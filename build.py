#!/usr/bin/env python3
"""
Génère docs/data.enc (données chiffrées) à partir des fichiers Excel du dossier donnees/.

Fichiers reconnus automatiquement (d'après leurs colonnes) :
  - Référentiel / stock   : colonne "Qté Stock (pr4)"       -> type MED ou DM selon le nom (DM_*.xlsx = DM)
  - Mouvements de sortie  : colonne "Date Mouvement (mvs)"
  - Dotations de service  : colonne "Prévision Conso (cat)"

Usage :
  python build.py                  (le mot de passe est demandé)
  MEDAPP_PASSWORD=xxx python build.py
"""
import base64
import getpass
import gzip
import json
import os
import re
import sys
from datetime import datetime, date
from pathlib import Path

import pandas as pd
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "donnees"
OUT = ROOT / "docs" / "data.enc"
PBKDF2_ITER = 250_000
KEY = "No Produit (pr)"
LIB = "Libelle1 Produit (pr)"


def clean_header(h):
    """'Qté Stock (pr4)' -> 'Qté Stock'"""
    return re.sub(r"\s*\((pr\d*|cat|uf|mvs|mvt|fr|ma\d*)\)\s*$", "", str(h)).strip()


def read_report(path):
    """Lit un export : trouve la ligne d'en-tête contenant 'No Produit (pr)'."""
    raw = pd.read_excel(path, header=None, dtype=object)
    hdr = None
    for i in range(min(20, len(raw))):
        if KEY in [str(v).strip() for v in raw.iloc[i].tolist()]:
            hdr = i
            break
    if hdr is None:
        raise ValueError(f"{path.name} : colonne '{KEY}' introuvable")
    cols = [str(v).strip() if pd.notna(v) else None for v in raw.iloc[hdr].tolist()]
    df = raw.iloc[hdr + 1:].copy()
    df.columns = [c if c else f"_vide{i}" for i, c in enumerate(cols)]
    df = df[[c for c in df.columns if not c.startswith("_vide")]]
    df = df[df[KEY].notna()]
    return df


def val(v):
    """Convertit une cellule en valeur JSON compacte."""
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float):
        if v.is_integer():
            return int(v)
        return round(v, 4)
    if hasattr(v, "item"):  # numpy
        return val(v.item())
    if isinstance(v, str):
        s = v.strip()
        return s if s else None
    return v


def pid(v):
    v = val(v)
    if isinstance(v, float):
        v = int(v)
    return str(v)


def main():
    files = sorted(p for p in SRC.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        sys.exit(f"Aucun fichier .xlsx dans {SRC}")

    products = {}   # id -> {l, t, a, s:[[...]], m:[[date, qty, uf]], d:[[cat, uf, prev, emp]]}
    stock_cols = {}  # type -> [colonnes]
    ufs, cats = {}, {}
    sources = []
    mv_dates = []

    def prod(i, lib=None, atc=None):
        p = products.setdefault(i, {"l": None, "t": [], "a": None, "s": [], "m": [], "d": []})
        if lib and not p["l"]:
            p["l"] = lib
        if atc and not p["a"]:
            p["a"] = atc
        return p

    for f in files:
        df = read_report(f)
        cols = list(df.columns)
        kind = None
        if "Qté Stock (pr4)" in cols:
            kind = "DM" if f.name.upper().startswith("DM") else "MED"
            keep = [c for c in cols if c not in (KEY, LIB)]
            stock_cols[kind] = [clean_header(c) for c in keep]
            for _, r in df.iterrows():
                p = prod(pid(r[KEY]), val(r.get(LIB)), val(r.get("Code ATC (pr1)")))
                if kind not in p["t"]:
                    p["t"].append(kind)
                p["s"].append([kind] + [val(r[c]) for c in keep])
        elif "Date Mouvement (mvs)" in cols:
            kind = "MVTS"
            for _, r in df.iterrows():
                d = val(r["Date Mouvement (mvs)"])
                uf = pid(r.get("No UF (mvs)"))
                ufs.setdefault(uf, val(r.get("Libellé UF (uf)")))
                p = prod(pid(r[KEY]), val(r.get(LIB)), val(r.get("Code ATC (pr1)")))
                p["m"].append([d, val(r.get("Qté Mvt (mvs)")) or 0, uf])
                if d:
                    mv_dates.append(d)
        elif "Prévision Conso (cat)" in cols:
            kind = "DOT"
            for _, r in df.iterrows():
                uf = pid(r.get("No UF (uf)"))
                ufs.setdefault(uf, val(r.get("Libellé UF (uf)")))
                cat = val(r.get("No Cat (cat)"))
                cats.setdefault(cat, val(r.get("Libellé catalogue (cat)")))
                p = prod(pid(r[KEY]), val(r.get(LIB)), val(r.get("Code ATC (pr1)")))
                p["d"].append([cat, uf, val(r.get("Prévision Conso (cat)")),
                               val(r.get("Autre Emplacement (cat)")), val(r.get("Code Ucd (pr1)"))])
        else:
            print(f"  ! {f.name} : format non reconnu, ignoré")
            continue
        sources.append({"fichier": f.name, "type": kind, "lignes": len(df)})
        print(f"  {f.name:30s} {kind:5s} {len(df):6d} lignes")

    for p in products.values():
        p["m"].sort(key=lambda x: x[0] or "", reverse=True)

    data = {
        "meta": {
            "genere": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "sources": sources,
            "mvts_du": min(mv_dates) if mv_dates else None,
            "mvts_au": max(mv_dates) if mv_dates else None,
            "stock_cols": stock_cols,
        },
        "uf": ufs,
        "cat": cats,
        "p": products,
    }

    raw = json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    packed = gzip.compress(raw, 9)

    pwd = os.environ.get("MEDAPP_PASSWORD") or getpass.getpass("Mot de passe de chiffrement : ")
    if len(pwd) < 8:
        sys.exit("Mot de passe trop court (8 caractères minimum).")
    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(hashes.SHA256(), 32, salt, PBKDF2_ITER).derive(pwd.encode("utf-8"))
    ct = AESGCM(key).encrypt(iv, packed, None)

    OUT.write_text(json.dumps({
        "v": 1, "iter": PBKDF2_ITER,
        "salt": base64.b64encode(salt).decode(),
        "iv": base64.b64encode(iv).decode(),
        "data": base64.b64encode(ct).decode(),
    }))
    print(f"\n{len(products)} produits — JSON {len(raw)/1e6:.1f} Mo, compressé+chiffré {OUT.stat().st_size/1e6:.2f} Mo")
    print(f"-> {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
