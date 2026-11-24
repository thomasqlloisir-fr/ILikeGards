# Fiche produit — Médicaments & DM

Consultation sur smartphone, produit par produit, des données croisées de :
- **Stock / référentiel** médicaments (`MED_STOCK.xlsx`) et dispositifs médicaux (`DM_STOCK.xlsx`)
- **Mouvements de sortie** (`MED_Mvts_MVS.xlsx`)
- **Dotations de service** (`MED_DOTATIONS.xlsx`)

Le site (`docs/`) est publié par GitHub Pages. **Les données sont chiffrées** (AES-256-GCM,
clé dérivée du mot de passe par PBKDF2-SHA256, 250 000 itérations) : seul `docs/data.enc`
est publié, il est déchiffré dans le navigateur du téléphone. Les fichiers Excel ne sont
**jamais** envoyés sur GitHub (voir `.gitignore`).

## Mettre à jour les données (sur un PC)

```bash
pip install -r requirements.txt
# déposer les exports Excel dans le dossier donnees/ (DM_*.xlsx = dispositifs médicaux)
python build.py          # demande le mot de passe
git add docs/data.enc && git commit -m "Mise à jour données" && git push
```

Les fichiers sont reconnus par leurs colonnes (`Qté Stock (pr4)`, `Date Mouvement (mvs)`,
`Prévision Conso (cat)`), la clé commune est `No Produit (pr)`.

## Sur le téléphone

Ouvrir l'adresse GitHub Pages dans Chrome → menu ⋮ → **Ajouter à l'écran d'accueil**.

## Limites
- Le mot de passe protège les données, mais le dépôt et la page sont publics.
- Changer le mot de passe = relancer `build.py` avec le nouveau et republier.
