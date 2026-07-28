# Import des ventes CSV — Exercice Le Comptoir du Marché

**Contexte :** Port-au-Prince, Haïti · prix menu en **$US** (USD).

Fichiers dans ce dossier : `ventes_csv/`

## Calendrier réel (juin–juillet 2026)

| Jour | Date | Événement / ventes |
|------|------|-------------------|
| **Lundi** | **29/06/2026** | Inventaire d'ouverture (Finalize) |
| **Mardi** | **30/06/2026** | Factures · **Till Tape** (1re vente) |
| **Mercredi** | **01/07/2026** | Import CSV · Waste laitue · **1 316,10 $** Gross |
| **Jeudi** | **02/07/2026** | Import CSV · Facture DC-0320 |
| **Vendredi** | **03/07/2026** | Import CSV · Facture BN |
| **Samedi** | **04/07/2026** | Import CSV · Facture EH · **clôture inventaire** |

## Fichiers CSV (noms = jour calendaire)

| Fichier | Date |
|---------|------|
| `ventes_mardi_2026-06-30.csv` | 30/06 (Till Tape ou import) |
| `ventes_mercredi_2026-07-01.csv` | 01/07 |
| `ventes_jeudi_2026-07-02.csv` | 02/07 |
| `ventes_vendredi_2026-07-03.csv` | 03/07 |
| `ventes_samedi_2026-07-04.csv` | 04/07 |
| `ventes_semaine_comptoir.csv` | Semaine complète |

Colonne 5 = **Gross Sales** (total ligne = Qté × prix unitaire). Spec Comptoir CSV : indices 1→5, **Selling Price** vide.

## Erreur « Column index… »

Fichier sans colonne **Date** ou seulement 2 colonnes (`ventes_minimal_*`) → incompatible avec Comptoir CSV.

## Regénérer les CSV

```bash
python generate_ventes_csv.py
```

Depuis `E:\OC DOCS`.
