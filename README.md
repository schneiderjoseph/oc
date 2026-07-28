# Optimum Control — Formation, audit & outils

Dépôt personnel pour la formation **Optimum Control v5**, l'exercice **Le Comptoir du Marché**, la mission **La Réserve**, et les scripts d'audit SQL.

## Structure

```
├── README.md
├── ROADMAP_V2.md          # Prochaines améliorations audit DB
├── requirements.txt
├── scripts/
│   ├── audit/             # Audit SQL → Excel (La Réserve)
│   ├── build/             # Génération docs Word exercice
│   ├── verify/            # Vérifications base ocdata (exercice)
│   └── tools/             # Utilitaires doc / mapping
├── docs/
│   ├── formation/         # Playlist, mapping vidéos ↔ livre
│   ├── la-reserve/        # Mails, mission, contexte client
│   └── schema/            # Schéma tables oc.* (référence)
├── transcriptions/
│   ├── clean/             # Transcriptions nettoyées (tutoriels OC)
│   └── raw/               # Transcriptions brutes
├── exercice_comptoir/     # Données exercice (ventes CSV, README)
└── livres/                # Guides Word (.docx à déposer ici)
```

## Démarrage rapide — Audit base OC

```powershell
pip install -r requirements.txt

# Base locale exercice (LocalDB)
python scripts/audit/oc_db_full_audit.py `
  --server "(localdb)\mssqllocaldb" `
  --database ocdata `
  --trusted `
  -o output/OC_Audit_Local.xlsx

# Chez un client (adapter serveur + base)
python scripts/audit/oc_db_full_audit.py `
  --server "NOM_PC\SQLEXPRESS" `
  --database "NomBaseOC" `
  --trusted `
  -o output/LaReserve_Audit.xlsx
```

**Prérequis :** [ODBC Driver 17 ou 18 for SQL Server](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server)

## Contenu principal

| Zone | Contenu |
|------|---------|
| **Audit** | `oc_db_full_audit.py` — 21 onglets Excel depuis la DB |
| **La Réserve** | `docs/la-reserve/` — mail mission, contexte audit |
| **Exercice** | `exercice_comptoir/` + scripts `build/` |
| **Formation** | Transcriptions 60+ vidéos OC + playlist ordonnée |
| **Schéma DB** | `docs/schema/oc_schema_tables.txt` — tables `oc.*` |

## V2

Voir [ROADMAP_V2.md](ROADMAP_V2.md) — doublons flous, orphelins recettes, plan d'action auto, config YAML client, mode avant/après.

## Notes

- Scripts en **lecture seule** sur la DB client (SELECT uniquement).
- Ne pas committer : exports Excel client, `bar.csv`, contrats signés (voir `.gitignore`).
- Les `.docx` des livres vont dans `livres/` (à copier depuis la machine principale).

## Licence

Usage personnel / mission client. Contenu TracRite / Optimum Control : droits TracRite Systems.
