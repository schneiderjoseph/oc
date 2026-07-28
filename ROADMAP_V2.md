# Roadmap — Audit OC V2

## Déjà en V1 (`scripts/audit/oc_db_full_audit.py`)

- Connexion SQL Server directe (lecture seule)
- Export Excel multi-onglets (21 feuilles)
- Items, preps, products, fournisseurs, case sizes, recettes
- Doublons de noms exacts
- Anomalies P1/P2/P3 (règles La Réserve)
- Factures, ventes POS, inventaires, volumes tables

## Prévu V2 — détection plus intelligente

### Qualité des fiches
- Items **sans fournisseur** / sans case size / sans UOM recette
- Items **sans emplacement** ou multi-emplacements incohérents
- **Orphelins** : ingrédients de recette pointant vers un item supprimé/inactif
- Products **sans POS ID** ou sans ingrédients
- Preps **sans yield** ou batch UOM manquant
- Fiches **inactives** encore utilisées dans des recettes

### Doublons & amalgames
- Similarité **floue** (Levenshtein / token sort) — pas seulement nom identique
- Suggestions d'amalgame : *garder ID X, fusionner Y, Z*
- Détection **même order code** chez un fournisseur sur 2 items
- Groupes par **fournisseur + famille** (bar, viande, etc.)

### Bar / alcools / vins
- Règles étendues : liqueurs, bières, vins au verre
- Liste **prioritaire** spiritueux top CA (si ventes dispo)
- Alertes **bottle sans ml/oz** conversion

### Achats & ventes
- Factures **déséquilibrées** (total ≠ lignes)
- Lignes facture à **coût zéro**
- Statuts POS **décodés** (Valid, Unlinked, Mismatched, Pending)
- Compteur **Unlinked par jour** + tendance

### Rapport & usage
- Onglet **Plan d'action** (semaine 1–4) auto-généré
- Onglet **Amalgames proposés** (à valider client)
- Fichier **config YAML** par client (seuils, exclusions, emplacements)
- Mode **comparaison** : 2 exports datés (avant/après corrections)
- Métadonnées de run (date, serveur, durée, version script)

### Technique
- Support **plusieurs magasins** (`--store-id`)
- Export **CSV** en parallèle de l'Excel
- Meilleure gestion **encodage console Windows**
- Tests automatiques sur base exercice `ocdata`
