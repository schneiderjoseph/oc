# Kit d'exercices — Le Comptoir du Marché

**Localisation : Port-au-Prince, Haïti · Devise : $US (USD)**

> Montants en dollars US. Si votre OC est en **gourdes (HTG)**, convertissez au taux du jour (~132 G / 1 $US) en gardant les mêmes proportions.

**Taxe :** TCA 10 % (Taxe sur le Chiffre d'Affaires) — pas de TPS/TVQ canadiennes.

**Fournisseurs fictifs :**
- Distrib. Caraïbes (gros alimentaire)
- Boulangerie Pétion (pains)
- Emballages Haïti (papier)
- Boissons Nationale (soft drinks)

## Fichiers Word (dans ce dossier)

| Fichier | Style |
|---------|--------|
| `Exercice_Parcours_Lineaire.docx` | 40 étapes, tutoriels dans l'ordre |
| `Exercice_Pratique_Optimum_Control.docx` | Par thème + défis bonus |
| `Exercice_Corrige_Detaille.docx` | Corrigé — après l'exercice |

Copies identiques à la racine `E:\OC DOCS\` (générées en même temps).

## Fichiers annexes

- `ventes_csv/` — import Sales Mix  
- `Exercice_Corrige_Detaille.docx` — après l'exercice  
- `exercice_locale.py` — paramètres Haïti/$US (modifier ici pour regénérer)

## Régénérer

```bash
python build_parcours_lineaire.py
python build_exercice.py
python build_corrige.py
python generate_ventes_csv.py
```
