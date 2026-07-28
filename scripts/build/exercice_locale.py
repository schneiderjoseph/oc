# Contexte local — Haïti, devise USD
# Importé par build_exercice.py et build_parcours_lineaire.py

import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent
KIT_DIR = BASE / "exercice_comptoir"


def publish_docx(doc, out: Path) -> None:
    """Enregistre le docx à la racine OC DOCS et copie dans exercice_comptoir/."""
    import tempfile
    KIT_DIR.mkdir(parents=True, exist_ok=True)
    kit = KIT_DIR / out.name
    tmp = Path(tempfile.gettempdir()) / f"oc_publish_{out.name}"
    doc.save(tmp)
    for target in (out, kit):
        try:
            shutil.copy2(tmp, target)
            print(f"OK -> {target}")
        except PermissionError:
            print(f"WARN: {target.name} verrouillé — fermez Word et relancez le script")
    try:
        tmp.unlink(missing_ok=True)
    except OSError:
        pass

PAYS = "Haïti"
VILLE = "Port-au-Prince"
ADRESSE = "45 rue Lamarre, Port-au-Prince, Haïti"
DEVISE = "USD ($US)"
DEVISE_COURTE = "$US"
SEMAINE_RAPPORT = "Dimanche → Samedi"

# Calendrier semaine type — Le Comptoir du Marché (juin–juillet 2026, jours réels)
EXO_DATE_OPENING = "29/06/2026"   # lundi soir — Count Inventory
EXO_DATE_CLOSING = "04/07/2026"   # samedi matin — 2e inventaire
EXO_FIRST_SALES_DAY = "Mardi"       # 1re journée ventes / Till Tape
EXO_WEEK_NOTE = (
    "Calendrier **aligné sur juin–juillet 2026** : inventaire d'ouverture **lundi 29/06/2026** (soir), "
    "exploitation **mardi 30/06 → samedi 04/07/2026**, clôture **samedi 04/07/2026** matin. "
    "**Invoice Date** = date de livraison sur la facture fournisseur (jour d'achat), "
    "pas la date d'inventaire d'ouverture."
)
EXO_WEEK_HEADERS = ["Jour", "Date (saisir dans OC)", "Événement clé"]
EXO_WEEK = [
    ["Lundi", "29/06/2026", "Inventaire d'ouverture — Finalize (soir)"],
    ["Mardi", "30/06/2026", "Factures BP-2026-0142 + DC-2026-0318 · Till Tape · Daily Sales"],
    ["Mercredi", "01/07/2026", "Import ventes · Waste laitue"],
    ["Jeudi", "02/07/2026", "Facture DC-2026-0320 · ventes · Waste frites"],
    ["Vendredi", "03/07/2026", "Facture BN-2026-0088 · ventes · Waste chili"],
    ["Samedi", "04/07/2026", "Facture EH-2026-0205 · ventes · Waste pain · inventaire clôture"],
]
EXO_DAY_TO_DATE = {day: date for day, date, _ in EXO_WEEK}
EXO_DAY_LOWER = {
    "Mardi": "mardi", "Mercredi": "mercredi", "Jeudi": "jeudi",
    "Vendredi": "vendredi", "Samedi": "samedi",
}

def exo_date_to_iso(ddmmyyyy: str) -> str:
    d, m, y = ddmmyyyy.split("/")
    return f"{y}-{m}-{d}"

def exo_sales_iso_by_day() -> dict:
    """Dates ISO (YYYY-MM-DD) pour import CSV ventes — lun à ven."""
    return {
        EXO_DAY_LOWER[day]: exo_date_to_iso(date)
        for day, date, _ in EXO_WEEK
        if day in EXO_DAY_LOWER
    }

def exo_week_rows():
    return [list(row) for row in EXO_WEEK]

# Taxes — TCA (Haïti)
# code, description, type, taux, recoverable, account, apply_purchase_on_item
TAXES = [
    ("TCA", "Taxe sur le Chiffre d'Affaires", "Percentage", "10 %", "Oui", "TCA recoverable", "Non"),
]

TAX_HEADERS = [
    "Code", "Description", "Type", "Taux", "Recoverable", "Account", "Apply purchase amount on item",
]

def tax_table_rows():
    return [list(row) for row in TAXES]

TAXES_DOC = "TCA 10 % recoverable (Haïti). Pas de TPS/TVQ canadiennes."

TAXES_NOTE = (
    "Setup → Taxes and Adjustments → Add. Renseignez Code, Description, Recoverable, "
    "Percentage (ou Value pour un dépôt), Account (compte OC lié à cette taxe) et "
    "Apply purchase amount on item : si coché, le montant de taxe est imputé au compte "
    "de l'item (Inventory Group) plutôt qu'au compte Account — en général Non pour une TCA recoverable. "
    "Note : le champ Account des taxes liste les comptes CostOfSales, Expense et Liability "
    "(pas Asset ni Income). Pour une taxe, préférez Liability (manuel TracRite : comptes de taxes). "
    "Recoverable = Oui distingue une taxe récupérable."
)

# Tax Groups — lie un ou plusieurs taxes/ajustements aux items (Case Size → Tax Group)
TAX_GROUP_CODE = "TCA-ACHAT"

TAX_GROUPS_NOTE = (
    "Setup → Tax Groups → Add (après Taxes and Adjustments, étape 10). "
    "Un Tax Group regroupe une ou plusieurs taxes/ajustements et se assigne sur chaque Case Size (champ Tax Group). "
    "C'est ce lien qui associe l'item aux taxes applicables — distinct de la saisie TCA sur l'onglet Taxes d'une facture. "
    "Price Includes Tax : cocher seulement si le prix fournisseur est TTC (bière, alcool, dépôts inclus)."
)

TAX_GROUP_HEADERS = ["Code groupe", "Description", "Taxes / ajustements cochés", "Usage dans l'exercice"]

TAX_GROUP_DOC = [
    (
        TAX_GROUP_CODE,
        "Achats soumis à TCA",
        "TCA 10 %",
        "Tous les items alimentaires/boissons/papier ; Price Includes Tax = Non",
    ),
]

def tax_group_rows():
    return [list(row) for row in TAX_GROUP_DOC]

# Option gourde — note pour l'utilisateur
NOTE_DEVISE = (
    "Montants de l'exercice en $US. Si votre base OC est en gourdes (HTG), "
    "multipliez par le taux du jour (ex. ~132 G pour 1 $US) ou saisissez directement en HTG "
    "en gardant les mêmes proportions."
)

# Fournisseurs — champs Setup → Suppliers (Haïti)
# nom, accounting_vendor_id, address, city, province, postal_code, country, contact, phone, email
FOURNISSEURS = [
    (
        "Distrib. Caraïbes", "576294",
        "12 zone industrielle Jalousie", "Port-au-Prince", "Ouest", "HT6120", "Haïti",
        "Jean-Claude Mercier", "+509 2810-4520", "commandes@distrib-caribes.ht",
    ),
    (
        "Boulangerie Pétion", "576301",
        "8 rue Pétion, Pétion-Ville", "Pétion-Ville", "Ouest", "HT6140", "Haïti",
        "Marie Dupont", "+509 3712-8891", "ventes@boulangerie-petion.ht",
    ),
    (
        "Emballages Haïti", "576318",
        "45 av. John Brown, Delmas 33", "Delmas", "Ouest", "HT6130", "Haïti",
        "Richard Louis", "+509 2945-6670", "info@emballages-haiti.ht",
    ),
    (
        "Boissons Nationale", "576325",
        "Route Nationale #1, Croix-des-Bouquets", "Croix-des-Bouquets", "Ouest", "HT6310", "Haïti",
        "Sophie Célestin", "+509 2234-1100", "distribution@boissons-nationale.ht",
    ),
]

SUPPLIER_HEADERS = [
    "Fournisseur", "Accounting Vendor ID", "Address", "City", "Province",
    "Postal Code", "Country", "Contact", "Téléphone", "Courriel",
]

def supplier_table_rows():
    return [list(row) for row in FOURNISSEURS]

SUPPLIERS_NOTE = (
    "Setup → Suppliers. Nom obligatoire. Renseignez aussi Address, City, Province, "
    "Postal Code, Country et les contacts si disponibles. "
    "Accounting Vendor ID = ID numérique du vendeur dans votre logiciel comptable "
    "(ex. 576294 — doit correspondre exactement, vidéo #49)."
)

# Alias pour textes dans les exercices
NOM_DISTRIB = "Distrib. Caraïbes"
NOM_BOULANGERIE = "Boulangerie Pétion"
NOM_EMBALLAGES = "Emballages Haïti"
NOM_BOISSONS = "Boissons Nationale"

# Case Size — liste « Prix pour » (Purchase Unit) après le montant
CASE_SIZE_PURCHASE_UNIT_NOTE = (
    "Case Size — après Tax Group : (1) saisir Purchase Price ; "
    "(2) liste déroulante à droite du prix = pour quoi ce montant s'applique — "
    "case, bag, each, gallon, dozen, lb, kg, L… (Purchase Unit · oc.CaseSize.PurchaseUom) ; "
    "(3) Split unit + quantité = contenu de cette unité d'achat (sacs, kg, lb, bidons…) ; "
    "(4) Pack unit + quantité = unité recette. "
    "Exercice Comptoir : colonne « Prix pour » = presque toujours case ; Split décrit ce qu'il y a dedans "
    "(12 each, 5 kg, 25 lb…). Rare : catch weight → Prix pour = lb ou kg au lieu de case."
)

# 3 premiers items — saisie complète (étape 17 / Partie 2.0)
ITEMS_FIRST3_NOTE = (
    "Créez ces 3 items en premier (vidéo #06). Core à gauche, Case Size à droite, puis Save. "
    f"Tax Group = {TAX_GROUP_CODE} sur chaque Case Size. "
    + CASE_SIZE_PURCHASE_UNIT_NOTE
    + " Pain : en-tête OC « 12 ea / cs and 1 ea / ea » = prix pour 1 case, 12 pains, 1 each / each — "
    "le libellé « how many each in a each » avec Pack = 1 est normal quand How used = Unit."
)

ITEMS_FIRST3_CORE_HEADERS = [
    "Item", "Fournisseur", "Inventory Group", "Main Location", "How used?",
    "Reporting UOM", "Ingredient UOM", "Key Item", "Actualize",
]

ITEMS_FIRST3_CORE = [
    [
        "Pain burger", NOM_BOULANGERIE, "Dry Goods", "Sec", "Unit",
        "each (chacun)", "each", "Oui", "Non",
    ],
    [
        "Bœuf haché 80/20", NOM_DISTRIB, "Proteins", "Congélateur", "Weight",
        "gram", "gram", "Oui", "Non",
    ],
    [
        "Pommes de terre", NOM_DISTRIB, "Produce", "Sec", "Weight",
        "gram", "gram", "Oui", "Non",
    ],
]

ITEMS_FIRST3_CASE_HEADERS = [
    "Item", "Purchase Price", "Prix pour", "Split unit", "Qty in split",
    "Pack unit", "Qty per split (pack)", "Yield %", "Tax Group",
    "Price incl. tax", "En-tête OC attendu", "Coût unitaire OC",
]

ITEMS_FIRST3_CASE = [
    [
        "Pain burger", "2,40 $US", "case", "each", "12",
        "each", "1", "100 %", TAX_GROUP_CODE,
        "Non", "12 ea / cs · 1 ea / ea", "0,20 $US / chacun",
    ],
    [
        "Bœuf haché 80/20", "42,00 $US", "case", "kg", "5",
        "gram", "1000", "100 %", TAX_GROUP_CODE,
        "Non", "5 kg / cs · 1000 g / kg", "0,0084 $US / g",
    ],
    [
        "Pommes de terre", "22,00 $US", "case", "lb", "50",
        "gram", "453,592", "85 %", TAX_GROUP_CODE,
        "Non", "50 lb / cs · 453,592 g / lb", "≈ 0,0011 $US / g utilisable",
    ],
]

def items_first3_core_rows():
    return [list(row) for row in ITEMS_FIRST3_CORE]

def items_first3_case_rows():
    return [list(row) for row in ITEMS_FIRST3_CASE]

# En-têtes partagés — tableaux détaillés items (Core + Case Size)
ITEMS_DETAIL_CORE_HEADERS = ITEMS_FIRST3_CORE_HEADERS
ITEMS_DETAIL_CASE_HEADERS = ITEMS_FIRST3_CASE_HEADERS

ITEM_MULTI_CASESIZE_NOTE = (
    "Un seul item = une seule fiche Core Information (Description, groupe, How used, UOM…). "
    "Plusieurs fournisseurs ou formats d'achat = plusieurs Case Sizes sur le MÊME item (bouton Add / Remove). "
    "En base OC : oc.Item (1) ← oc.CaseSize (N) — chaque ligne lie Item + Supplier + prix + unités. "
    "Le stock (qty on hand) est la somme de toutes les case sizes actives. "
    "Ne jamais recréer un 2e item « Bacon » pour un autre fournisseur : Add case size."
)

# Exercice labo — 2e case size sur Bacon tranché (étape 18 / Partie 2.1)
ITEM_MULTI_CASESIZE_EXAMPLE_NOTE = (
    "Après avoir créé « Bacon tranché » avec la 1re ligne du tableau Case Size (Distrib. Caraïbes) : "
    "Inventory → Items → ouvrir Bacon tranché → barre d'outils Add (ne pas New). "
    "Saisir la 2e case size ci-dessous — Supplier = Boulangerie Pétion, Core inchangé. "
    "Save → View All : 2 lignes pour le même item. "
    "Dans la vraie vie, un 2e fournisseur doit exister dans Setup → Suppliers (déjà le cas ici). "
    "Default / Current Case Size : garder Distrib. Caraïbes comme format principal ; la boulangerie = format secours petit format."
)

ITEM_MULTI_CASESIZE_EXAMPLE_HEADERS = [
    "Case #", "Item", "Supplier", "Purchase Price", "Prix pour", "Split unit", "Qty in split",
    "Pack unit", "Qty per split (pack)", "Yield %", "Tax Group", "Price incl. tax",
    "En-tête OC attendu", "Coût unitaire OC",
]

ITEM_MULTI_CASESIZE_EXAMPLE = [
    [
        "1 — déjà dans le tableau principal", "Bacon tranché", NOM_DISTRIB,
        "18,50 $US", "case", "each", "2", "gram", "500", "95 %", TAX_GROUP_CODE,
        "Non", "2 sacs × 500 g / cs", "≈ 0,019 $US / g",
    ],
    [
        "2 — à ajouter (Add)", "Bacon tranché", NOM_BOULANGERIE,
        "42,00 $US", "case", "each", "12", "gram", "200", "95 %", TAX_GROUP_CODE,
        "Non", "12 sacs × 200 g / cs", "≈ 0,018 $US / g",
    ],
]

def items_multicasize_example_rows():
    return [list(row) for row in ITEM_MULTI_CASESIZE_EXAMPLE]

def item_multicasize_book_blocks():
    """Paragraphes Word pour Comprendre_Optimum_Control (rebuild_final.py)."""
    blocks = [
        ("Heading2", "Exemple labo — Bacon tranché, 2e fournisseur (Add case size)"),
        ("Normal", ITEM_MULTI_CASESIZE_NOTE),
        ("Normal", ITEM_MULTI_CASESIZE_EXAMPLE_NOTE),
    ]
    for row in ITEM_MULTI_CASESIZE_EXAMPLE:
        case_num, _item, supplier, price, purch_u, split_u, split_q, pack_u, pack_q = row[:9]
        yield_pct, tax_grp, incl_tax, header, cost = row[9:]
        blocks.append(("Heading3", case_num))
        blocks.append((
            "Normal",
            f"• Supplier — {supplier} ; Purchase Price — {price} pour 1 {purch_u}. "
            f"Split — {split_q} {split_u} ; Pack — {pack_q} {pack_u} par split unit.",
        ))
        blocks.append((
            "Normal",
            f"• Yield % — {yield_pct} ; Tax Group — {tax_grp} ; Price Includes Tax — {incl_tax}.",
        ))
        blocks.append((
            "Normal",
            f"• En-tête OC attendu : {header} ; Coût unitaire {cost}.",
        ))
    blocks.append((
        "Normal",
        "Vérification : View All affiche 2 lignes pour un seul item « Bacon tranché ». "
        "Default Ordering Case Size = Distrib. Caraïbes (format principal). "
        "En base : oc.Item (1) ← oc.CaseSize (2) avec Supplier différent sur chaque ligne.",
    ))
    return blocks

# Actualize Usage Values — huile friture (exercice Comptoir)
ACTUALIZE_HUILE_NOTE = (
    "Pourquoi Actualize Usage Values = Oui sur Huile friture (seul item de l'exercice avec cette case cochée) : "
    "l'huile du bain de friture s'absorbe dans les aliments, s'évapore et se complète en cours de semaine — "
    "l'usage réel ne colle pas à l'idéal calculé par les ventes et recettes (même si 200 ml figure dans le prep Frites). "
    "Sans Actualize, Usage Summary montrerait un écart permanent difficile à expliquer (faux « gaspillage » chaque semaine). "
    "Avec Actualize : le logiciel considère l'usage réel comme l'usage théorique — l'écart disparaît "
    "(astérisque * à côté du montant idéal dans les rapports). "
    "Le stock (inventaire, qty on hand) reste compté normalement ; seul le rapport d'écart idéal vs réel est neutralisé pour cet item. "
    "À l'inverse : bœuf, bacon, laitue → Actualize = Non pour voir les vraies variances (portion, perte, vol)."
)

# 13 items alimentaires restants (étape 18) — après pain, bœuf, pommes
ITEMS_REST13_NOTE = (
    "Même méthode que l'étape 17 : Core à gauche, Case Size à droite, Save. "
    f"Tax Group = {TAX_GROUP_CODE}. Sel : Key Item = Non (option comptage allégé). "
    "Huile friture : Actualize = Oui — voir encadré ci-dessous. "
    "Ketchup, mayo, bacon : ajouter Ligne chaude en Locations (secondaire) après Save."
)

ITEMS_REST13_CORE = [
    ["Bacon tranché", NOM_DISTRIB, "Proteins", "Congélateur", "Weight", "gram", "gram", "Oui", "Non"],
    ["Laitue romaine", NOM_DISTRIB, "Produce", "Réfrigérateur", "Weight", "gram", "gram", "Oui", "Non"],
    ["Parmesan râpé", NOM_DISTRIB, "Dairy", "Réfrigérateur", "Weight", "gram", "gram", "Oui", "Non"],
    ["Croûtons", NOM_DISTRIB, "Dry Goods", "Sec", "Weight", "gram", "gram", "Oui", "Non"],
    ["César dressing", NOM_DISTRIB, "Dairy", "Réfrigérateur", "Volume", "ml", "ml", "Oui", "Non"],
    ["Huile friture", NOM_DISTRIB, "Dry Goods", "Friterie", "Volume", "ml", "ml", "Oui", "Oui"],
    ["Sel fin", NOM_DISTRIB, "Spices", "Sec", "Weight", "gram", "gram", "Non", "Non"],
    ["Tomates", NOM_DISTRIB, "Produce", "Réfrigérateur", "Weight", "gram", "gram", "Oui", "Non"],
    ["Oignons", NOM_DISTRIB, "Produce", "Réfrigérateur", "Weight", "gram", "gram", "Oui", "Non"],
    ["Haricots rouges secs", NOM_DISTRIB, "Dry Goods", "Sec", "Weight", "gram", "gram", "Oui", "Non"],
    ["Bœuf haché chili", NOM_DISTRIB, "Proteins", "Congélateur", "Weight", "gram", "gram", "Oui", "Non"],
    ["Ketchup", NOM_DISTRIB, "Dry Goods", "Sec", "Volume", "ml", "ml", "Oui", "Non"],
    ["Mayonnaise", NOM_DISTRIB, "Dairy", "Réfrigérateur", "Volume", "ml", "ml", "Oui", "Non"],
]

ITEMS_REST13_CASE = [
    ["Bacon tranché", "18,50 $US", "case", "each", "2", "gram", "500", "95 %", TAX_GROUP_CODE,
     "Non", "2 sacs × 500 g / cs", "≈ 0,019 $US / g · conv. tranches étape 20"],
    ["Laitue romaine", "28,00 $US", "case", "each", "12", "gram", "500", "75 %", TAX_GROUP_CODE,
     "Non", "12 têtes / cs · 500 g / tête", "≈ 0,0062 $US / g utilisable"],
    ["Parmesan râpé", "24,00 $US", "case", "kg", "2", "gram", "1000", "100 %", TAX_GROUP_CODE,
     "Non", "2 kg / cs · 1000 g / kg", "0,012 $US / g"],
    ["Croûtons", "8,00 $US", "case", "kg", "1", "gram", "1000", "100 %", TAX_GROUP_CODE,
     "Non", "1 kg / cs", "0,008 $US / g"],
    ["César dressing", "52,00 $US", "case", "each", "4", "ml", "3780", "100 %", TAX_GROUP_CODE,
     "Non", "4 × 3,78 L / cs · 3780 ml / bidon", "≈ 0,0034 $US / ml"],
    ["Huile friture", "38,00 $US", "case", "L", "16", "ml", "1000", "100 %", TAX_GROUP_CODE,
     "Non", "16 L / cs · 1000 ml / L", "0,0024 $US / ml"],
    ["Sel fin", "3,50 $US", "case", "kg", "2", "gram", "1000", "100 %", TAX_GROUP_CODE,
     "Non", "2 kg / cs", "0,0018 $US / g"],
    ["Tomates", "45,00 $US", "case", "lb", "25", "gram", "453,592", "90 %", TAX_GROUP_CODE,
     "Non", "25 lb / cs · 453,592 g / lb", "≈ 0,0044 $US / g utilisable"],
    ["Oignons", "18,00 $US", "case", "lb", "25", "gram", "453,592", "88 %", TAX_GROUP_CODE,
     "Non", "25 lb / cs · 453,592 g / lb", "≈ 0,0018 $US / g utilisable"],
    ["Haricots rouges secs", "55,00 $US", "case", "kg", "20", "gram", "1000", "100 %", TAX_GROUP_CODE,
     "Non", "20 kg / cs", "0,0028 $US / g"],
    ["Bœuf haché chili", "40,00 $US", "case", "kg", "5", "gram", "1000", "100 %", TAX_GROUP_CODE,
     "Non", "5 kg / cs", "0,008 $US / g"],
    ["Ketchup", "38,00 $US", "case", "each", "6", "ml", "3371", "100 %", TAX_GROUP_CODE,
     "Non", "6 × 114 oz / cs · ~3371 ml / bouteille", "≈ 0,0019 $US / ml"],
    ["Mayonnaise", "48,00 $US", "case", "each", "4", "ml", "3780", "100 %", TAX_GROUP_CODE,
     "Non", "4 × 3,78 L / cs", "≈ 0,0032 $US / ml"],
]

def items_rest13_core_rows():
    return [list(row) for row in ITEMS_REST13_CORE]

def items_rest13_case_rows():
    return [list(row) for row in ITEMS_REST13_CASE]

# 5 items boissons + papier (étape 19)
ITEMS_LAST5_NOTE = (
    "Même structure Core + Case Size. Serviette : Key Item = Non si comptage allégé."
)

ITEMS_LAST5_CORE = [
    ["Cola 355 ml", NOM_BOISSONS, "Soft Drinks", "Bar", "Unit", "each", "each", "Oui", "Non"],
    ["Eau 500 ml", NOM_BOISSONS, "Soft Drinks", "Bar", "Unit", "each", "each", "Oui", "Non"],
    ["Boîte burger", NOM_EMBALLAGES, "Paper", "Sec", "Unit", "each", "each", "Oui", "Non"],
    ["Serviette", NOM_EMBALLAGES, "Paper", "Sec", "Unit", "each", "each", "Non", "Non"],
    ["Fourchette plastique", NOM_EMBALLAGES, "Paper", "Sec", "Unit", "each", "each", "Oui", "Non"],
]

ITEMS_LAST5_CASE = [
    ["Cola 355 ml", "12,00 $US", "case", "each", "24", "each", "1", "100 %", TAX_GROUP_CODE,
     "Non", "24 canettes / cs · 1 ea / ea", "0,50 $US / canette"],
    ["Eau 500 ml", "8,00 $US", "case", "each", "24", "each", "1", "100 %", TAX_GROUP_CODE,
     "Non", "24 bouteilles / cs", "≈ 0,33 $US / bouteille"],
    ["Boîte burger", "45,00 $US", "case", "each", "500", "each", "1", "100 %", TAX_GROUP_CODE,
     "Non", "500 / cs", "0,09 $US / boîte"],
    ["Serviette", "35,00 $US", "case", "each", "5000", "each", "1", "100 %", TAX_GROUP_CODE,
     "Non", "5000 / cs", "0,007 $US / serviette"],
    ["Fourchette plastique", "12,00 $US", "case", "each", "1000", "each", "1", "100 %", TAX_GROUP_CODE,
     "Non", "1000 / cs", "0,012 $US / fourchette"],
]

def items_last5_core_rows():
    return [list(row) for row in ITEMS_LAST5_CORE]

def items_last5_case_rows():
    return [list(row) for row in ITEMS_LAST5_CASE]

# ---- Preps (Partie 3 / étape 22) ----
PREPS_NOTE = (
    "Recipe → Preps → New. Ordre de création : d'abord les preps qui n'utilisent que des Items, "
    "puis ceux qui utilisent un autre Prep (bouton Add Prep en haut du tableau Ingrédients). "
    "Créer « Frites maison — batch » AVANT « Portion frites 200 g » (prep dans prep). "
    "Chaque prep doit lister tous les ingrédients réels : ex. boulette = bœuf + sel (pas le bœuf seul)."
)

PREPS_NESTED_NOTE = (
    "Prep dans prep (obligatoire exercice) : Portion frites 200 g consomme le prep Frites maison — batch. "
    "Ingrédients → Add Prep → Frites maison — batch, 200 gram. "
    "Le product Frites moyennes (étape 24) utilise Portion frites, pas le batch entier."
)

PREPS_YIELD_NOTES = (
    "Rendements (Batch Yield) — ne pas confondre avec la somme des lignes ingrédients : "
    "Frites batch : 5050 g entrés → 4000 g sortis (eau + parures friture). "
    "Sauce burger : 1000 ml entrés → 950 ml (léger écart au mélange). "
    "Chili batch : ~7040 g d'ingrédients secs → 12 000 g sortis : les haricots secs absorbent l'eau "
    "à la cuisson (l'eau n'est pas un item — rendement pesé en sortie). "
    "Boulette / portion frites : 1 each = une unité portionnée."
)

PREPS_BATCH_UNIT_NOTE = (
    "Erreur OC « It can't be 1 gram on a gram » (ou ml/ml) : ne mettez PAS Batch Yield = 4000 gram "
    "ET Recipe Unit = gram avec « per 1 gram » — OC refuse un ratio 1:1 sur la même unité. "
    "Méthode exercice (recommandée) : Batch Uom = Batch · Batch Yield = 1 batch · "
    "Recipe Unit = gram (ou ml) · Quantité = poids/volume total produit (4000 g, 950 ml, 12 000 g). "
    "Alternative : Batch Yield = 4 kg · Recipe Unit = gram · Quantité = 4000."
)

PREPS_ACTUALIZE_NOTE = (
    "Actualize Usage Values : en v5, cette case n'apparaît PAS sur l'écran New Prep (vidéo #16). "
    "Elle se règle uniquement sur Inventory → Items → New / fiche Item (vidéo #06). "
    "Exercice : Actualize = Oui seulement sur l'item Huile friture — pas sur les preps. "
    "L'huile du prep Frites est suivie via l'item Huile friture (200 ml dans la recette)."
)

# Guide écran New Prep (labels OC v5 — capture exercice)
PREPS_SCREEN_INTRO = (
    "Un Prep = ce que vous transformez en cuisine (sauce, batch de frites, boulette) avant de l'utiliser "
    "dans un Product. La fenêtre se divise en Core Information (gauche), Inventory + Prep Sheet (gauche bas), "
    "Ingrédients (droite). Onglets en haut : Recipe Instruction, Image, Location, Conversions, Ingredient Nutrition."
)

PREPS_FIELD_GUIDE_HEADERS = ["Section / champ (OC)", "Rôle", "Exercice Le Comptoir"]

PREPS_FIELD_GUIDE = [
    ["Description", "Nom de la préparation", "Voir tableau « Valeurs par prep »"],
    ["Inventory Group", "Groupe d'inventaire du prep", "Prep (groupe par défaut OC) — ou Preps si vous l'avez créé"],
    ["How is it used?", "Poids, Volume ou Unité — verrouillé après Save", "Weight sauf Sauce burger = Volume ; Boulette et Portion frites = Unit"],
    ["Batch Uom", "Contenant de production (batch, bac, bouteille…)", "Batch pour tous les preps poids/volume"],
    ["Batch Yield", "Quantité produite par lot (ou 1 batch — voir note)", "1 batch + qty recette, ou 4 kg / 0,95 L"],
    ["Recipe Unit (per ___)", "Unité pour portionner dans d'autres recettes", "gram ou ml — pas le même ratio 1:1 que Batch Yield"],
    ["Quantité (sous Recipe Unit)", "Grammes ou ml produits par lot", "4000 g frites · 950 ml sauce · 12 000 g chili"],
    ["Actual Cost Per ___", "Coût par unité recette (calculé)", "Vérifier à l'étape Recipe Costing"],
    ["Include on Count Sheets?", "Prep compté à l'inventaire physique", "Oui batch frites, chili, sauce ; Non portion frites"],
    ["Qty on Hand (stock prep)", "Non affiché sur fiche Prep en v5", "Reports → Prep Sheet (On Hand) ou Count Inventory → Summarize"],
    ["Main Location", "Où le stock prep est rangé", "Friterie, Ligne chaude, Réfrigérateur, Congélateur"],
    ["Barcode", "Scan OC Mobile (optionnel)", "Laisser vide"],
    ["Include on Prep Sheets?", "Prep Par Levels + rapports prep sheet", "Oui batch frites + chili ; Non les autres"],
    ["Prep Station", "Station responsable (prep sheet)", "Friterie ou Marmite selon prep"],
    ["Shelf Life", "Durée de conservation (jours)", "1 à 5 jours selon tableau"],
    ["Ingrédients (droite)", "Items ou autres Preps + qty + unité", "Tableau ingrédients ci-dessous ; Total = coût du batch"],
]

PREPS_CORE_HEADERS = [
    "Prep", "How is it used?", "Batch Uom", "Batch Yield", "Recipe Unit", "Qty recette",
    "Main Location", "Count sheet?", "Prep sheet?", "Prep Station", "Shelf life",
]

PREPS_CORE = [
    ["Boulette 80 g", "Unit", "Batch", "1 batch", "each", "1", "Congélateur", "Oui", "Non", "—", "2 j"],
    ["Frites maison — batch", "Weight", "Batch", "1 batch", "gram", "4000", "Friterie", "Oui", "Oui", "Friterie", "1 j"],
    ["Portion frites 200 g", "Unit", "Batch", "1 batch", "each", "1", "Friterie", "Non", "Non", "—", "1 j"],
    ["Sauce burger maison", "Volume", "Batch", "1 batch", "ml", "950", "Réfrigérateur", "Oui", "Non", "—", "5 j"],
    ["Chili — batch", "Weight", "Batch", "1 batch", "gram", "12000", "Ligne chaude", "Oui", "Oui", "Marmite", "3 j"],
]

PREPS_SUMMARY_HEADERS = [
    "Prep", "Batch Yield", "Shelf life", "Include prep sheet", "Station / note",
]

PREPS_SUMMARY = [
    ["Boulette 80 g", "1 each (80 g)", "2 jours", "Non", "77 g bœuf + 3 g sel"],
    ["Frites maison — batch", "4000 g", "1 jour", "Oui", "Pommes + huile + sel"],
    ["Portion frites 200 g", "1 each (200 g)", "1 jour", "Non", "Prep ← batch (nested)"],
    ["Sauce burger maison", "950 ml", "5 jours", "Non", "Mayo + ketchup"],
    ["Chili — batch", "12000 g", "3 jours", "Oui", "Haricots secs + eau cuisson"],
]

PREPS_INGREDIENT_HEADERS = [
    "Prep", "Type", "Ingrédient", "Qty", "UOM",
]

PREPS_INGREDIENTS = [
    ["Boulette 80 g", "Item", "Bœuf haché 80/20", "77", "gram"],
    ["Boulette 80 g", "Item", "Sel fin", "3", "gram"],
    ["Frites maison — batch", "Item", "Pommes de terre", "5000", "gram"],
    ["Frites maison — batch", "Item", "Huile friture", "200", "ml"],
    ["Frites maison — batch", "Item", "Sel fin", "50", "gram"],
    ["Portion frites 200 g", "Prep", "Frites maison — batch", "200", "gram"],
    ["Sauce burger maison", "Item", "Mayonnaise", "800", "ml"],
    ["Sauce burger maison", "Item", "Ketchup", "200", "ml"],
    ["Chili — batch", "Item", "Haricots rouges secs", "2000", "gram"],
    ["Chili — batch", "Item", "Bœuf haché chili", "3000", "gram"],
    ["Chili — batch", "Item", "Tomates", "1500", "gram"],
    ["Chili — batch", "Item", "Oignons", "500", "gram"],
    ["Chili — batch", "Item", "Sel fin", "40", "gram"],
]

PREPS_COST_HINTS = [
    ["Boulette 80 g", "≈ 0,65 $US / boulette", "77 g bœuf + 3 g sel"],
    ["Frites maison — batch", "≈ 0,0016 $US / g sortie", "5000 g pommes + 200 ml huile + 50 g sel → 4000 g"],
    ["Portion frites 200 g", "≈ 0,32 $US / portion", "200 g batch (prep nested)"],
    ["Sauce burger maison", "≈ 0,003 $US / ml", "800 ml mayo + 200 ml ketchup → 950 ml"],
    ["Chili — batch", "≈ 0,003 $US / g sortie", "5 items dont haricots secs → 12 000 g"],
]

# Livre Comprendre OC — champs écran New Prep (vidéo #16, labels v5)
PREPS_CREATE_NOTE = (
    "Recipe → Preps → New. Panneau gauche : Core Information, Inventory Information, "
    "Prep Sheet Information ; panneau droit : Ingrédients (boutons Add Item et Add Prep en haut). "
    "Onglets après sauvegarde : Recipe Instruction, Image, Location, Conversions, Ingredient Nutrition. "
    "Un Prep = transformation interne (sauce, batch, portion) avant usage dans un Product. "
    "Si le prix d'achat d'un ingrédient change (facture), le coût du Prep se met à jour partout."
)

PREPS_CREATE_DOC = [
    ("Core Information", "Description", "Nom unique de la préparation.", "Différent de tout item ou product existant."),
    ("Core Information", "Inventory Group", "Groupe d'inventaire du prep.", "Prep (groupe par défaut OC) ou groupe personnalisé."),
    ("Core Information", "How is it used?", "Poids, Volume ou Unité — verrouillé après Save.", "Weight : frites, chili. Volume : sauce. Unit : boulette, portion."),
    ("Core Information", "Batch Uom", "Contenant ou unité de production du lot.", "Batch (défaut) — ne pas mettre gram en Batch Uom si Recipe Unit = gram."),
    ("Core Information", "Batch Yield", "Quantité produite par lot.", "1 batch (exercice) ; la qty recette porte le poids réel (4000 g, 950 ml…)."),
    ("Core Information", "Recipe Unit (per ___)", "Unité pour portionner dans Products ou autres Preps.", "gram ou ml — voir note « 1 gram on a gram »."),
    ("Core Information", "Quantité (sous Recipe Unit)", "Poids ou volume total produit par 1 batch.", "4000 g frites · 950 ml sauce · 12 000 g chili — pas « 1 » si même unité que Batch Yield en gram."),
    ("Core Information", "Actual Cost Per ___", "Coût par unité recette (lecture seule).", "Vérifier après Save ; mis à jour si prix ingrédient change."),
    ("Core Information", "Key Inventory Item", "Prioritaire sur feuilles Key Items au comptage.", "Comme les items ; ON par défaut si visible."),
    ("Inventory Information", "Include on Count Sheets?", "Prep compté à l'inventaire physique.", "Oui : batch frites, chili, sauce ; Non : portion frites si on compte le batch."),
    ("Inventory Information", "Qty on Hand", "Stock courant du prep.", "Non affiché sur la fiche Prep en v5 — voir Reports → Prep Sheet (On Hand) ou Count Inventory → Summarize."),
    ("Inventory Information", "Main Location", "Emplacement principal du stock prep.", "Friterie, Ligne chaude, Réfrigérateur, Congélateur…"),
    ("Inventory Information", "Barcode", "Scan avec OC Mobile.", "Optionnel."),
    ("Prep Sheet Information", "Include on Prep Sheets?", "Prep Par Levels + rapports On Hand / Make.", "Oui : preps produits régulièrement (batch frites, chili)."),
    ("Prep Sheet Information", "Prep Station", "Station responsable sur la prep sheet.", "Friterie, Marmite… — créer la station en Setup si besoin."),
    ("Prep Sheet Information", "Shelf Life", "Durée de conservation (jours).", "Alimente Prep Par Levels ; ex. chili 3 j, sauce 5 j."),
    ("Ingrédients", "Add Item / Add Prep", "Items bruts ou autre Prep déjà créé.", "Créer les preps parents avant les enfants (batch frites avant portion)."),
    ("Ingrédients", "Quantity + Unit", "Quantité par ligne ; unités convertibles.", "Unité par défaut proposée ; menu déroulant si besoin."),
    ("Ingrédients", "Cost (colonne)", "Coût de la ligne ingrédient.", "Somme en bas = Total cost du batch."),
    ("Ingrédients", "Total (bas de fenêtre)", "Coût total des ingrédients pour le lot.", "Comparer avec Recipe Costing après Save."),
    ("Onglet Recipe Instruction", "Instructions / Import Word", "Méthode de préparation.", "Optionnel ; utile pour formation employés."),
    ("Onglet Image", "Browse image", "Photo du résultat.", "Apparaît sur fiche imprimée / Recipe Book."),
    ("Onglet Location", "Main + Secondary", "Plusieurs emplacements de stockage.", "Une ligne par location au comptage."),
    ("Onglet Conversions", "Add conversion", "Conversions supplémentaires.", "Rare pour les preps — généralement inutile."),
    ("Onglet Ingredient Nutrition", "Calculate nutrition", "Valeurs nutritionnelles agrégées.", "Si module nutrition actif."),
]

def prep_book_blocks():
    """Paragraphes Word pour Comprendre_Optimum_Control (rebuild_final.py)."""
    blocks = [
        ("Heading2", "Exercice Comptoir — 5 preps (étape 22)"),
        ("Normal", PREPS_NOTE),
        ("Normal", PREPS_NESTED_NOTE),
        ("Normal", PREPS_YIELD_NOTES),
        ("Normal", "Inventory Group = Prep (défaut OC) ; Batch Uom = Batch pour tous les preps ci-dessous."),
        ("Heading3", "Valeurs par prep — Core, Inventory et Prep Sheet"),
    ]
    for row in PREPS_CORE:
        name, how, buom, batch, runit, rqty, loc, count, psheet, station, shelf = row
        blocks.append((
            "Normal",
            f"• {name} — How is it used? : {how} ; Batch Uom : {buom} ; Batch Yield : {batch} ; "
            f"Recipe Unit : {runit} (qty {rqty}). Main Location : {loc} ; "
            f"Count sheet : {count} ; Prep sheet : {psheet} ; Station : {station} ; "
            f"Shelf life : {shelf}.",
        ))
    blocks.append(("Normal", PREPS_BATCH_UNIT_NOTE))
    blocks.append(("Normal", PREPS_ACTUALIZE_NOTE))
    blocks.append(("Heading3", "Ingrédients — Items et Preps"))
    current = None
    for prep, typ, ing, qty, uom in PREPS_INGREDIENTS:
        if prep != current:
            current = prep
            blocks.append(("Normal", f"{prep} :"))
        blocks.append(("Normal", f"  – {typ} : {ing}, {qty} {uom}"))
    blocks.append(("Heading3", "Coûts attendus (Recipe Costing)"))
    for prep, cost, calc in PREPS_COST_HINTS:
        blocks.append(("Normal", f"• {prep} — {cost} ({calc})"))
    blocks.append((
        "Normal",
        "Product Frites moyennes (POS 201) : utiliser le prep Portion frites 200 g (1 each), "
        "pas le batch Frites maison directement.",
    ))
    return blocks

def prep_field_guide_rows():
    return [list(row) for row in PREPS_FIELD_GUIDE]

def prep_core_rows():
    return [list(row) for row in PREPS_CORE]

def prep_summary_rows():
    return [list(row) for row in PREPS_SUMMARY]

def prep_ingredient_rows():
    return [list(row) for row in PREPS_INGREDIENTS]

def prep_cost_hint_rows():
    return [list(row) for row in PREPS_COST_HINTS]

# ---- Products (Partie 4 / étape 24) ----
PRODUCTS_NOTE = (
    "Recipe → Products → New. Un Product = plat vendu au menu ; lie POS ID# aux ventes importées. "
    "Ingrédients : Add Item, Add Prep ou Add Product (ex. Burger bacon peut réutiliser Burger classique)."
)

# ---- POS ID# — liaison caisse ↔ Product (vidéo #21, ch.9) ----
POS_ID_NOTE = (
    "Le **POS ID#** (aussi **PLU** sur la caisse) est le **numéro du bouton** sur le système de caisse (Point Of Sale). "
    "Quand le serveur vend un « Burger classique », la caisse enregistre par exemple **101**. "
    "Sur la fiche Product, ce champ dit à Optimum Control : « quand l'import contient 101, c'est **ce** plat — "
    "extraire **ces** ingrédients du stock ». Ce n'est pas un code inventé par OC : c'est le numéro que **votre caisse** "
    "envoie dans le fichier de ventes (CSV ou intégration)."
)

POS_ID_CHAIN_NOTE = (
    "Chaîne complète : **bouton caisse** (PLU 101) → **import** (CSV ou Sales Export) → **Product OC** « Burger classique » "
    "(POS ID# 101 + recette) → **usage idéal** (pain, bœuf, sauce…) → **Usage Summary** (écart réel vs théorique). "
    "Sans POS ID renseigné, l'import arrive en **Unlinked** — OC ne sait pas quelle recette appliquer."
)

POS_ID_NOT_NOTE = (
    "Ne pas confondre : le POS ID n'est **pas** le nom du plat (champs Description / POS Description), "
    "ni un code **Item** d'inventaire (pain, bœuf). **Un POS ID = un Product** (plat menu vendu), pas un ingrédient. "
    "Règle : **un POS ID ne doit pointer que vers un seul Product** — sinon OC ne sait plus quelle recette utiliser."
)

POS_ID_TILL_TAPE_NOTE = (
    "**Till Tape** (saisie manuelle) : vous entrez les quantités par **nom de Product** — pas de colonne POS ID à l'écran. "
    "Le POS ID reste sur la fiche Product et sert surtout à l'**import CSV** et à **Pending Sales**. "
    "Les deux méthodes (Till Tape et import) alimentent le même Sales Mix et le même calcul de stock."
)

POS_ID_MULTI_HEADERS = ["Situation", "Ce qui se passe", "Que faire dans OC"]
POS_ID_MULTI = [
    [
        "Plusieurs terminaux, **même caisse** (2 iPads, 2 caisses)",
        "Même plan de boutons — PLU 101 = toujours le burger, peu importe la caisse",
        "**1 Product · 1 POS ID#** — toutes les ventes se cumulent dans le même export",
    ],
    [
        "**Même plat**, 2 boutons caisse (salle vs livraison, ex. 101 et 1101)",
        "La caisse envoie **deux numéros PLU** différents pour le même assiette",
        "**2 Products** (chacun son POS ID#) — ou le 2e réutilise le 1er en ingrédient (comme Burger bacon → Burger classique)",
    ],
    [
        "**Plusieurs systèmes** de caisse (salle + plateforme livraison)",
        "Chaque système exporte un fichier au format différent",
        "**1 spec POS** par format (Preferences → POS) ; lier chaque PLU via **Pending Sales → Link To**",
    ],
    [
        "**Plusieurs magasins** (OC Enterprise / multi-store)",
        "Même plat, numéro PLU **différent par succursale**",
        "Champ **POS IDs** au pluriel par établissement (hors exercice mono-magasin Le Comptoir)",
    ],
]

POS_ID_COMPTOIR_HEADERS = ["POS ID# (PLU)", "Product OC", "Note"]
POS_ID_COMPTOIR = [
    ["101", "Burger classique", "Till Tape mar. : 25 vendus"],
    ["102", "Burger bacon", "Réutilise Burger classique en ingrédient"],
    ["201", "Frites moyennes", "Accompagnement"],
    ["301", "Salade César", "Entrée / plat"],
    ["401", "Bol chili", "Selling Price 9,95 $US"],
    ["501", "Cola 355 ml", "Boisson — lier si CSV dit « Cola »"],
    ["502", "Eau 500 ml", "Boisson"],
]

POS_ID_FIELD_HINT = (
    "Pricing Information → **POS ID#** : numéro PLU caisse (unique par Product en mono-magasin). "
    "**POS Description** : libellé exact du bouton caisse — utile si **Mismatched** à l'import. "
    "Laisser vide à la création si inconnu ; remplir via **Link Products to POS** ou Pending Sales."
)

def pos_id_multi_rows():
    return [list(row) for row in POS_ID_MULTI]

def pos_id_comptoir_rows():
    return [list(row) for row in POS_ID_COMPTOIR]

def pos_id_book_blocks():
    """Paragraphes Word — POS ID (rebuild_final.py, ch.9 / Products)."""
    blocks = [
        ("Heading2", "POS ID# — comprendre la liaison caisse ↔ Optimum Control"),
        ("Normal", POS_ID_NOTE),
        ("Normal", POS_ID_CHAIN_NOTE),
        ("Normal", POS_ID_NOT_NOTE),
        ("Heading3", "Till Tape vs import CSV"),
        ("Normal", POS_ID_TILL_TAPE_NOTE),
        ("Heading3", "Même produit vendu sur plusieurs POS ?"),
    ]
    for sit, happens, action in POS_ID_MULTI:
        blocks.append(("Normal", f"• **{sit}** — {happens}. → {action}"))
    blocks.append(("Heading3", "Exercice Le Comptoir — table de correspondance"))
    blocks.append((
        "Normal",
        "Caisse fictive « Comptoir CSV » : chaque ligne du fichier ventes_*.csv utilise ces PLU en colonne 2.",
    ))
    for pos, prod, note in POS_ID_COMPTOIR:
        blocks.append(("Normal", f"• PLU **{pos}** → {prod} — {note}"))
    blocks.append((
        "Normal",
        "Après création des 7 Products avec POS ID# (étape 24), un import mer→sam doit afficher "
        "**Valid** dans Pending Sales — pas Unlinked.",
    ))
    return blocks

PRODUCTS_SCREEN_INTRO = (
    "Product = recette « à l'assiette » (menu). Panneau gauche : Core Information + Pricing Information + Comments ; "
    "panneau droit : Ingrédients. Onglets : Recipe Instruction, Ingredient Nutrition, Allergènes (si module actif)."
)

PRODUCTS_FIELD_GUIDE_HEADERS = ["Section / champ (OC)", "Rôle", "Exercice Le Comptoir"]

PRODUCTS_FIELD_GUIDE = [
    ["Description", "Nom du plat dans OC (peut différer du POS)", "Voir tableau « Valeurs par product »"],
    ["Sales Group", "Groupe de vente (Setup → Sales Groups)", "Entrées / Plats, Accompagnements ou Boissons"],
    ["Barcode", "Scan retail (rare en resto)", "Laisser vide"],
    ["Actual Product Cost", "Coût recette calculé (somme ingrédients)", "Vérifier après Save — Cost Percent se remplit"],
    ["Key Product", "Prioritaire sur certains rapports / filtres", "Décoché sauf besoin"],
    ["POS ID#", "Numéro PLU caisse — lien import ventes", POS_ID_FIELD_HINT],
    ["POS Description", "Libellé exact caisse (matching import)", "Optionnel ; utile si nom POS différent"],
    ["Selling Price", "Prix menu TTC ou HT selon votre caisse", "14,95 $US burger classique, etc."],
    ["Gross Margin / Cost % / Gross %", "Marge et food cost % (calculés)", "Noter Cost Percent — cible exercice ~32 %"],
    ["Tax Group", "Taxes sur vente (si applicable)", "[No Tax Group] pour l'exercice"],
    ["Comments", "Notes libres", "Optionnel"],
    ["Ingrédients (droite)", "Items, Preps ou Products + qty + unité", "Tableau ingrédients ci-dessous"],
]

PRODUCTS_CORE_HEADERS = [
    "Product", "Sales Group", "POS ID#", "Selling Price", "Tax Group", "Key Product?",
]

PRODUCTS_CORE = [
    ["Burger classique", "Entrées / Plats", "101", "14,95 $US", "No Tax Group", "Non"],
    ["Burger bacon", "Entrées / Plats", "102", "16,95 $US", "No Tax Group", "Non"],
    ["Frites moyennes", "Accompagnements", "201", "5,50 $US", "No Tax Group", "Non"],
    ["Salade César", "Entrées / Plats", "301", "13,50 $US", "No Tax Group", "Non"],
    ["Bol chili", "Entrées / Plats", "401", "9,95 $US", "No Tax Group", "Non"],
    ["Cola", "Boissons", "501", "3,25 $US", "No Tax Group", "Non"],
    ["Eau", "Boissons", "502", "2,75 $US", "No Tax Group", "Non"],
]

PRODUCTS_INGREDIENT_HEADERS = [
    "Product", "Type", "Ingrédient", "Qty", "UOM",
]

PRODUCTS_INGREDIENTS = [
    ["Burger classique", "Item", "Pain burger", "1", "each"],
    ["Burger classique", "Prep", "Boulette 80 g", "1", "each"],
    ["Burger classique", "Prep", "Sauce burger maison", "15", "ml"],
    ["Burger classique", "Item", "Tomates", "30", "gram"],
    ["Burger classique", "Item", "Boîte burger", "1", "each"],
    ["Burger classique", "Item", "Serviette", "1", "each"],
    ["Burger bacon", "Product", "Burger classique", "1", "each"],
    ["Burger bacon", "Item", "Bacon tranché", "2", "each"],
    ["Frites moyennes", "Prep", "Portion frites 200 g", "1", "each"],
    ["Frites moyennes", "Item", "Serviette", "1", "each"],
    ["Salade César", "Item", "Laitue romaine", "150", "gram"],
    ["Salade César", "Item", "Parmesan râpé", "15", "gram"],
    ["Salade César", "Item", "Croûtons", "20", "gram"],
    ["Salade César", "Item", "César dressing", "45", "ml"],
    ["Bol chili", "Prep", "Chili — batch", "400", "gram"],
    ["Bol chili", "Item", "Fourchette plastique", "1", "each"],
    ["Cola", "Item", "Cola 355 ml", "1", "each"],
    ["Eau", "Item", "Eau 500 ml", "1", "each"],
]

PRODUCTS_NESTED_NOTE = (
    "Product dans product (exercice) : Burger bacon = Add Product « Burger classique » (1 each) + 2 tranches bacon. "
    "Alternative : recopier toutes les lignes du burger + bacon. "
    "Frites moyennes : prep Portion frites 200 g — pas le batch frites."
)

PRODUCTS_TOMATES_NOTE = (
    "Tomates dans les products : l'item Tomates est en gram (How is it used? = Weight). "
    "Burger classique → 30 gram (pas « 2 tranches » ni « 3 tranches » dans OC). "
    "En cuisine ≈ 2 tranches de ~15 g chacune — équivalent indicatif seulement."
)

SALADE_CESAR_NOTE = (
    "Salade César (POS 301) : **4 Items seulement** — laitue, parmesan, croûtons, César dressing. "
    "Pas de poulet dans cet exercice : les **20 items** du Comptoir sont fixes (étapes 17–19) et « Poulet grillé » "
    "n'en fait pas partie. Si vous l'avez ajouté sur la fiche Product par erreur, supprimez la ligne ingrédient."
)

PRODUCTS_CREATE_NOTE = (
    "Recipe → Products → New. Sales Group obligatoire (bordure rouge si vide). "
    "POS ID# unique — lien avec Sales Mix. Actual Product Cost et Cost Percent se calculent après ingrédients. "
    "Rapports : Recipe Book (Products), Menu Product Detail Problems (seuil food cost)."
)

PRODUCTS_CREATE_DOC = [
    ("Core Information", "Description", "Nom du plat dans OC.", "Différent des items/preps ; peut différer du libellé POS."),
    ("Core Information", "Sales Group", "Groupe Setup → Sales Groups.", "Entrées / Plats, Accompagnements, Boissons — hérite Sales Cat."),
    ("Core Information", "Barcode", "Scan (retail).", "Vide en resto."),
    ("Core Information", "Actual Product Cost", "Coût total recette (lecture seule).", "Somme colonne Cost ingrédients."),
    ("Core Information", "Key Product", "Filtre / rapports.", "Décoché par défaut."),
    ("Pricing Information", "POS ID#", "PLU caisse — numéro du bouton POS.", "101–502 exercice ; unique par Product (mono-magasin)."),
    ("Pricing Information", "POS Description", "Libellé bouton caisse pour matching.", "Utile si Mismatched ; peut différer de Description OC."),
    ("Pricing Information", "Selling Price", "Prix menu.", "Alimente marge et Cost Percent."),
    ("Pricing Information", "Gross Margin / Cost % / Gross %", "Calculés automatiquement.", "Cost Percent = food cost %."),
    ("Pricing Information", "Tax Group", "Taxe vente.", "No Tax Group — exercice."),
    ("Comments", "Comments", "Notes.", "Optionnel."),
    ("Ingrédients", "Add Item / Prep / Product", "Composition du plat.", "Voir tableau exercice ; Product nested pour Burger bacon."),
    ("Ingrédients", "Quantity + Unit", "Portion par assiette.", "Unités convertibles."),
    ("Onglet Recipe Instruction", "Instructions / image", "Méthode ou photo du plat.", "Optionnel."),
    ("Onglet Ingredient Nutrition", "Calculate / From Label", "Nutrition agrégée.", "Si module actif."),
    ("Onglet Allergènes", "Allergènes hérités", "Depuis items liés.", "Si module actif."),
]

def product_book_blocks():
    blocks = [
        ("Heading2", "POS ID# — rappel (voir aussi « Lier les produits POS »)"),
        ("Normal", POS_ID_NOTE),
        ("Normal", POS_ID_CHAIN_NOTE),
        ("Normal", POS_ID_TILL_TAPE_NOTE),
        ("Heading2", "Exercice Comptoir — 7 products (étape 24)"),
        ("Normal", PRODUCTS_NOTE),
        ("Normal", PRODUCTS_NESTED_NOTE),
        ("Normal", PRODUCTS_TOMATES_NOTE),
        ("Normal", SALADE_CESAR_NOTE),
        ("Heading3", "Valeurs par product — Core et Pricing"),
    ]
    for row in PRODUCTS_CORE:
        name, sg, pos, price, tax, keyp = row
        blocks.append((
            "Normal",
            f"• {name} — Sales Group : {sg} ; POS ID# : {pos} ; Selling Price : {price} ; "
            f"Tax Group : {tax} ; Key Product : {keyp}.",
        ))
    blocks.append(("Heading3", "Ingrédients — Items, Preps et Products"))
    current = None
    for prod, typ, ing, qty, uom in PRODUCTS_INGREDIENTS:
        if prod != current:
            current = prod
            blocks.append(("Normal", f"{prod} :"))
        blocks.append(("Normal", f"  – {typ} : {ing}, {qty} {uom}"))
    return blocks

def product_field_guide_rows():
    return [list(row) for row in PRODUCTS_FIELD_GUIDE]

def product_core_rows():
    return [list(row) for row in PRODUCTS_CORE]

def product_ingredient_rows():
    return [list(row) for row in PRODUCTS_INGREDIENTS]

# Types de compte OC — champ « Account Type » dans Settings → Setup → Accounts
ACCOUNT_TYPES = ("CostOfSales", "Asset", "Expense", "Income", "Liability")

# Type OC, signification, exemple dans un restaurant
ACCOUNT_TYPES_DOC = [
    (
        "CostOfSales",
        "Coût des ventes — ce que vous consommez pour produire ce que vous vendez "
        "(matières, boissons servies, emballage lié aux plats).",
        "Food Cost, Beverage Cost, Paper / Supplies",
    ),
    (
        "Asset",
        "Actif — créance ou avoir (somme que l'entreprise a payée et peut récupérer).",
        "Dépôts consignes, avances fournisseurs",
    ),
    (
        "Expense",
        "Charge — frais d'exploitation qui ne sont pas le coût direct des ventes.",
        "Loyer, salaires admin, réparations, publicité",
    ),
    (
        "Income",
        "Revenu — entrées d'argent liées à l'activité (ventes).",
        "Sales Food, Sales Beverage",
    ),
    (
        "Liability",
        "Passif — dans OC, type recommandé pour les comptes de taxes (manuel TracRite). "
        "Visible dans Taxes → Account (avec CostOfSales et Expense ; pas Asset ni Income).",
        "TCA recoverable, TCA à remettre, GST, dépôts consignes",
    ),
]

# Nom OC, GL ref (export comptable), Account Type
COMPTES_GL = [
    ("Food Cost", "5100", "CostOfSales"),
    ("Beverage Cost", "5101", "CostOfSales"),
    ("Paper / Supplies", "5120", "CostOfSales"),
    ("TCA recoverable", "2310", "Liability"),
    ("Sales Food", "4100", "Income"),
    ("Sales Beverage", "4101", "Income"),
]

# Setup OC — trois écrans liés (vidéo #02 + manuel)
# 1) Categories : on CRÉE les catégories ici (+ Income Account)
# 2) Inventory Groups : Group desc. | Sales Cat. (dropdown) | Account
# 3) Sales Groups : Group desc. | Sales Cat. (dropdown) — pour les Products
# Le champ « Sales Cat. » / « Sales Category » partout = liste des Categories créées à l'étape 1

CATEGORIES = [
    ("Food", "Sales Food", "Plats : burgers, frites, salade, chili (+ emballages liés)"),
    ("Beverage", "Sales Beverage", "Boissons : Cola, eau"),
]

# Inventory Groups : Group desc. | Sales Cat. | Account
INVENTORY_GROUPS = [
    ("Proteins", "Food", "Food Cost"),
    ("Produce", "Food", "Food Cost"),
    ("Dry Goods", "Food", "Food Cost"),
    ("Spices", "Food", "Food Cost"),
    ("Dairy", "Food", "Food Cost"),
    ("Paper", "Food", "Paper / Supplies"),
    ("Soft Drinks", "Beverage", "Beverage Cost"),
]

# Sales Groups : Group desc. | Sales Cat. (dropdown = Categories)
SALES_GROUPS = [
    ("Entrées / Plats", "Food"),
    ("Accompagnements", "Food"),
    ("Boissons", "Beverage"),
]

SETUP_CATEGORIES_NOTE = (
    "OC a trois écrans Setup distincts : Categories, Inventory Groups et Sales Groups. "
    "Les catégories se créent dans Categories (avec Income Account). "
    "Partout ailleurs, le champ Sales Cat. / Sales Category est un menu déroulant "
    "qui affiche exactement ces catégories."
)

# Settings → Preferences → Inventory (écran OC — pas confondre avec Key Item / Actualize sur fiche Item)
INVENTORY_PREFS_NOTE = (
    "Settings → Preferences → Inventory. Key Item et Actualize Usage Values se règlent sur "
    "chaque fiche Item (pas sur l'écran Preps en v5 — vidéo #16)."
)

INVENTORY_PREFS_HEADERS = ["Champ", "Valeur exercice", "Rôle"]

INVENTORY_PREFS_DOC = [
    ("Value Inventory Using", "FIFO", "Valorise le stock au coût des premiers achats (First-In, First-Out)."),
    ("", "Last Cost (alt.)", "Alternative : valorise au dernier prix de caisse (fiche item) — utile pour réévaluation."),
    ("", "Weighted Average (alt.)", "Alternative : moyenne pondérée des achats de la période."),
    ("Ask to Summarize on Close", "Coché", "Rappel avant de fermer un comptage : ne pas oublier Summarize."),
    ("Require Reason for Inventory Adjustments", "Coché", "Oblige un motif si vous modifiez une quantité sur la feuille."),
    ("Countsheet Column Display", "All Columns", "Colonnes au comptage : unité d'achat, caisse (split) et recette."),
    ("Inventory Interface", "[None]", "Intégration bar (Freepour, Bluestreak) — laisser None pour cet exercice."),
    ("Print countsheet groupings on separate pages", "Décoché", "Chaque regroupement (location/groupe) sur une page séparée à l'impression."),
    ("Last Prep Amount Calculation Range", "(vide)", "Plage de dates pour le calcul Prep Par Levels (rapports prep sheet)."),
    ("Default Prep Margin", "10 %", "Marge par défaut sur les prep sheets (+10 % sur quantités calculées)."),
    ("Prep Amount Factor", "100 %", "Facteur appliqué aux quantités de prep calculées."),
    ("Warning Threshold (over stock)", "2", "Alerte jaune si quantité en stock dépasse le par × 2."),
    ("Critical Threshold (over stock)", "3", "Alerte rouge si quantité dépasse le par × 3."),
    ("Disable perpetual amounts for non-counted items", "Décoché", "Si coché : items non comptés ne recalculent pas leur qty théorique."),
    ("Set negative quantities to zero", "Décoché", "Force les quantités négatives à 0 après comptage."),
]

def inventory_prefs_rows():
    return [list(row) for row in INVENTORY_PREFS_DOC]

# Settings → Preferences → POS (OC v5 — plus de menu « Setup → Configure POS »)
POS_PREFS_NOTE = (
    "Dans OC v5 : Settings → Preferences → POS (pas Setup → Configure POS). "
    "L'icône engrenage à côté de POS configuration ouvre le mappage des colonnes CSV."
)

CSV_IMPORT_NOTE = (
    "Spec **Comptoir CSV** (engrenage POS) : **5 colonnes** — 1=Date, 2=POS ID, 3=Description, 4=Quantity, "
    "5=**Gross Sales** (total de **ligne** = Quantity × prix unitaire, **pas** le prix unitaire seul). "
    "Champs obligatoires (*) : PLU Number, PLU Description, Amount Sold, **Gross Sales** — laisser **Selling Price** vide. "
    "Si la col. 5 contient 14,95 au lieu de 373,75 (25×14,95), **Avg Price** sera rouge (14,95÷25≈0,60). "
    "Régénérez les CSV : `python generate_ventes_csv.py`."
)

CSV_IMPORT_HEADERS = ["Index colonne", "Champ OC (Comptoir CSV)", "En-tête CSV / contenu col. 5"]
CSV_IMPORT_MAP = [
    ["1", "Sales Date", "Date"],
    ["2", "*PLU Number", "POS ID"],
    ["3", "*PLU Description", "POS Description"],
    ["4", "*Amount Sold", "Quantity"],
    ["5", "*Gross Sales", "Gross Sales (= Qté × prix unitaire)"],
    ["—", "Selling Price", "(vide — ne pas mapper si Gross Sales = 5)"],
]

def csv_import_map_rows():
    return [list(row) for row in CSV_IMPORT_MAP]

POS_PREFS_HEADERS = ["Champ", "Valeur exercice", "Rôle"]

POS_PREFS_DOC = [
    ("Default POS import folder", "exercice_comptoir/ventes_csv/", "Dossier par défaut quand Sales → New → … (vidéo #42)."),
    ("POS configuration", "New Pos Specification + engrenage", "Assistant Import Specification Settings → Import Specification Fields."),
    ("Update selling price after import?", "Never", "Ne pas écraser le prix menu OC à chaque import."),
    ("Default Category when Creating Product", "Food", "Catégorie par défaut si vous créez un product depuis Pending Sales."),
    ("Type of sales to view in reporting", "Gross Sales", "Ventes brutes dans les rapports (vs Net Sales)."),
    ("Import Sales at Start?", "Décoché", "Ne pas importer automatiquement au lancement d'OC."),
]

# Écran 1 — Import Specification Settings (engrenage POS)
POS_SPEC_WIZARD_HEADERS = ["Champ assistant", "Valeur type CSV exercice"]
POS_SPEC_WIZARD_DOC = [
    ("Configuration", "[New Pos Specification]"),
    ("Description", "Comptoir CSV (nom libre)"),
    ("File Format", "Csv"),
    ("Extension", ".csv"),
    ("File Path", "Fichier exemple : ventes_semaine_comptoir.csv"),
    ("Date Format", "yyyy-MM-dd"),
    ("Delimiter", ", (virgule)"),
    ("Lines To Skip", "1 — la 1re ligne est l'en-tête"),
    ("String in Quote", "Décoché"),
]

def pos_spec_wizard_rows():
    return [list(row) for row in POS_SPEC_WIZARD_DOC]

# Écran Import Specification Fields — colonnes du fichier ventes_semaine_comptoir.csv
# (1=Date, 2=POS ID, 3=POS Description, 4=Quantity, 5=Gross Sales = Qté × prix unitaire)
POS_SPEC_HEADERS = ["Champ OC (Index)", "Colonne CSV", "Exemple ligne 2 (mar., 25 burgers)"]
POS_SPEC_DOC = [
    ("*PLU Number → 2", "POS ID", "101"),
    ("*PLU Description → 3", "POS Description", "Burger classique"),
    ("*Amount Sold → 4", "Quantity", "25"),
    ("*Gross Sales → 5", "Gross Sales", "373,75 (= 25 × 14,95 — total ligne, pas prix unitaire)"),
    ("Selling Price", "(vide)", "Ne pas mapper — Gross Sales (*) suffit"),
    ("Sales Date → 1", "Date", exo_date_to_iso(EXO_DAY_TO_DATE[EXO_FIRST_SALES_DAY])),
    ("Department", "(vide)", "—"),
    ("Net Sales", "(vide)", "—"),
    ("Promo Number", "(vide)", "—"),
    ("Group Name", "(vide)", "—"),
]

def pos_spec_rows():
    return [list(row) for row in POS_SPEC_DOC]

def pos_prefs_rows():
    return [list(row) for row in POS_PREFS_DOC]

# ---- Till Tape / Sales Mix manuel (Phase F — vidéo #43) ----
TILL_TAPE_NOTE = (
    "Un **Till Tape** dans Optimum Control, ce n'est pas un fichier : c'est la **saisie manuelle du Sales Mix** "
    "(ce qui a été vendu), comme si vous recopiez le **Z de caisse** quand il n'y a pas d'import POS automatique. "
    "OC enregistre les quantités par **Product** (recette menu), calcule l'**usage idéal** et **diminue le stock** au Save."
)

TILL_TAPE_TWO_STEP_NOTE = (
    "Deux écrans différents — ne pas confondre :\n"
    "• **Till Tape List** (Sales → Till Tapes → New) = **modèle** : quels products entrer (7 du menu). "
    "Pas de date, pas de quantités, pas de colonne POS ID.\n"
    "• **Till Tape Sale** (Sales → New → Till Tape) = **vente d'un jour** : date + quantités vendues (# Sold) "
    "à côté du **nom du product** (Burger classique, Cola…). Le POS ID (101, 102…) est déjà sur la fiche Product — "
    "il sert surtout à l'import CSV et Pending Sales."
)

TILL_TAPE_POS_NOTE = (
    "**Preferences → POS → Comptoir CSV** (étape 14) sert à l'**import de fichiers CSV**, pas à créer la liste Till Tape. "
    "Ne **désélectionnez pas** Comptoir CSV pour « débloquer » Till Tape — sans config POS, **Sales → New** peut tout griser "
    "(message *Must Select Sales Configuration*). En pratique v5 : **créez et Save la liste Till Tape d'abord** "
    "(Description + 7 products) ; ensuite **Sales → New → Till Tape** se dégrise. "
    "Mar→ven : **Import From File** avec Comptoir CSV (étape 33). "
    "Alternative 1re vente : importer ventes_mardi_2026-06-30.csv — même résultat qu'un Till Tape manuel."
)

TILL_TAPE_LIST_FLOW = [
    "Sales → Till Tapes (ruban) → New.",
    "Description obligatoire : ex. « Tous les produits ».",
    "Gauche : les 7 Products — Burger classique, Burger bacon, Frites moyennes, Salade César, Bol chili, Cola, Eau.",
    "Double-clic ou flèche → panneau droit (Till Tape Products). Shift+clic pour tout sélectionner.",
    "Save → fermer. C'est une configuration unique — pas la vente du jour.",
]

TILL_TAPE_SALE_FLOW = [
    f"Sales → New (ruban, pas Till Tapes) → Till Tape / New Till Tape Sale.",
    f"Sales Date : {EXO_DAY_TO_DATE[EXO_FIRST_SALES_DAY]} ({exo_date_to_iso(EXO_DAY_TO_DATE[EXO_FIRST_SALES_DAY])}) — **mardi**.",
    "Choisir la liste « Tous les produits » → Continue.",
    "Saisir # Sold à côté de chaque nom de product (pas de colonne POS — voir tableau ci-dessous).",
    "Save — OC extrait le stock (ex. Pain burger : 48 → 13 ea après 25 burgers + 10 bacon burgers).",
]

TILL_TAPE_QTY_HEADERS = [
    "Product (nom à l'écran)", "POS ID (fiche Product)",
    f"Qté {EXO_DAY_TO_DATE[EXO_FIRST_SALES_DAY]} (mar.)",
]
TILL_TAPE_MARDI_QTY = [
    ["Burger classique", "101", "25"],
    ["Burger bacon", "102", "10"],
    ["Frites moyennes", "201", "30"],
    ["Salade César", "301", "8"],
    ["Bol chili", "401", "15"],
    ["Cola 355 ml", "501", "40"],
    ["Eau 500 ml", "502", "12"],
]

def till_tape_mardi_rows():
    return [list(row) for row in TILL_TAPE_MARDI_QTY]

# Alias legacy scripts
till_tape_lundi_rows = till_tape_mardi_rows

TILL_TAPE_VALIDATE = [
    "Liste Till Tape Save (7 products).",
    f"Ventes {EXO_DAY_TO_DATE[EXO_FIRST_SALES_DAY]} (mardi) Save — 140 unités au total.",
    "Calendrier Sales : date du mardi affichée.",
    "Qty on Hand baissée (ex. Pain ≈ 13 ea, Cola −40).",
    "Imports mer→sam faits (étape 33) — puis Pending Sales (étape 34).",
]

DAILY_SALES_GROSS_NOTE = (
    "Daily Sales (étape 36) : le tableau indique des **totaux Gross arrondis** (ex. mardi 1 420 $US). "
    "Le Till Tape calculé depuis vos **Selling Price** OC peut différer (ex. ≈ 1 182 $US) — "
    "saisissez le montant **simulation** du tableau pour l'exercice comptable."
)

# ---- Pending Sales (Phase F — étape 34, vidéo #44) ----
PENDING_SALES_NOTE = (
    "Après un **import CSV** (ou parfois un Till Tape), Optimum Control classe chaque ligne de vente POS "
    "dans l'écran **Pending Sales Mix** (double-clic sur un jour du calendrier Sales). "
    "Tant qu'une ligne reste **Unlinked**, OC **ne calcule pas l'usage idéal** pour ce produit — "
    "le stock n'est pas correctement extrait par les recettes. "
    "Objectif de l'exercice : **0 Unlinked** (toutes les lignes **Valid** ou **Ignored** volontairement)."
)

PENDING_SALES_STATES_HEADERS = ["État OC", "Signification", "Action"]
PENDING_SALES_STATES = [
    [
        "Valid (liée)",
        "Le POS ID (ex. 101) correspond à un Product OC — description OK.",
        "Rien à faire — la vente compte dans le Sales Mix et l'usage idéal.",
    ],
    [
        "Unlinked (non liée)",
        "Le POS ID du fichier CSV **n'est associé à aucun Product** dans OC.",
        "**Link To** (loupe) → choisir le Product · **Create Product** · ou **Ignore** (hors menu).",
    ],
    [
        "Mismatched (incohérente)",
        "Le POS ID existe, mais la **description POS** du fichier ≠ celle attendue sur la fiche Product.",
        "**Switch Description** si c'est le même plat (bouton renommé en caisse). Sinon **Unlink** (voir ci-dessous).",
    ],
    [
        "Ignored (ignorée)",
        "Ligne exclue du calcul (pourboires, « bien cuit », etc.).",
        "**Include** pour annuler un ignore par erreur.",
    ],
]

PENDING_SALES_UNLINK_NOTE = (
    "**Unlink** n'apparaît que sur une vente **Mismatched** — pas sur Unlinked.\n"
    "• **Unlinked** = OC ne sait **pas encore** quel Product lier (POS ID inconnu).\n"
    "• **Unlink** = vous **cassez** un lien existant parce que la description ne correspond **pas** au bon Product "
    "(ex. le fichier dit « Burger meal » mais le lien pointe vers « Burger classique » par erreur). "
    "La ligne retourne en section **Unlinked** ; vous pouvez alors **Link To** le bon Product ou **Create Product**.\n"
    "• **Switch Description** = même Product, seul le libellé POS a changé — préférez cette option si c'est le même plat.\n"
    "Dans l'exercice Comptoir : vos 7 Products ont déjà POS ID 101–502 (étape 24). "
    "Après import, tout devrait être **Valid**. Si une ligne est Unlinked, vérifiez le POS ID sur la fiche Product "
    "(Recipe → Products) puis **Link To**."
)

PENDING_SALES_FLOW = [
    "Sales → calendrier : icône « yield » (triangle) = ventes en attente — survolez pour voir Unlinked / Mismatched.",
    "Double-clic sur le jour (ex. 01/07/2026) → fenêtre Pending Sales Mix.",
    "Traiter d'abord **Unlinked** : surligner la ligne → loupe **Link To** → double-clic sur le Product (ex. Burger classique pour POS 101).",
    "Puis **Mismatched** : **Switch Description** si même produit ; **Unlink** seulement si mauvais lien.",
    "Ne pas ignorer les 7 produits du menu — ils doivent rester **Valid**.",
    "Save en haut → Open pour revenir au calendrier. Répéter pour chaque jour importé si besoin.",
    "Calendrier : plus d'icône pending sur les jours traités ; totaux Gross inchangés si déjà importés.",
]

PENDING_SALES_COMPTOIR_NOTE = (
    "Cas typique exercice : import mer→sam avec Comptoir CSV — les 7 POS ID sont déjà sur vos Products. "
    "Vous ouvrez Pending Sales par précaution : tout est **Valid** → Save sans modification. "
    "Si **Unlinked** sur « Cola 355 ml » : Link To → product Cola (POS 501). "
    "Erreur fréquente : avoir créé le Product sans POS ID — le corriger sur la fiche Product évite les Unlinked futurs."
)

PENDING_SALES_VALIDATE = [
    "Chaque jour importé (01/07 → 04/07) : 0 Unlinked.",
    "7 lignes Valid par jour (ou équivalent regroupé par POS ID).",
    "Aucun des 7 products du menu en Ignored.",
    "Icône pending absente sur le calendrier Sales (ou seulement Ignored hors menu).",
]

def pending_sales_state_rows():
    return [list(row) for row in PENDING_SALES_STATES]

# ---- Waste (Phase F — étape 35, vidéo #45) ----
WASTE_NOTE = (
    "Le **Waste** (pertes) enregistre ce qui a été **jeté ou perdu** sans être vendu : fané, brûlé, fin de shelf life, etc. "
    "OC **diminue le stock** (comme une vente) et affiche le montant dans **Usage Summary** (colonne Waste). "
    "Sans waste saisi, un écart « mystère » peut apparaître sur la laitue ou le chili à la clôture. "
    "Menu : **Sales → Waste** (calendrier) → **New**."
)

WASTE_FLOW = [
    "Sales → Waste → New.",
    "Date : choisir le jour de la perte (calendrier).",
    "Type : **Item** (matière) ou **Prep** (batch préparé) — pas Product pour cet exercice.",
    "Sélectionner l'item/prep dans la liste → saisir **Quantity** + **UOM** + **Reason** (raison).",
    "Vérifier Unit Cost et Total → **Save** (une entrée = une ligne ; répéter New pour chaque perte).",
    "Le calendrier Waste affiche le jour avec une quantité / valeur.",
]

WASTE_TABLE_HEADERS = ["Date", "Jour", "Type", "Item / Prep", "Quantité", "Raison"]
WASTE_TABLE = [
    [EXO_DAY_TO_DATE["Mercredi"], "Mercredi", "Item", "Laitue romaine", "500 g", "Fanée"],
    [EXO_DAY_TO_DATE["Jeudi"], "Jeudi", "Prep", "Frites maison — batch (prep)", "400 g", "Brûlées"],
    [EXO_DAY_TO_DATE["Vendredi"], "Vendredi", "Prep", "Chili — batch (prep)", "800 g", "Fin de shelf life"],
    [EXO_DAY_TO_DATE["Samedi"], "Samedi", "Item", "Pain burger", "6 ea", "Sec / invendable"],
]

WASTE_VALIDATE = [
    "4 entrées Waste Save (mer. → sam.).",
    "Qty on Hand diminuée sur laitue, frites prep, chili prep, pain.",
    "Usage Summary (étape 38) : colonne Waste > 0 sur ces items.",
]

def waste_rows():
    return [list(row) for row in WASTE_TABLE]

# ---- Daily Sales (Phase F — étape 36, ch.10) ----
DAILY_SALES_NOTE = (
    "**Sales Mix** (Till Tape / import, étapes 32–33) = **quantités par product** → OC calcule l'usage idéal et "
    "préremplit souvent la colonne de gauche (Food / Beverage). "
    "**Daily Sales** (étape 36) = **photographie financière du jour** : totaux caisse, remises, dépôts, main-d'œuvre — "
    "comme le Z de caisse / rapport de fin de journée POS. Les deux sont complémentaires : sans Daily Sales, "
    "Usage Summary peut afficher « Sales not defined » et le **Labour %** ne se calcule pas."
)

DAILY_SALES_SCREEN_NOTE = (
    "Écran Daily Sales (3 colonnes) :\n"
    "• **Gauche** — Ventes par catégorie (Food, Beverage…) : souvent **auto** depuis le Sales Mix du jour.\n"
    "• **Milieu** — Date, taxes, **Net / Gross Sales**, **Variance**, remises (Comp, Promo…), **Sales Voids**, "
    "gestion cash (Cash In Drawer, Cash Payout, Cash Received).\n"
    "• **Droite** — **Dépôts** (cash, carte, gift card), **Labour** (horaire + management), "
    "**Customer Count**, **Ticket Average**, commentaires.\n"
    "Vous n'avez pas à tout remplir chaque jour : voir tableau « champs à saisir par jour »."
)

DAILY_SALES_GROSS_NOTE = (
    "**Gross Sales** : si le Sales Mix du jour est déjà Save, OC préremplit Food + Beverage et le total "
    "(ex. mardi ≈ 1 181,75 $US depuis le Till Tape). **Ne changez pas Gross** sauf scénario comptable spécial. "
    "Concentrez-vous sur les champs **opérationnels** du tableau ci-dessous (remises, dépôts, labour…). "
    "**Net Sales** se met à jour quand vous saisissez des remises ; sinon Net ≈ Gross."
)

DAILY_SALES_DISTRIB_NOTE = (
    "Exercice Comptoir : les montants opérationnels sont **répartis sur la semaine** pour vous faire pratiquer "
    "tous les champs sans surcharger chaque jour. Cellule **—** = laisser **0** ou vide. "
    "Après saisie → **Save** (disquette) → **Open** → jour suivant → **New**."
)

DAILY_SALES_FIELDS_HEADERS = ["Zone", "Champ OC", "Rôle", "Exercice Comptoir"]
DAILY_SALES_FIELDS = [
    [
        "Gauche — Catégories",
        "Food / Beverage / … (Net & Gross par ligne)",
        "Ventilation des ventes par Sales Category (liée aux Sales Groups des Products).",
        "Auto depuis Sales Mix — vérifier seulement que Food + Beverage > 0.",
    ],
    [
        "Milieu — Totaux",
        "Gross Sales",
        "Total brut des ventes du jour (avant ou selon config remises).",
        "Prérempli depuis Sales Mix — ne pas modifier sauf consigne formateur.",
    ],
    [
        "Milieu — Totaux",
        "Net Sales",
        "Ventes nettes après remises / ajustements.",
        "Se met à jour si remises saisies ; sinon = Gross.",
    ],
    [
        "Milieu — Totaux",
        "Variance",
        "Écart entre Daily Sales saisi et total Sales Mix.",
        "0 ou faible si Gross vient du mix ; augmente si vous modifiez Gross manuellement.",
    ],
    [
        "Milieu — Cash",
        "Cash In Drawer",
        "Fond de caisse / argent compté en tiroir en fin de journée.",
        "Saisir **samedi 04/07** seulement (320 $US) — simulation comptage tiroir.",
    ],
    [
        "Milieu — Cash",
        "Cash Payout",
        "Sorties de caisse (petites dépenses payées cash depuis le tiroir).",
        "Saisir **jeudi 02/07** : 38 $US (courses urgentes).",
    ],
    [
        "Milieu — Cash",
        "Cash Received",
        "Entrées cash hors ventes (rare).",
        "Laisser 0 dans l'exercice.",
    ],
    [
        "Milieu — Remises",
        "Comp",
        "Repas offerts / gratuits (staff, erreur service).",
        "Mar. 12 $US · Sam. 8 $US.",
    ],
    [
        "Milieu — Remises",
        "Promo",
        "Promotions marketing (rabais affiché).",
        "Mer. 22,50 $US · Sam. 14 $US.",
    ],
    [
        "Milieu — Remises",
        "Employee Discount",
        "Rabais employés.",
        "Ven. 18 $US.",
    ],
    [
        "Milieu — Remises",
        "Other Discounts",
        "Autres rabais non classés.",
        "Jeu. 8 $US.",
    ],
    [
        "Milieu — Remises",
        "Total Discounts",
        "Somme des remises (calculé).",
        "Vérifie après saisie Comp / Promo / etc.",
    ],
    [
        "Milieu",
        "Sales Voids",
        "Ventes annulées après coup (ticket void au POS).",
        "Jeu. 9,95 $US (1 bol chili annulé).",
    ],
    [
        "Droite — Dépôts",
        "Cash Deposit",
        "Dépôt bancaire espèces du jour.",
        "Mer. 280 $US · Sam. 320 $US.",
    ],
    [
        "Droite — Dépôts",
        "Credit Card Deposits",
        "Encaissements carte déposés / batch terminal.",
        "Mer. 980 $US · Sam. 1 350 $US.",
    ],
    [
        "Droite — Dépôts",
        "Gift Card Deposits",
        "Ventes ou encaissements cartes cadeau.",
        "Ven. 45 $US.",
    ],
    [
        "Droite — Dépôts",
        "Total Deposits",
        "Somme des dépôts (calculé).",
        "Contrôle : dépôts ≈ partie des ventes non laissées en tiroir.",
    ],
    [
        "Droite",
        "Overage +/-",
        "Écart de caisse (tiroir vs attendu) — surplus ou manque.",
        "Sam. +4,25 $US (léger surplus).",
    ],
    [
        "Droite — Labour",
        "Hourly Labour",
        "Coût main-d'œuvre horaire du jour (saisie manuelle ou import paie).",
        "Mer. 175 · Ven. 195 · Sam. 240 $US.",
    ],
    [
        "Droite — Labour",
        "Management Labour",
        "Salaire / coût management attribué au jour.",
        "Ven. et Sam. : 115 $US chacun.",
    ],
    [
        "Droite — Labour",
        "Total Labour / Labour %",
        "Somme MO et % sur ventes du jour.",
        "Se calcule après Hourly + Management ; utile au Usage Summary / KPI.",
    ],
    [
        "Droite — Stats",
        "Customer Count",
        "Nombre de clients / couverts / tickets.",
        "Mar. 82 · Mer. 90 · Jeu. 76 · Ven. 98 · Sam. 112.",
    ],
    [
        "Droite — Stats",
        "Ticket Average",
        "Panier moyen (souvent Gross ÷ clients).",
        "Auto si OC calcule ; sinon laisser 0.",
    ],
    [
        "Droite",
        "Comments",
        "Notes libres sur la journée.",
        "Mar. « Ouverture semaine » · Jeu. « Void chili » · Sam. « Samedi marché ».",
    ],
]

DAILY_SALES_FLOW = [
    "Sales → Daily Sales → New (ou double-clic sur un jour vide du calendrier).",
    "Vérifier **Sales Date** (mar. 30/06 → sam. 04/07).",
    "Colonne **gauche** : Food + Beverage préremplis depuis Sales Mix — ne pas retaper les produits.",
    "Colonne **milieu** : vérifier Gross / Net ; saisir uniquement les champs indiqués pour CE jour (tableau ops).",
    "Colonne **droite** : dépôts, labour, customer count, commentaire selon le jour.",
    "Save → Open → jour suivant. Répéter **5 jours**.",
]

# Gross = total Sales Mix connu (Till Tape + imports) — champs ops répartis sur la semaine
DAILY_SALES_OPS_HEADERS = [
    "Date", "Jour", "Gross (auto)", "Comp", "Promo", "Emp. disc.", "Other disc.",
    "Voids", "Cash payout", "Cash drawer", "Overage", "Cash dep.", "CC dep.",
    "Gift dep.", "Hourly $", "Mgmt $", "Clients", "Commentaire",
]
DAILY_SALES_OPS_BY_DAY = [
    [EXO_DAY_TO_DATE["Mardi"], "Mar.", "1 181,75", "12,00", "—", "—", "—", "—", "—", "—", "—", "—", "—", "—", "—", "—", "82", "Ouverture semaine"],
    [EXO_DAY_TO_DATE["Mercredi"], "Mer.", "1 316,10", "—", "22,50", "—", "—", "—", "—", "—", "—", "280,00", "980,00", "—", "175,00", "—", "90", "—"],
    [EXO_DAY_TO_DATE["Jeudi"], "Jeu.", "1 130,50", "—", "—", "—", "8,00", "9,95", "38,00", "—", "—", "—", "—", "—", "—", "—", "76", "Void bol chili"],
    [EXO_DAY_TO_DATE["Vendredi"], "Ven.", "1 415,45", "—", "—", "18,00", "—", "—", "—", "—", "—", "—", "—", "45,00", "195,00", "115,00", "98", "—"],
    [EXO_DAY_TO_DATE["Samedi"], "Sam.", "1 717,50", "8,00", "14,00", "—", "—", "—", "—", "320,00", "+4,25", "320,00", "1 350,00", "—", "240,00", "115,00", "112", "Samedi marché"],
]

DAILY_SALES_TABLE_HEADERS = ["Date", "Jour", "Gross ($US, Sales Mix)", "Champs ops ce jour (résumé)"]
DAILY_SALES_TABLE = [
    [EXO_DAY_TO_DATE["Mardi"], "Mardi", "1 181,75", "Comp 12 · 82 clients"],
    [EXO_DAY_TO_DATE["Mercredi"], "Mercredi", "1 316,10", "Promo 22,50 · dépôts cash/CC · MO 175 · 90 clients"],
    [EXO_DAY_TO_DATE["Jeudi"], "Jeudi", "1 130,50", "Other disc. 8 · Void 9,95 · Cash payout 38 · 76 clients"],
    [EXO_DAY_TO_DATE["Vendredi"], "Vendredi", "1 415,45", "Emp. disc. 18 · Gift card 45 · MO 195+115 · 98 clients"],
    [EXO_DAY_TO_DATE["Samedi"], "Samedi", "1 717,50", "Comp+Promo · tiroir · overage · dépôts · MO 240+115 · 112 clients"],
]

DAILY_SALES_VALIDATE = [
    "5 jours Daily Sales Save (mar. → sam.).",
    "Gross = prérempli depuis Sales Mix (ne pas forcer 1 420 $ simulation).",
    "Au moins 3 types de champs ops pratiqués sur la semaine (remise, dépôt, labour, void…).",
    "Customer Count saisi au moins 3 jours.",
    "Usage Summary (étape 38) : Gross Sales période > 0, pas « Sales not defined ».",
]

def daily_sales_field_rows():
    return [list(row) for row in DAILY_SALES_FIELDS]

def daily_sales_ops_rows():
    return [list(row) for row in DAILY_SALES_OPS_BY_DAY]

def daily_sales_rows():
    return [list(row) for row in DAILY_SALES_TABLE]

def daily_sales_book_blocks():
    """Paragraphes Word — Daily Sales (rebuild_final.py, ch.10)."""
    blocks = [
        ("Heading2", "Exercice Comptoir — Daily Sales détaillé (étape 36)"),
        ("Normal", DAILY_SALES_NOTE),
        ("Normal", DAILY_SALES_SCREEN_NOTE),
        ("Normal", DAILY_SALES_GROSS_NOTE),
        ("Normal", DAILY_SALES_DISTRIB_NOTE),
        ("Heading3", "Lexique — chaque champ de l'écran"),
    ]
    for zone, champ, role, exo in DAILY_SALES_FIELDS:
        blocks.append(("Normal", f"• **{champ}** ({zone}) — {role} {exo}"))
    blocks.append(("Heading3", "Valeurs à saisir par jour (répartition sur la semaine)"))
    blocks.append((
        "Normal",
        "Gross = auto depuis Sales Mix. Colonne — = laisser 0. CC = Credit Card.",
    ))
    for row in DAILY_SALES_OPS_BY_DAY:
        date, jour = row[0], row[1]
        gross = row[2]
        filled = [f"{h}={v}" for h, v in zip(DAILY_SALES_OPS_HEADERS[3:], row[3:]) if v != "—"]
        blocks.append(("Normal", f"• {jour} {date} — Gross {gross} $US · " + (" · ".join(filled) if filled else "vérif. auto seulement")))
    blocks.append(("Heading3", "Ordre de saisie"))
    for step in DAILY_SALES_FLOW:
        blocks.append(("Normal", f"• {step}"))
    return blocks

# ---- Inventaire de clôture (Phase G — étape 37, vidéo #25) ----
CLOSING_INVENTORY_NOTE = (
    "Deuxième inventaire physique : **samedi matin 04/07/2026**, après une semaine d'achats, ventes et pertes. "
    "OC compare **ouverture (29/06) + achats − clôture (04/07)** = **usage actual** ; "
    "les ventes donnent l'**usage ideal**. L'écart (variance) apparaît dans Usage Summary. "
    "Même logique que l'ouverture : Count Inventory → New → date → comptage → Summarize → **Finalize/Save**."
)

CLOSING_INVENTORY_FLOW = [
    f"Count Inventory → New → date **{EXO_DATE_CLOSING}** (samedi matin).",
    "Tri **Location** (comme à l'ouverture) — compter shelf-to-sheet par emplacement.",
    "Saisir les quantités **terrain** ci-dessous pour les items clés ; compléter le reste de façon cohérente.",
    "Summarize Count → vérifier valorisation et rappel factures/ventes de la période.",
    "Finalize / Save — verrouille la période pour Usage Summary (29/06 → 04/07).",
]

CLOSING_INVENTORY_QTY_HEADERS = ["Item / Prep clé", "Quantité comptée", "UOM", "Emplacement type"]
CLOSING_INVENTORY_QTY = [
    ["Pain burger", "8", "ea", "Sec"],
    ["Bœuf haché 80/20", "2500", "g", "Congélateur / Ligne chaude"],
    ["Pommes de terre", "8000", "g", "Sec"],
    ["Huile friture", "9000", "ml", "Friterie"],
    ["Chili — batch (prep)", "1500", "g", "Ligne chaude"],
    ["Cola 355 ml", "22", "ea", "Bar"],
    ["Boîte burger", "120", "ea", "Sec"],
]

CLOSING_INVENTORY_VALIDATE = [
    f"Inventaire clôture Save / Finalize — date {EXO_DATE_CLOSING}.",
    "Période couverte : ouverture 29/06 → clôture 04/07.",
    "Summarize sans erreur bloquante.",
]

def closing_inventory_qty_rows():
    return [list(row) for row in CLOSING_INVENTORY_QTY]

# ---- Usage Summary (Phase G — étape 38, vidéo #48) ----
USAGE_SUMMARY_NOTE = (
    "Rapport central : **actual vs ideal** par item. "
    "**Actual** = inventaire ouverture + achats période − inventaire clôture. "
    "**Ideal** = ingrédients consommés selon les ventes (recettes × Sales Mix). "
    "**Waste** = pertes saisies. **Net variance** = actual − ideal (hors waste selon config). "
    "Prérequis : ventes Valid, factures, waste, **deux inventaires** Finalize."
)

USAGE_SUMMARY_FORMULA = (
    "Formule mentale : **Actual usage** ≈ stock début + factures − stock fin. "
    "Si Actual et Ideal divergent beaucoup sur la laitue → vérifier waste mercredi + yield 75 % sur la prep. "
    "Huile friture : écart possible si **Actualize** activé (item non recetté au gramme près)."
)

USAGE_SUMMARY_FLOW = [
    "Reports → Usage Summary.",
    f"Période / dates d'inventaire : **{EXO_DATE_OPENING}** (ouverture) → **{EXO_DATE_CLOSING}** (clôture).",
    "Type : **Count Amounts** ou **Values** (vidéo #48).",
    "Run Report — en-tête : Gross Sales, variance globale, food cost %.",
    "Double-clic **Period Purchases** → vos factures BP, DC, BN, EH.",
    "Double-clic **Ideal Usage** → quels products ont consommé l'item.",
    "Double-clic **Waste** → détail des 4 pertes.",
]

USAGE_SUMMARY_QUESTIONS = [
    "Quel item a le plus grand écart ($US) Actual vs Ideal ?",
    "Quel item a un ideal à zéro et un actual positif — pourquoi ?",
    "Le food cost % global Food est-il proche de 32 % ?",
    "La colonne Waste montre-t-elle laitue, frites, chili, pain ?",
    "Pouvez-vous expliquer un écart en 3 phrases à un propriétaire ?",
]

USAGE_SUMMARY_VALIDATE = [
    "Rapport généré sans « Sales not defined » sur Food.",
    "Drill-down Purchases et Ideal testés sur au moins 2 items.",
    "Variance laitue / chili cohérente avec waste saisi.",
]

# ---- Backup (Phase G — étape 39) ----
BACKUP_NOTE = (
    "Sauvegardez la base après Usage Summary : **File → Backup Data**. "
    "Conservez le fichier .bak — en production, backup hebdomadaire après inventaire."
)

BACKUP_FLOW = [
    "Recipe → Recipe Book → Export PDF (au moins Burgers + Chili).",
    "File → Backup Data → choisir dossier → Save.",
    "Optionnel : Reports → Cost of Sales Analysis, Item Activity (bœuf).",
]

def sales_till_tape_book_blocks():
    """Paragraphes Word — Till Tape / Phase F (rebuild_final.py)."""
    blocks = [
        ("Heading2", "Exercice Comptoir — Till Tape (Phase F, vidéo #43)"),
        ("Normal", TILL_TAPE_NOTE),
        ("Normal", TILL_TAPE_TWO_STEP_NOTE),
        ("Normal", TILL_TAPE_POS_NOTE),
        ("Heading3", "Étape A — Créer la liste Till Tape (une fois)"),
    ]
    for step in TILL_TAPE_LIST_FLOW:
        blocks.append(("Normal", f"• {step}"))
    blocks.append(("Heading3", f"Étape B — Saisir le mardi {EXO_DAY_TO_DATE[EXO_FIRST_SALES_DAY]}"))
    for step in TILL_TAPE_SALE_FLOW:
        blocks.append(("Normal", f"• {step}"))
    blocks.append(("Heading3", f"Quantités mardi {EXO_DAY_TO_DATE[EXO_FIRST_SALES_DAY]}"))
    for prod, pos, qty in TILL_TAPE_MARDI_QTY:
        blocks.append(("Normal", f"• {prod} (POS {pos}) — {qty}"))
    blocks.append(("Normal", DAILY_SALES_GROSS_NOTE))
    blocks.append(("Normal", CSV_IMPORT_NOTE))
    return blocks

# Inventory → Items → New — tous les champs (vidéo #06 + manuel OC v5)
COUNTSHEET_TRACK_NOTE = (
    "OC v5 : Track Inventory n'est plus sur la fiche Item. "
    "C'est dans Count Inventory → Countsheet Setup (vidéo #22, étape 25) : "
    "colonnes Track Inventory (suivre le stock / qty on hand) et Should Count (sur la feuille de comptage). "
    "Par défaut Track = coché pour les items actifs."
)

OPENING_INVENTORY_NOTE = (
    "L'inventaire d'ouverture est la première « photo » du stock dans Optimum Control. "
    "Vous la saisissez manuellement sur la feuille de comptage (count sheet) : OC ne devine pas vos quantités "
    "à partir des factures. C'est normal et attendu à ce stade : vous n'avez souvent encore aucune facture entrée, "
    "et les Qty on Hand sur les fiches Item restent à 0 tant que vous n'avez pas cliqué Finalize sur l'inventaire "
    "d'ouverture (ou reçu des achats). L'assistant Create Inventory (New) s'ouvre quand même — il crée seulement "
    "la session de comptage ; c'est vous qui entrez les quantités ligne par ligne. "
    "Dans cet exercice, l'ouverture est volontairement avant les factures (Partie 6) : vous fixez le stock de départ "
    "dimanche soir, puis les achats de la semaine s'y ajoutent."
)

OPENING_INVENTORY_FLOW = [
    "Count Inventory → Countsheet Setup (étape 25) : Track Inventory + Should Count.",
    "Count Inventory → Customize Sort (optionnel, étape 26) : ordre shelf-to-sheet.",
    "Count Inventory → New → assistant Create Inventory : date, All Items, Finish.",
    "Feuille générée : tri Location (ou Custom) → saisir les quantités (colonnes purchase / case / pack selon l'item).",
    "Summarize Count → vérifier totaux et valorisation → Finalize (verrouille le stock de départ).",
    "Ensuite seulement : Purchasing → Invoices (Partie 6 / étape 29).",
]

OPENING_INVENTORY_WIZARD_HEADERS = ["Champ Create Inventory", "Rôle", "Exercice Comptoir"]
OPENING_INVENTORY_WIZARD = [
    (
        "Date",
        "Stock enregistré à la fin de ce jour.",
        "Comptage lundi 29/06 soir → date = 29/06/2026 (inventaire d'ouverture).",
    ),
    (
        "Multiple count sheets required?",
        "Une feuille par filtre si plusieurs Groupes ou Hot Lists en parallèle.",
        "Inventaire complet : laisser décoché.",
    ),
    (
        "What kind of inventory count?",
        "All Items = tout ce qui est sur Countsheet Setup ; autres options = comptage partiel.",
        "All Items → Finish.",
    ),
]

def opening_inventory_wizard_rows():
    return [list(row) for row in OPENING_INVENTORY_WIZARD]

OPENING_INVENTORY_COUNT_NOTE = (
    "Sur la feuille de comptage : saisir les Items dans **Purchase Count** en **cs** (caisses — unité d'achat), "
    "sauf exception indiquée (Pak Count). Les Preps inclus au comptage se saisissent en **batch** dans Purchase Count. "
    "L'en-tête sous chaque ligne rappelle le case size (ex. « 12 ea / cs », « 50 lb / cs »). "
    "Portion frites 200 g n'est pas sur la feuille — on compte le prep Frites maison — batch."
)

OPENING_INVENTORY_QTY_HEADERS = [
    "Item / Prep", "Emplacement", "Colonne", "Saisie", "Uom comptage", "≈ stock (reporting)",
]
# Colonne = Purchase Count sauf si Pak Count indiqué. Preps = batch dans Purchase Count.
OPENING_INVENTORY_QTY = [
    # Bar
    ["Cola 355 ml", "Bar", "Purchase Count", "2", "cs", "48 each"],
    ["Eau 500 ml", "Bar", "Purchase Count", "2", "cs", "48 each"],
    # Congélateur
    ["Bœuf haché 80/20", "Congélateur", "Purchase Count", "2", "cs", "10 000 g"],
    ["Bacon tranché", "Congélateur", "Purchase Count", "4", "cs", "≈ 4 000 g (8 sacs × 500 g)"],
    ["Bœuf haché chili", "Congélateur", "Purchase Count", "2", "cs", "10 000 g"],
    ["Boulette 80 g (prep)", "Congélateur", "Purchase Count", "30", "batch", "30 each"],
    # Friterie
    ["Huile friture", "Friterie", "Purchase Count", "1", "cs", "16 000 ml"],
    ["Frites maison — batch (prep)", "Friterie", "Purchase Count", "0", "batch", "0 g — batch mardi"],
    # Ligne chaude
    ["Chili — batch (prep)", "Ligne chaude", "Purchase Count", "0", "batch", "0 g — batch mardi"],
    # Réfrigérateur
    ["Laitue romaine", "Réfrigérateur", "Purchase Count", "2", "cs", "≈ 6 000 g brut"],
    ["Parmesan râpé", "Réfrigérateur", "Purchase Count", "1", "cs", "2 000 g"],
    ["César dressing", "Réfrigérateur", "Purchase Count", "1", "cs", "≈ 15 120 ml"],
    ["Tomates", "Réfrigérateur", "Purchase Count", "1", "cs", "≈ 11 340 g"],
    ["Oignons", "Réfrigérateur", "Purchase Count", "1", "cs", "≈ 11 340 g"],
    ["Mayonnaise", "Réfrigérateur", "Purchase Count", "1", "cs", "≈ 15 120 ml"],
    ["Sauce burger maison (prep)", "Réfrigérateur", "Purchase Count", "2", "batch", "≈ 1 900 ml"],
    # Sec
    ["Pain burger", "Sec", "Purchase Count", "2", "cs", "24 each"],
    ["Pommes de terre", "Sec", "Purchase Count", "1", "cs", "≈ 22 680 g (50 lb / cs)"],
    ["Croûtons", "Sec", "Purchase Count", "2", "cs", "2 000 g"],
    ["Haricots rouges secs", "Sec", "Purchase Count", "1", "cs", "20 000 g"],
    ["Ketchup", "Sec", "Purchase Count", "1", "cs", "≈ 20 226 ml"],
    ["Sel fin", "Sec", "Purchase Count", "1", "cs", "2 000 g"],
    ["Boîte burger", "Sec", "Pak Count", "200", "each", "200 each"],
    ["Serviette", "Sec", "Purchase Count", "1", "cs", "5 000 each"],
    ["Fourchette plastique", "Sec", "Purchase Count", "2", "cs", "2 000 each"],
]

def opening_inventory_qty_rows():
    return [list(row) for row in OPENING_INVENTORY_QTY]

OPENING_INVENTORY_VIEW_NOTE = (
    "OC v5 — ne confondez pas la feuille de comptage (ce que vous avez saisi à une date) et le Qty on Hand "
    "(stock courant dans le système). Après Finalize, les Items affichent Qty on Hand sur Inventory → Items ; "
    "la fiche Recipe → Preps n'affiche pas ce champ — utilisez Reports → Prep Sheet (colonne On Hand) ou "
    "Count Inventory → Summarize pour les preps."
)

OPENING_INVENTORY_VIEW_HEADERS = ["Besoin", "Menu OC", "Exercice Comptoir"]
OPENING_INVENTORY_VIEW_DOC = [
    (
        "Revoir l'inventaire saisi",
        "Inventory → Count Inventory → calendrier → cliquer la date finalisée.",
        "Feuille par Location ; Boulette 30 (Congélateur), Sauce 2 batch (Réfrigérateur). Adjust Count Sheet pour corriger.",
    ),
    (
        "Summarize / valorisation",
        "Même fenêtre → Summarize Count (en haut).",
        "Liste surtout les Items avec qty et valeur. Preps parfois absents du Summary — voir ligne ci-dessous.",
    ),
    (
        "Qty on Hand — Items",
        "Inventory → Items → ouvrir l'item → panneau Info (bas).",
        "Vérifier Pain burger = 24 each, Pommes = 1 cs comptée, etc. Track Inventory = Countsheet Setup (pas ici).",
    ),
    (
        "Qty on Hand — Preps",
        "Recipe → Preps → fiche Boulette / Sauce : pas de Qty on Hand affiché en v5.",
        "Reports → Prep Sheet Daily (ou Weekly) → colonne On Hand. Si 0 alors que la feuille montre 30 / 2 : Adjust Count Sheet → Re-summarize, ou produire le prep en semaine.",
    ),
]

def opening_inventory_view_rows():
    return [list(row) for row in OPENING_INVENTORY_VIEW_DOC]

def opening_inventory_book_blocks():
    """Paragraphes Word — chapitre 6 inventaire (rebuild_final.py)."""
    blocks = [
        ("Heading2", "Inventaire d'ouverture (Opening Inventory)"),
        ("Normal", OPENING_INVENTORY_NOTE),
        ("Heading3", "Pourquoi avant les factures ?"),
        (
            "Normal",
            "OC distingue deux façons d'alimenter le stock : (1) vous déclarez un inventaire physique "
            "(ouverture ou clôture) ; (2) les factures ajoutent des achats. À la mise en service, il n'y a "
            "pas encore d'historique d'achats — l'ouverture pose la baseline. Formule de la semaine : "
            "stock fin = ouverture + achats − usage théorique (ventes, waste) ; l'inventaire de clôture "
            "compare le physique à ce calcul (Usage Summary).",
        ),
        ("Heading3", "Assistant Create Inventory (Count Inventory → New)"),
        ("Normal", "• Date — fin de journée retenue pour le stock."),
        ("Normal", "• Multiple count sheets — cocher seulement pour plusieurs Hot Lists / Groupes simultanés."),
        ("Normal", "• Type — All Items pour un inventaire complet ; Key / Category / Group / Hot List = partiel."),
        ("Heading3", "Feuille de comptage : c'est ici que le stock entre dans OC"),
        ("Normal", OPENING_INVENTORY_COUNT_NOTE),
        (
            "Normal",
            "Après Finish, choisissez le tri Location (recommandé). Saisissez les quantités dans la colonne "
            "indiquée du tableau exercice : en général Purchase Count en cs (items) ou batch (preps). "
            "OC additionne purchase / case / pack selon les unités compatibles. Aucune facture requise.",
        ),
        ("Heading3", "Summarize Count et Finalize"),
        (
            "Normal",
            "Summarize Count affiche surtout les Items avec quantité et valorisation. Finalize verrouille l'inventaire "
            "et met à jour les Qty on Hand des Items. Sans Finalize, le stock officiel n'est pas enregistré. "
            "Erreur après coup : Adjust Count Sheet → modifier → Re-summarize → Save (motif demandé si préférence activée).",
        ),
        ("Heading3", "Où voir l'inventaire et Qty on Hand (OC v5)"),
        ("Normal", OPENING_INVENTORY_VIEW_NOTE),
    ]
    for need, menu, exo in OPENING_INVENTORY_VIEW_DOC:
        blocks.append(("Normal", f"• {need} — {menu} {exo}"))
    blocks.extend([
        ("Heading3", "Exercice Comptoir — saisie lundi 29/06 soir (tri Location)"),
    ])
    for row in OPENING_INVENTORY_QTY:
        name, loc, col, qty, uom, approx = row
        blocks.append((
            "Normal",
            f"• {name} — {loc} : {col} = {qty} {uom} → {approx}",
        ))
    blocks.append((
        "Normal",
        "Chili et Frites batch à 0 batch = volontaire (production mardi). Finalize, puis factures (Partie 6).",
    ))
    return blocks

# ---- Factures manuelles (Phase E / Partie 6 — vidéo #37) ----
INVOICE_PHASE_NOTE = (
    "Phase E (étapes 29–31) = saisir les achats après l'inventaire d'ouverture Finalize (29/06/2026). "
    "Chaque facture augmente le Qty on Hand (en plus du stock déjà posé à l'ouverture), met à jour "
    "le Last Cost sur les case sizes et recalcule les coûts des preps/products. "
    "Ordre : Factures (29) → Commande (30) → Order Reminder (31) → puis ventes (Phase F). "
    + EXO_WEEK_NOTE
)

INVOICE_MANUAL_NOTE = (
    "Purchasing → Invoices → New (vidéo #37). Fenêtre en deux panneaux : à gauche le catalogue du "
    "fournisseur choisi ; à droite les lignes de la facture. Double-clic ou glisser-déposer pour ajouter "
    "un item. Prix = case size du fournisseur (Price Includes Tax = Non sur nos items). "
    "Onglet Taxes and Adjustments : TCA 10 % recoverable. Onglet Account Balance : doit être à **0** avant Save. "
    "Preferences → Purchasing → Show advanced tax adjustment fields = coché (étape 12)."
)

INVOICE_SCREEN_HEADERS = ["Zone / onglet", "Champ OC", "Rôle", "Exercice Comptoir"]
INVOICE_SCREEN_DOC = [
    ("En-tête", "Supplier", "Fournisseur — filtre le catalogue à gauche.", "Choisir avant d'ajouter des lignes."),
    ("En-tête", "Invoice Date", "Date comptable de l'achat / livraison.", "Voir colonne Date du tableau (ex. 30/06/2026 mardi)."),
    ("En-tête", "Invoice Number", "N° facture fournisseur.", "Ex. BP-2026-0142 — unique par facture."),
    ("En-tête", "Invoice Total", "Total TTC attendu (contrôle).", "Doit égaler sous-total + TCA après saisie."),
    ("Lignes (droite)", "Item + Case Size", "Une ligne par item commandé.", "Qté en caisses (cs) sauf Split Case."),
    ("Lignes", "Quantity", "Nombre de caisses achetées.", "2 cs pain = Quantity 2."),
    ("Lignes", "Unit Cost / Price", "Prix par caisse.", "Pré-rempli depuis case size ; Price Override si changement."),
    ("Lignes", "Split Case", "Fraction de caisse.", "Décoché pour cet exercice (caisses entières)."),
    ("Onglet", "Taxes and Adjustments", "Détail TCA.", "TCA 10 % sur sous-total HT — Tax Group TCA-ACHAT sur items."),
    ("Onglet", "Account Balance", "Équilibre comptable.", "Doit afficher 0 avant Save."),
    ("Après Save", "Price Variance", "Écart vs dernier achat.", "Normal au premier achat ; lire le popup."),
]

def invoice_screen_rows():
    return [list(row) for row in INVOICE_SCREEN_DOC]

INVOICE_ENTRY_FLOW = [
    "Purchasing → Invoices → New.",
    "Supplier → **Invoice Date** (date réelle du tableau) → Invoice Number.",
    "Gauche : chercher l'item → double-clic → ligne à droite.",
    "Quantity = nb de caisses ; vérifier Unit Cost = prix case size.",
    "Onglet Taxes and Adjustments → TCA 10 % (recoverable).",
    "Onglet Account Balance → 0 → Save.",
    "Répéter pour les 5 factures (dates 30/06, 02/07, 03/07, 04/07).",
]

# Lignes détaillées — prix/cs = Purchase Price fiche item ; Extension = qty × prix
INVOICE_LINE_HEADERS = [
    "Date", "Jour", "Fournisseur", "N° facture", "Item", "Qté", "Uom", "Prix/cs HT", "Extension HT",
]
INVOICE_LINES = [
    ["30/06/2026", "Mardi", NOM_BOULANGERIE, "BP-2026-0142", "Pain burger", "2", "cs", "2,40 $US", "4,80 $US"],
    ["30/06/2026", "Mardi", NOM_DISTRIB, "DC-2026-0318", "Bœuf haché 80/20", "2", "cs", "42,00 $US", "84,00 $US"],
    ["30/06/2026", "Mardi", NOM_DISTRIB, "DC-2026-0318", "Pommes de terre", "2", "cs", "22,00 $US", "44,00 $US"],
    ["30/06/2026", "Mardi", NOM_DISTRIB, "DC-2026-0318", "Laitue romaine", "1", "cs", "28,00 $US", "28,00 $US"],
    ["30/06/2026", "Mardi", NOM_DISTRIB, "DC-2026-0318", "Oignons", "1", "cs", "18,00 $US", "18,00 $US"],
    ["02/07/2026", "Jeudi", NOM_DISTRIB, "DC-2026-0320", "Bacon tranché", "2", "cs", "18,50 $US", "37,00 $US"],
    ["02/07/2026", "Jeudi", NOM_DISTRIB, "DC-2026-0320", "Parmesan râpé", "1", "cs", "24,00 $US", "24,00 $US"],
    ["02/07/2026", "Jeudi", NOM_DISTRIB, "DC-2026-0320", "Haricots rouges secs", "1", "cs", "55,00 $US", "55,00 $US"],
    ["02/07/2026", "Jeudi", NOM_DISTRIB, "DC-2026-0320", "Sel fin", "1", "cs", "3,50 $US", "3,50 $US"],
    ["02/07/2026", "Jeudi", NOM_DISTRIB, "DC-2026-0320", "Tomates", "1", "cs", "45,00 $US", "45,00 $US"],
    ["03/07/2026", "Vendredi", NOM_BOISSONS, "BN-2026-0088", "Cola 355 ml", "4", "cs", "12,00 $US", "48,00 $US"],
    ["04/07/2026", "Samedi", NOM_EMBALLAGES, "EH-2026-0205", "Boîte burger", "1", "cs", "45,00 $US", "45,00 $US"],
    ["04/07/2026", "Samedi", NOM_EMBALLAGES, "EH-2026-0205", "Serviette", "1", "cs", "35,00 $US", "35,00 $US"],
]

def invoice_line_rows():
    return [list(row) for row in INVOICE_LINES]

INVOICE_SUMMARY_HEADERS = [
    "Date", "Jour", "Fournisseur", "N° facture", "Sous-total HT", "TCA 10 %", "Total TTC",
]
INVOICE_SUMMARY = [
    ["30/06/2026", "Mardi", NOM_BOULANGERIE, "BP-2026-0142", "4,80 $US", "0,48 $US", "5,28 $US"],
    ["30/06/2026", "Mardi", NOM_DISTRIB, "DC-2026-0318", "174,00 $US", "17,40 $US", "191,40 $US"],
    ["02/07/2026", "Jeudi", NOM_DISTRIB, "DC-2026-0320", "164,50 $US", "16,45 $US", "180,95 $US"],
    ["03/07/2026", "Vendredi", NOM_BOISSONS, "BN-2026-0088", "48,00 $US", "4,80 $US", "52,80 $US"],
    ["04/07/2026", "Samedi", NOM_EMBALLAGES, "EH-2026-0205", "80,00 $US", "8,00 $US", "88,00 $US"],
]

def invoice_summary_rows():
    return [list(row) for row in INVOICE_SUMMARY]

INVOICE_LAB_BOULANGERIE = (
    "Exemple papier — Facture Boulangerie Pétion (1re facture à saisir dans OC)\n"
    "────────────────────────────────────────\n"
    "BOULANGERIE PÉTION · Port-au-Prince\n"
    "Facture n° BP-2026-0142 · Date livraison : **30/06/2026** (mardi)\n"
    "Client : Le Comptoir du Marché\n"
    "────────────────────────────────────────\n"
    "  Pain burger (12 pains / cs)    2 cs  ×  2,40 $US  =   4,80 $US\n"
    "────────────────────────────────────────\n"
    "  Sous-total HT                              4,80 $US\n"
    "  TCA recoverable 10 %                       0,48 $US\n"
    "  TOTAL TTC                                  5,28 $US\n"
    "────────────────────────────────────────\n"
    "Dans OC : Supplier = Boulangerie Pétion · Invoice Date = **30/06/2026** · Invoice Number = BP-2026-0142 · "
    "1 ligne Pain burger · Qty 2 · Taxes → TCA · Account Balance = 0 · Save."
)

INVOICE_LAB_DISTRIB_LUNDI = (
    "Facture Distrib. Caraïbes n° DC-2026-0318 · Date livraison : **30/06/2026** (mardi)\n"
    "  Bœuf haché 80/20 (5 kg/cs)     2 cs × 42,00 =  84,00 $US\n"
    "  Pommes de terre (50 lb/cs)     2 cs × 22,00 =  44,00 $US\n"
    "  Laitue romaine (12 têtes/cs)  1 cs × 28,00 =  28,00 $US\n"
    "  Oignons (25 lb/cs)            1 cs × 18,00 =  18,00 $US\n"
    "  Sous-total HT 174,00 $US · TCA 17,40 $US · Total TTC 191,40 $US"
)

ORDER_MARDI_HEADERS = ["Date", "Item", "Fournisseur", "Qté commande", "Uom", "Note"]
ORDER_MARDI = [
    ["30/06/2026", "Bœuf haché 80/20", NOM_DISTRIB, "1", "cs", "Réappro — par level ou manuel"],
    ["30/06/2026", "Pommes de terre", NOM_DISTRIB, "2", "cs", "Round to Case si Order Reminder"],
]

def order_mardi_rows():
    return [list(row) for row in ORDER_MARDI]

# Alias legacy
ORDER_LUNDI_HEADERS = ORDER_MARDI_HEADERS
order_lundi_rows = order_mardi_rows

def invoice_book_blocks():
    """Paragraphes Word — achats / factures (rebuild_final.py)."""
    blocks = [
        ("Heading2", "Exercice Comptoir — factures manuelles (Phase E)"),
        ("Normal", INVOICE_PHASE_NOTE),
        ("Normal", INVOICE_MANUAL_NOTE),
        ("Heading3", "Écran Invoices → New"),
    ]
    for zone, field, role, exo in INVOICE_SCREEN_DOC:
        blocks.append(("Normal", f"• {field} ({zone}) — {role} {exo}"))
    blocks.append(("Heading3", "Calendrier de la semaine"))
    blocks.append(("Normal", EXO_WEEK_NOTE))
    for day, date, event in EXO_WEEK:
        blocks.append(("Normal", f"• {day} {date} — {event}"))
    blocks.append(("Heading3", "Facture labo — Boulangerie (mardi 30/06/2026)"))
    blocks.append(("Normal", INVOICE_LAB_BOULANGERIE))
    blocks.append(("Heading3", "Récapitulatif — 5 factures de la semaine"))
    for row in INVOICE_SUMMARY:
        date, day, sup, num, ht, tax, ttc = row
        blocks.append(("Normal", f"• {date} ({day}) — {sup} {num} : HT {ht} + TCA {tax} = TTC {ttc}"))
    return blocks

ITEM_CREATE_NOTE = (
    "Inventory → Items → New. Panneau gauche = Core Information ; panneau droit = Case Size. "
    "Onglets après sauvegarde : Locations, Conversions, Nutrition (opt.), Allergens (opt.). "
    "Un seul item par produit — plusieurs fournisseurs = bouton Add (case size), pas un nouvel item. "
    "Track Inventory : voir Countsheet Setup (pas sur cet écran en v5)."
)

ITEM_CREATE_HEADERS = ["Zone", "Champ OC", "Rôle", "Exercice Comptoir"]
ITEM_CREATE_DOC = [
    ("Core", "Description", "Nom affiché partout (rapports, recettes, commandes).", "Ex. « Bœuf haché 80/20 » — nom simple, pas le libellé fournisseur."),
    ("Core", "Inventory Group", "Groupe d'inventaire (Setup → Inventory Groups).", "Proteins, Dry Goods, Produce, Paper, Soft Drinks… — pas Category."),
    ("Core", "Main Location", "Emplacement principal (commande + comptage par défaut).", "Sec, Congélateur, Réfrigérateur, Friterie, Bar…"),
    ("Core", "How it is used?", "Verrouille le type d'unité en recette : Weight, Volume ou Unit.", "Poids : bœuf, légumes. Volume : huile, ketchup. Unit : pain, cola."),
    ("Core", "Reporting Unit", "Unité des rapports (achats, stock, usage).", "gramme, ml, chacun — alimente le Split Unit par défaut."),
    ("Core", "Default Ingredient Unit", "Unité proposée en premier dans les recettes.", "Même famille que Reporting ; modifiable dans la recette."),
    ("Core", "Key Inventory Item", "Si coché : prioritaire sur les feuilles de comptage (Key Items).", "ON par défaut ; OFF sur sel / serviettes si vous allégez le comptage."),
    ("Core", "Actualize Usage Values", "Actual = idéal dans Usage Summary (pas de variance).", "ON : huile friture (bain — écarts inévitables). OFF : bœuf, bacon, laitue (variances utiles)."),
    ("Countsheet Setup", "Track Inventory", "OC v5 — pas sur fiche Item. Count Inventory → Countsheet Setup.", "Coché = suivi qty on hand + eligible comptage (défaut : oui)."),
    ("Countsheet Setup", "Should Count", "Même écran — par case size.", "Default case size (icône panier) toujours sur feuille ; cocher autres tailles si besoin."),
    ("Case Size", "Supplier", "Fournisseur de cet achat.", "Distrib. Caraïbes, Boulangerie Pétion, etc."),
    ("Case Size", "Order Code", "Code article fournisseur (import EDI).", "Vide OK pour l'exercice."),
    ("Case Size", "Case Description", "Libellé caisse sur facture fournisseur.", "Laisser vide sauf import EDI."),
    ("Case Size", "Barcode", "Scan inventaire (OC Mobile).", "Vide pour l'exercice."),
    ("Case Size", "Tax Group", "Groupe Setup → Tax Groups ; lie l'item à une ou plusieurs taxes.", f"{TAX_GROUP_CODE} — Price Includes Tax = Non."),
    ("Case Size", "Price Includes Tax", "Coché si prix fournisseur TTC (bière, alcool, dépôts).", "Décoché — TCA sur facture."),
    ("Case Size", "Purchase Price", "Montant payé au fournisseur.", "Ex. 42,00 $US — puis choisir l'unité du prix dans la liste."),
    ("Case Size", "Prix pour (liste après le montant)", "Unité à laquelle le prix s'applique : case, bag, each, gallon, dozen, lb, kg, L…", "Presque toujours case ; lb/kg si catch weight."),
    ("Case Size", "Split Unit + qty", "Contenu de l'unité d'achat (sacs, kg, lb, bidons…).", "Pain : 12 each dans 1 case. Bœuf : 5 kg dans 1 case."),
    ("Case Size", "Pack Unit + qty", "Unités recette dans 1 split unit.", "Pain : 1 each / each (normal). Bœuf : 1000 g / kg."),
    ("Case Size", "Yield %", "Rendement utilisable (pertes épluchage, etc.).", "Laitue 75 %, bacon 95 %, pommes 85 %, défaut 100 %."),
    ("Case Size", "Actual Cost by Recipe Unit", "Coût calculé à l'unité recette (lecture seule).", "Vérifier : bœuf ≈ 0,0084 $US/g."),
    ("Case Size", "Allow Split Case?", "Autoriser commande/achat en fraction de caisse.", "Décoché sauf besoin."),
    ("Case Size", "Default Ordering Case Size", "Caisse par défaut si plusieurs case sizes.", "Après ajout 2e fournisseur."),
    ("Case Size", "Current Case Size", "Dernière caisse achetée (après factures).", "Auto — ne pas saisir à la création."),
    ("Info", "Qty on Hand", "Stock actuel (unité reporting).", "0 à la création ; après Finalize inventaire d'ouverture ou factures."),
    ("Info", "Current Value", "Valeur du stock.", "Calculé."),
    ("Info", "Last Purchased", "Date du dernier achat.", "Après 1re facture."),
    ("Info", "Comments/Notes", "Notes libres.", "Optionnel."),
    ("Onglet Locations", "Main + Secondary", "Emplacements de comptage (épingle = principal).", "Ketchup/mayo/bacon : + Ligne chaude en secondaire."),
    ("Onglet Conversions", "Add conversion", "Conversion propre à cet item (ex. lb → tranches).", "Bacon : tranches ; voir vidéos #08–10 si besoin."),
    ("Onglet Nutrition", "Link to USDA / label", "Analyse nutritionnelle (module optionnel).", "Ignorer pour l'exercice."),
    ("Onglet Allergens", "Allergènes", "Allergènes pour cartes recette / rapports.", "Ignorer sauf si module actif."),
    ("Barre d'outils", "Add", "Ajouter une 2e case size (autre fournisseur/format).", "Même item, pas de doublon Description."),
    ("Barre d'outils", "View All", "Liste toutes les case sizes de l'item.", "Après plusieurs fournisseurs."),
    ("Barre d'outils", "Save", "Enregistrer l'item.", "Obligatoire avant onglets Locations/Conversions."),
]

def item_create_rows():
    return [list(row) for row in ITEM_CREATE_DOC]
