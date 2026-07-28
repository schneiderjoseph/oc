#!/usr/bin/env python3
"""Calcule les coûts unitaires et coûts recette pour le corrigé."""
from dataclasses import dataclass

@dataclass
class Item:
    name: str
    case_cost: float
    case_qty: float  # in recipe uom
    uom: str
    yield_pct: float = 100.0

    @property
    def unit_cost(self):
        return (self.case_cost / self.case_qty) * (100.0 / self.yield_pct)


ITEMS = {
    "pain": Item("Pain burger", 2.40, 12, "chacun"),
    "boeuf": Item("Bœuf 80/20", 42.00, 5000, "g"),
    "bacon": Item("Bacon", 18.50, 66, "tranche", 95),  # ~66 tranches estimées
    "mayo": Item("Mayonnaise", 48.00, 4 * 3780, "ml"),
    "ketchup": Item("Ketchup", 38.00, 6 * 114 * 29.5735, "ml"),  # oz -> ml
    "pommes": Item("Pommes de terre", 22.00, 50 * 453.592, "g", 85),
    "huile": Item("Huile friture", 38.00, 16000, "ml"),
    "sel": Item("Sel", 3.50, 2000, "g"),
    "laitue": Item("Laitue", 28.00, 12 * 500, "g", 75),  # ~500g/head
    "parmesan": Item("Parmesan", 24.00, 2000, "g"),
    "croûtons": Item("Croûtons", 8.00, 1000, "g"),
    "dressing": Item("César dressing", 52.00, 4 * 3780, "ml"),
    "chili_beef": Item("Bœuf chili", 40.00, 5000, "g"),
    "haricots": Item("Haricots secs", 55.00, 20000, "g"),
    "tomates": Item("Tomates", 45.00, 25 * 453.592, "g", 90),
    "oignons": Item("Oignons", 18.00, 25 * 453.592, "g", 88),
    "boite": Item("Boîte burger", 45.00, 500, "chacun"),
    "serviette": Item("Serviette", 35.00, 5000, "chacun"),
    "fourchette": Item("Fourchette", 12.00, 1000, "chacun"),
    "cola": Item("Cola", 12.00, 24, "chacun"),
    "eau": Item("Eau", 8.00, 24, "chacun"),
}


def prep_boulette():
    return 77 * ITEMS["boeuf"].unit_cost + 3 * ITEMS["sel"].unit_cost

def prep_frites_batch():
    # 5000g potato + 200ml oil + 50g salt -> 4000g yield
    cost = 5000 * ITEMS["pommes"].unit_cost + 200 * ITEMS["huile"].unit_cost + 50 * ITEMS["sel"].unit_cost
    return cost / 4000  # per gram output

def prep_frites_g():
    return prep_frites_batch()

def prep_sauce_ml():
    cost = 800 * ITEMS["mayo"].unit_cost + 200 * ITEMS["ketchup"].unit_cost
    return cost / 950

def prep_chili_g():
    cost = (
        2000 * ITEMS["haricots"].unit_cost
        + 3000 * ITEMS["chili_beef"].unit_cost
        + 1500 * ITEMS["tomates"].unit_cost
        + 500 * ITEMS["oignons"].unit_cost
        + 40 * ITEMS["sel"].unit_cost
    )
    return cost / 12000


def product_cost(name, parts):
    total = sum(qty * cost for qty, cost in parts)
    return total


PRODUCTS = {}

PRODUCTS["Burger classique"] = product_cost("Burger classique", [
    (1, ITEMS["pain"].unit_cost),
    (1, prep_boulette()),
    (15, prep_sauce_ml()),
    (30, ITEMS["tomates"].unit_cost),  # 30 g en recette (item en gram)
    (1, ITEMS["boite"].unit_cost),
    (1, ITEMS["serviette"].unit_cost),
])

PRODUCTS["Burger bacon"] = PRODUCTS["Burger classique"] + 2 * ITEMS["bacon"].unit_cost

PRODUCTS["Frites moyennes"] = product_cost("Frites", [
    (200, prep_frites_g()),
    (1, ITEMS["serviette"].unit_cost),
])

PRODUCTS["Salade César"] = product_cost("César", [
    (150, ITEMS["laitue"].unit_cost),
    (15, ITEMS["parmesan"].unit_cost),
    (20, ITEMS["croûtons"].unit_cost),
    (45, ITEMS["dressing"].unit_cost),
])

PRODUCTS["Bol chili"] = product_cost("Chili", [
    (400, prep_chili_g()),
    (1, ITEMS["fourchette"].unit_cost),
])

PRICES = {
    "Burger classique": 14.95,
    "Burger bacon": 16.95,
    "Frites moyennes": 5.50,
    "Salade César": 13.50,
    "Bol chili": 9.95,
    "Cola": 3.25,
    "Eau": 2.75,
}

PRODUCTS["Cola"] = ITEMS["cola"].unit_cost
PRODUCTS["Eau"] = ITEMS["eau"].unit_cost


def fmt_money(x):
    return f"{x:.4f} $".replace(".", ",") if x < 1 else f"{x:.2f} $"


def main():
    lines = ["=== COÛTS UNITAIRES ITEMS ==="]
    for k, it in ITEMS.items():
        lines.append(f"{it.name}: {it.unit_cost:.6f} $ / {it.uom}")

    lines.append("\n=== PREPS ===")
    lines.append(f"Boulette 80 g: {prep_boulette():.4f} $")
    lines.append(f"Frites (coût/g sortie): {prep_frites_g():.6f} $/g")
    lines.append(f"Sauce burger (coût/ml): {prep_sauce_ml():.6f} $/ml")
    lines.append(f"Chili (coût/g): {prep_chili_g():.6f} $/g")

    lines.append("\n=== PRODUCTS ===")
    for name, cost in PRODUCTS.items():
        price = PRICES.get(name, 0)
        pct = (cost / price * 100) if price else 0
        lines.append(f"{name}: coût {cost:.4f} $ | prix {price:.2f} $ | food cost {pct:.1f} %")

    # weekly sales totals
    sales = {
        "Lundi": {"101": 25, "102": 10, "201": 30, "301": 8, "401": 15, "501": 40, "502": 12},
        "Mardi": {"101": 28, "102": 12, "201": 35, "301": 10, "401": 18, "501": 45, "502": 15},
        "Mercredi": {"101": 22, "102": 8, "201": 28, "301": 12, "401": 20, "501": 38, "502": 10},
        "Jeudi": {"101": 30, "102": 15, "201": 40, "301": 9, "401": 16, "501": 50, "502": 18},
        "Vendredi": {"101": 35, "102": 18, "201": 45, "301": 14, "401": 22, "501": 55, "502": 20},
    }
    pos_map = {
        "101": "Burger classique", "102": "Burger bacon", "201": "Frites moyennes",
        "301": "Salade César", "401": "Bol chili", "501": "Cola", "502": "Eau",
    }
    lines.append("\n=== VENTES HEBDO (quantités) ===")
    totals = {k: 0 for k in pos_map}
    for day, d in sales.items():
        s = sum(d.values())
        for pid, q in d.items():
            totals[pid] += q
        lines.append(f"{day}: {s} transactions produit")
    lines.append("Totaux semaine:")
    for pid, q in totals.items():
        lines.append(f"  POS {pid} {pos_map[pid]}: {q}")

    open(r"E:\OC DOCS\corrigé_calc.txt", "w", encoding="utf-8").write("\n".join(lines))
    print("OK calculs")

if __name__ == "__main__":
    main()
