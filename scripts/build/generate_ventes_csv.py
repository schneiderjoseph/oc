#!/usr/bin/env python3
"""Génère les fichiers CSV de ventes pour l'exercice Comptoir du Marché."""
from pathlib import Path
import csv
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from exercice_locale import exo_sales_iso_by_day

OUT_DIR = Path(r"E:\OC DOCS\exercice_comptoir\ventes_csv")

PRODUCTS = [
    ("101", "Burger classique", "14.95"),
    ("102", "Burger bacon", "16.95"),
    ("201", "Frites moyennes", "5.50"),
    ("301", "Salade César", "13.50"),
    ("401", "Bol chili", "9.95"),
    ("501", "Cola 355 ml", "3.25"),
    ("502", "Eau 500 ml", "2.75"),
]

# Clés = jour calendaire réel 2026 (mar → sam)
SALES_QTY = {
    "mardi": {"101": 25, "102": 10, "201": 30, "301": 8, "401": 15, "501": 40, "502": 12},
    "mercredi": {"101": 28, "102": 12, "201": 35, "301": 10, "401": 18, "501": 45, "502": 15},
    "jeudi": {"101": 22, "102": 8, "201": 28, "301": 12, "401": 20, "501": 38, "502": 10},
    "vendredi": {"101": 30, "102": 15, "201": 40, "301": 9, "401": 16, "501": 50, "502": 18},
    "samedi": {"101": 35, "102": 18, "201": 45, "301": 14, "401": 22, "501": 55, "502": 20},
}

CSV_FIELDS = ["Date", "POS ID", "POS Description", "Quantity", "Gross Sales"]


def line_gross(qty: int, unit_price: str) -> str:
    return f"{qty * float(unit_price):.2f}"


def sale_row(date: str, pid: str, qty: int, price_map: dict, desc_map: dict) -> dict:
    return {
        "Date": date,
        "POS ID": pid,
        "POS Description": desc_map[pid],
        "Quantity": qty,
        "Gross Sales": line_gross(qty, price_map[pid]),
    }


def write_csv(path, rows, fieldnames):
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    price_map = {pid: price for pid, _, price in PRODUCTS}
    desc_map = {pid: desc for pid, desc, _ in PRODUCTS}
    iso_by_day = exo_sales_iso_by_day()

    # Supprimer anciens fichiers (mauvais jour/date ou noms legacy)
    stale_patterns = (
        "ventes_lundi_*.csv",
        "ventes_mardi_2026-07-*.csv",
        "ventes_mercredi_2026-07-02.csv",
        "ventes_jeudi_2026-07-03.csv",
        "ventes_vendredi_2026-07-04.csv",
        "ventes_minimal_lundi.csv",
        "*.occ",
    )
    for pattern in stale_patterns:
        for old in OUT_DIR.glob(pattern):
            old.unlink()

    all_rows = []
    for day, day_sales in SALES_QTY.items():
        date = iso_by_day[day]
        for pid, qty in day_sales.items():
            if qty:
                all_rows.append(sale_row(date, pid, qty, price_map, desc_map))
    write_csv(OUT_DIR / "ventes_semaine_comptoir.csv", all_rows, CSV_FIELDS)

    for day, day_sales in SALES_QTY.items():
        date = iso_by_day[day]
        rows = [sale_row(date, pid, qty, price_map, desc_map) for pid, qty in day_sales.items() if qty]
        write_csv(OUT_DIR / f"ventes_{day}_{date}.csv", rows, CSV_FIELDS)

    for day, day_sales in SALES_QTY.items():
        rows = [{"POS ID": pid, "Quantity": qty} for pid, qty in day_sales.items() if qty]
        write_csv(OUT_DIR / f"ventes_minimal_{day}.csv", rows, ["POS ID", "Quantity"])

    till_lines = [
        "# Référence Till Tape — Sales → New → Till Tape",
        "# 1re vente = MARDI 30/06/2026 (voir ventes_mardi_2026-06-30.csv pour import)",
        "",
    ]
    for day, qty_map in SALES_QTY.items():
        date = iso_by_day[day]
        till_lines.append(f"## {day.capitalize()} ({date})")
        for pid, _, _ in PRODUCTS:
            till_lines.append(f"POS {pid}: {qty_map.get(pid, 0)}")
        till_lines.append("")
    (OUT_DIR / "reference_till_tape.txt").write_text("\n".join(till_lines), encoding="utf-8")

    print(f"OK -> {OUT_DIR} ({len(list(OUT_DIR.glob('*.csv')))} fichiers CSV)")


if __name__ == "__main__":
    main()
