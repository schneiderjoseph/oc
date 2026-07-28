#!/usr/bin/env python3
"""Compare exercice item list vs ocdata query result."""
from pathlib import Path

EXPECTED = [
    "Pain burger",
    "Bœuf haché 80/20",
    "Pommes de terre",
    "Bacon tranché",
    "Laitue romaine",
    "Parmesan râpé",
    "Croûtons",
    "César dressing",
    "Huile friture",
    "Sel fin",
    "Tomates",
    "Oignons",
    "Haricots rouges secs",
    "Bœuf haché chili",
    "Ketchup",
    "Mayonnaise",
    "Cola 355 ml",
    "Eau 500 ml",
    "Boîte burger",
    "Serviette",
    "Fourchette plastique",
]

ALIASES = {
    "boeuf 80/20": "Bœuf haché 80/20",
    "bœuf 80/20": "Bœuf haché 80/20",
    "boeuf hache 80/20": "Bœuf haché 80/20",
}


def norm(s: str) -> str:
    return (
        s.strip()
        .lower()
        .replace("œ", "oe")
        .replace("é", "e")
        .replace("è", "e")
        .replace("ê", "e")
        .replace("à", "a")
        .replace("ô", "o")
        .replace("û", "u")
        .replace("ï", "i")
        .replace("haché", "hach")
        .replace("hache", "hach")
    )


def parse_db(path: Path):
    rows = []
    in_section = False
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        if line in ("INVENTORY_ITEMS_TYPE_I", "ITEMS_WITH_CASESIZE"):
            in_section = True
            continue
        if in_section and line.startswith("ItemId|"):
            continue
        if in_section:
            if line.startswith("TOTALS") or line.startswith("ITEM_TYPES"):
                break
            if not line.strip():
                continue
            parts = line.split("|")
            if len(parts) >= 5:
                rows.append({
                    "id": parts[0].strip(),
                    "type": parts[1].strip(),
                    "name": parts[2].strip(),
                    "cs": int(float(parts[3])),
                    "active_cs": int(float(parts[4])),
                })
    return rows


def match_expected(db_name: str, exp_norm: dict) -> str | None:
    n = norm(db_name)
    if n in exp_norm:
        return exp_norm[n]
    if n in ALIASES:
        return ALIASES[n]
    for en, display in exp_norm.items():
        if en in n or n in en:
            return display
    return None


def main():
    db_path = Path(r"E:\OC DOCS\verify_items_db.txt")
    if not db_path.exists():
        print("MANQUE verify_items_db.txt — lancer verify_exo_items.ps1")
        return 1

    db_items = parse_db(db_path)
    exp_norm = {norm(n): n for n in EXPECTED}
    found = []
    missing = list(EXPECTED)

    for r in db_items:
        display = match_expected(r["name"], exp_norm)
        if display:
            found.append((display, r, r["name"]))
            if display in missing:
                missing.remove(display)

    extra = [r for r in db_items if not match_expected(r["name"], exp_norm)]

    out = Path(r"E:\OC DOCS\verify_items_report.txt")
    lines = [
        f"Exercice attendu : {len(EXPECTED)} items",
        f"En base (avec case size) : {len(db_items)} items",
        "",
        f"TROUVÉS ({len(found)}/{len(EXPECTED)}) :",
    ]
    for name, r, db_name in sorted(found, key=lambda x: x[0]):
        cs_note = f"{r['cs']} case size(s)"
        if r["cs"] > 1:
            cs_note += " (multi OK)"
        if r["active_cs"] == 0:
            cs_note += " ⚠ pas de CaseSizeCost actif"
        alias = f"  [DB: {db_name}]" if db_name != name else ""
        lines.append(f"  OK  {name}{alias}  [id={r['id']}, {cs_note}]")

    if missing:
        lines += ["", f"MANQUANTS ({len(missing)}) :"]
        for m in missing:
            lines.append(f"  --  {m}")

    if extra:
        lines += ["", f"EN BASE MAIS PAS DANS L'EXO ({len(extra)}) :"]
        for r in extra:
            lines.append(f"  ++  {r['name']}  [id={r['id']}, {r['cs']} cs]")

    for line in db_path.read_text(encoding="utf-8").splitlines():
        if line.startswith("items|"):
            lines += ["", "TOTAUX DB : " + line.replace("|", " ")]
            break

    lines += [
        "",
        "NOTE : Bacon = 2 case sizes (Distrib. + Boulangerie) = 22 lignes pour 21 items.",
    ]
    text = "\n".join(lines)
    out.write_text(text, encoding="utf-8")
    print(text)
    return 0 if not missing else 1


if __name__ == "__main__":
    raise SystemExit(main())
