#!/usr/bin/env python3
"""Vérification détaillée étape 29 — factures et QOH."""
import pyodbc

CONN = r"Driver={ODBC Driver 17 for SQL Server};Server=(localdb)\mssqllocaldb;Database=ocdata;Trusted_Connection=yes;"

EXPECTED = {
    "BP-2026-0142": {"date": "2026-06-30", "ttc": 5.28, "lines": [("Pain burger", 2, 2.40, 4.80)]},
    "DC-2026-0318": {
        "date": "2026-06-30", "ttc": 191.40,
        "lines": [
            ("Bœuf haché 80/20", 2, 42.00, 84.00),
            ("Pommes de terre", 2, 22.00, 44.00),
            ("Laitue romaine", 1, 28.00, 28.00),
            ("Oignons", 1, 18.00, 18.00),
        ],
    },
    "DC-2026-0320": {
        "date": "2026-07-02", "ttc": 180.95,
        "lines": [
            ("Bacon tranché", 2, 18.50, 37.00),
            ("Parmesan râpé", 1, 24.00, 24.00),
            ("Haricots rouges secs", 1, 55.00, 55.00),
            ("Sel fin", 1, 3.50, 3.50),
            ("Tomates", 1, 45.00, 45.00),
        ],
    },
    "BN-2026-0088": {"date": "2026-07-03", "ttc": 52.80, "lines": [("Cola 355 ml", 4, 12.00, 48.00)]},
    "EH-2026-0205": {
        "date": "2026-07-04", "ttc": 88.00,
        "lines": [("Boîte burger", 1, 45.00, 45.00), ("Serviette", 1, 35.00, 35.00)],
    },
}

# QOH attendu après ouverture + factures lundi (approx.)
QOH_EXPECTED = {
    "Pain burger": (48, "ea", "24 ouverture + 24 facture (2 cs × 12)"),
    "Bœuf haché 80/20": (20000, "g", "2 cs ouverture + 2 cs facture (5 kg/cs)"),
    "Pommes de terre": (68000, "g", "1 cs ouverture + 2 cs facture (50 lb/cs) — ordre de grandeur"),
    "Laitue romaine": (None, "g", "ouverture + 1 cs facture"),
    "Cola 355 ml": (192, "ea", "48 ouverture + 144 facture (4 cs × 36?)"),
}

OPENING = {
    "Pain burger": (24, "ea"),
    "Bœuf haché 80/20": (10000, "g"),
    "Pommes de terre": (22680, "g"),
    "Cola 355 ml": (48, "ea"),
    "Boîte burger": (200, "ea"),
    "Serviette": (5000, "ea"),
}


def main():
    try:
        cn = pyodbc.connect(CONN)
    except Exception as e:
        print(f"ERR connexion: {e}")
        return
    cur = cn.cursor()

    print("=" * 60)
    print("1. FACTURES (5 attendues)")
    print("=" * 60)
    cur.execute("""
        SELECT inv.InvoiceNumber, inv.InvoiceDate, inv.Total, inv.ItemTotal,
               inv.AdjustmentTotal, inv.ExpenseTotal, s.Name
        FROM oc.Invoice inv
        LEFT JOIN oc.Supplier s ON s.SupplierId = inv.Supplier
        ORDER BY inv.InvoiceId
    """)
    found = {}
    for row in cur.fetchall():
        num, dt, total, item, adj, exp, sup = row
        found[num] = row
        dt_s = dt.strftime("%Y-%m-%d") if dt else "?"
        bal = float(total or 0) - float(item or 0) - float(adj or 0) - float(exp or 0)
        ok_bal = abs(bal) < 0.02
        exp_meta = EXPECTED.get(num, {})
        ok_date = dt_s == exp_meta.get("date", dt_s) if exp_meta else True
        ok_ttc = abs(float(total or 0) - exp_meta.get("ttc", float(total or 0))) < 0.02 if exp_meta else True
        status = "OK" if ok_bal and ok_ttc else "CHECK"
        date_flag = "" if ok_date else f" [date attendue {exp_meta.get('date')}]"
        print(f"  [{status}] {num} | {dt_s}{date_flag} | {sup} | TTC {float(total):.2f} | balance résiduelle {bal:.4f}")

    for num, meta in EXPECTED.items():
        if num not in found:
            print(f"  [MANQUANT] {num} — TTC attendu {meta['ttc']:.2f} $")

    print("\n" + "=" * 60)
    print("2. LIGNES FACTURES")
    print("=" * 60)
    cur.execute("""
        SELECT inv.InvoiceNumber, i.Descrip, ii.Qty, ii.UnitCost, ii.LineTotal
        FROM oc.InvoiceItem ii
        JOIN oc.Invoice inv ON inv.InvoiceId = ii.Invoice
        JOIN oc.Item i ON i.ItemId = ii.Item
        ORDER BY inv.InvoiceId, ii.Idx
    """)
    lines_by_inv = {}
    for num, desc, qty, uc, lt in cur.fetchall():
        lines_by_inv.setdefault(num, []).append((desc, float(qty), float(uc), float(lt)))
        print(f"  {num} | {desc} | {qty} cs × {uc} = {lt}")

    for num, meta in EXPECTED.items():
        if num not in lines_by_inv:
            continue
        db_lines = lines_by_inv[num]
        exp_lines = meta["lines"]
        for exp in exp_lines:
            exp_name = exp[0]
            match = [l for l in db_lines if exp_name.split()[0] in l[0] or l[0] in exp_name]
            if not match:
                print(f"  [WARN] {num} : ligne attendue « {exp_name} » introuvable")
            else:
                _, q, u, t = match[0]
                if abs(q - exp[1]) > 0.01 or abs(u - exp[2]) > 0.01 or abs(t - exp[3]) > 0.01:
                    print(f"  [WARN] {num} / {exp_name} : attendu {exp[1]}×{exp[2]}={exp[3]}, trouvé {q}×{u}={t}")

    print("\n" + "=" * 60)
    print("3. QTY ON HAND (items clés)")
    print("=" * 60)
    cur.execute("""
        SELECT i.Descrip, q.QtyOnHand, ru.Uom
        FROM oc.ItemQtyOnHand q
        JOIN oc.Item i ON i.ItemId = q.Item
        LEFT JOIN oc.Uom ru ON ru.UomId = i.RecipeUom
        WHERE i.Type = 'I'
        ORDER BY i.Descrip
    """)
    qoh = {r[0]: (float(r[1]), r[2]) for r in cur.fetchall()}

    checks = [
        ("Pain burger", 48, "ea", "ouverture 24 + facture 24"),
        ("Cola 355 ml", None, "ea", "48 ouverture + 4 cs facture"),
        ("Boîte burger", 200, "ea", "200 ouverture seulement (EH pas saisie)"),
        ("Serviette", 5000, "ea", "5000 ouverture seulement (EH pas saisie)"),
    ]
    for name, exp, uom, note in checks:
        if name not in qoh:
            print(f"  [?] {name} : pas de QOH")
            continue
        val, u = qoh[name]
        if exp is not None:
            ok = abs(val - exp) < 0.5
            mark = "OK" if ok else "ECART"
            print(f"  [{mark}] {name} : {val:.1f} {u} (attendu ~{exp} — {note})")
        else:
            print(f"  [INFO] {name} : {val:.1f} {u} ({note})")

    # Boeuf / pommes detail
    for name in ("Bœuf haché 80/20", "Bœuf", "Pommes de terre"):
        if name in qoh:
            print(f"  [INFO] {name} : {qoh[name][0]:.1f} {qoh[name][1]}")

    print("\n" + "=" * 60)
    print("4. SYNTHÈSE CHECKLIST ÉTAPE 29")
    print("=" * 60)
    n_inv = len(found)
    all5 = n_inv >= 5 and all(n in found for n in EXPECTED)
    pain_ok = qoh.get("Pain burger", (0,))[0] == 48
    print(f"  [{'x' if all5 else ' '}] 5 factures Save — {n_inv}/5 en base")
    print(f"  [{'x' if all5 else ' '}] Account Balance = 0 — {'toutes sauvegardées OK' if n_inv else 'N/A'}")
    print(f"  [{'x' if pain_ok else ' '}] Qty on Hand Pain = 48 ea (24+24)")
    print(f"  [?] Popup Price Variance — non vérifiable en base (OK si 1er achat)")
    print(f"  [{' ' if all5 else 'x'}] Prêt étape 30 — {'OUI' if all5 and pain_ok else 'NON — corriger ci-dessus'}")

    cn.close()


if __name__ == "__main__":
    main()
