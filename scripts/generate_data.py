"""
generate_data.py
----------------
Creates the synthetic (made-up) raw data used in this project.

All companies, suppliers, customers, prices and quantities are invented for
learning purposes. HS codes are illustrative only and must not be used for any
real customs or trade decision.

Some data quality problems are added ON PURPOSE so the workbook can show how
they are detected (see "KNOWN ISSUES" comments below).

Usage:
    python scripts/generate_data.py            # writes CSV files to data/raw/
"""

from __future__ import annotations

import csv
import random
from datetime import date, timedelta
from pathlib import Path

SEED = 2026  # fixed seed so the output is the same every time
OUT_DIR = Path(__file__).resolve().parent.parent / "data" / "raw"

# ---------------------------------------------------------------------------
# 1. Finished products (items)
#    conversion_cost_cad = estimated labour + plant overhead per unit (assumption)
# ---------------------------------------------------------------------------
ITEMS = [
    # item_no, description, category, uom, hs_code, list_price_cad, conversion_cost_cad, concrete_m3
    ("PC-RCP-0600", "Reinforced concrete pipe 600 mm x 2.44 m", "Pipe", "EA", "6810.91", 640, 310, 0.38),
    ("PC-RCP-0900", "Reinforced concrete pipe 900 mm x 2.44 m", "Pipe", "EA", "6810.91", 1150, 540, 0.70),
    ("PC-RCP-1200", "Reinforced concrete pipe 1200 mm x 2.44 m", "Pipe", "EA", "6810.91", 1890, 860, 1.10),
    ("PC-BOX-1812", "Box culvert section 1.8 m x 1.2 m x 1.8 m", "Box Culvert", "EA", "6810.91", 4200, 1950, 2.60),
    ("PC-MH-1050B", "Manhole base 1050 mm", "Manhole", "EA", "6810.91", 980, 470, 0.75),
    ("PC-MH-1200B", "Manhole base 1200 mm", "Manhole", "EA", "6810.91", 1240, 600, 0.95),
    ("PC-MH-1200R", "Manhole riser 1200 mm x 600 mm", "Manhole", "EA", "6810.91", 520, 300, 0.42),
    ("PC-CB-0600", "Catch basin 600 mm x 600 mm", "Catch Basin", "EA", "6810.91", 690, 170, 0.45),
    ("PC-BAR-F38", "F-shape traffic barrier 3.8 m", "Barrier", "EA", "6810.99", 1450, 650, 1.52),
    ("PC-BAR-SS60", "Single-slope traffic barrier 6.0 m", "Barrier", "EA", "6810.99", 2350, 980, 2.60),
    ("PC-HC-200", "Hollow-core slab 200 mm (per m2)", "Slab", "M2", "6810.91", 118, 52, 0.13),
    # KNOWN ISSUE: missing HS code
    ("PC-HC-250", "Hollow-core slab 250 mm (per m2)", "Slab", "M2", "", 136, 60, 0.16),
]

# ---------------------------------------------------------------------------
# 2. Suppliers  (KNOWN ISSUE: SUP-11 is used by a component but is not listed)
# ---------------------------------------------------------------------------
SUPPLIERS = [
    # supplier_id, supplier_name, country, currency, lead_time_days
    ("SUP-01", "Prairie Cement Co.", "CA", "CAD", 7),
    ("SUP-02", "Great Plains Ash Inc.", "US", "USD", 14),
    ("SUP-03", "Red Valley Aggregates", "CA", "CAD", 3),
    ("SUP-04", "Midwest Admix Solutions", "US", "USD", 10),
    ("SUP-05", "Northern Rebar Mills", "CA", "CAD", 10),
    ("SUP-06", "Anatolia Steel Trading", "TR", "USD", 60),
    ("SUP-07", "Lakeside Wire Products", "US", "USD", 14),
    ("SUP-08", "Hanbit Strand Co.", "KR", "USD", 55),
    ("SUP-09", "Sierra Rubber Works", "MX", "USD", 21),
    ("SUP-10", "Rhein Lifting Systems", "DE", "EUR", 45),
    ("SUP-12", "Westway Site Supply", "CA", "CAD", 5),
]

# ---------------------------------------------------------------------------
# 3. Purchased components (raw materials and parts)
# ---------------------------------------------------------------------------
COMPONENTS = [
    # comp_code, description, uom, unit_cost, currency, supplier_id, origin_country, hs_code
    ("CMT-GU", "Portland cement Type GU", "T", 235.00, "CAD", "SUP-01", "CA", "2523.29"),
    ("SCM-FA", "Fly ash Class F", "T", 72.00, "USD", "SUP-02", "US", "2621.90"),
    ("AGG-20", "Coarse aggregate 20 mm", "T", 28.00, "CAD", "SUP-03", "CA", "2517.10"),
    ("AGG-SND", "Concrete sand", "T", 24.00, "CAD", "SUP-03", "CA", "2505.90"),
    ("ADM-SP", "Superplasticizer admixture", "L", 3.10, "USD", "SUP-04", "US", "3824.40"),
    # KNOWN ISSUE: HS code written without the dot
    ("ADM-AE", "Air-entraining admixture", "L", 2.40, "USD", "SUP-04", "US", "382440"),
    ("RB-10M", "Rebar 10M", "KG", 1.38, "CAD", "SUP-05", "CA", "7214.20"),
    ("RB-15M", "Rebar 15M", "KG", 1.31, "CAD", "SUP-05", "CA", "7214.20"),
    ("RB-20M", "Rebar 20M", "KG", 0.93, "USD", "SUP-06", "TR", "7214.20"),
    ("WWM-152", "Welded wire mesh 152x152 MW25.8", "M2", 4.60, "USD", "SUP-07", "US", "7314.20"),
    ("STR-127", "Prestressing strand 12.7 mm", "KG", 1.50, "USD", "SUP-08", "KR", "7312.10"),
    ("GSK-0600", "Rubber pipe gasket 600 mm", "EA", 18.50, "USD", "SUP-09", "MX", "4016.93"),
    ("GSK-0900", "Rubber pipe gasket 900 mm", "EA", 26.00, "USD", "SUP-09", "MX", "4016.93"),
    # KNOWN ISSUE: country of origin missing
    ("GSK-1200", "Rubber pipe gasket 1200 mm", "EA", 34.00, "USD", "SUP-09", "", "4016.93"),
    ("LFT-ANC", "Lifting anchor 2.5 t", "EA", 4.60, "EUR", "SUP-10", "DE", "7326.90"),
    # KNOWN ISSUE: supplier SUP-11 does not exist in the supplier list
    ("CST-FRM", "Catch basin frame and grate casting", "EA", 210.00, "CAD", "SUP-11", "IN", "7325.10"),
    # KNOWN ISSUE: unit cost missing
    ("SEAL-BUT", "Butyl joint sealant", "M", None, "CAD", "SUP-12", "CA", "3214.10"),
    ("FRM-OIL", "Form release agent", "L", 1.15, "CAD", "SUP-12", "CA", "2710.19"),
    # KNOWN ISSUE: duplicate of RB-15M (lower case + trailing space, different cost)
    ("rb-15m ", "Rebar 15M (old price)", "KG", 1.29, "CAD", "SUP-05", "CA", "7214.20"),
]

# Concrete mix per 1 m3 (assumption, simplified)
MIX_PER_M3 = [
    ("CMT-GU", 0.350, "T"),
    ("SCM-FA", 0.070, "T"),
    ("AGG-20", 1.050, "T"),
    ("AGG-SND", 0.750, "T"),
    ("ADM-SP", 2.000, "L"),
    ("ADM-AE", 0.400, "L"),
]

# Non-concrete parts per item: (comp_code, qty, uom)
PARTS = {
    "PC-RCP-0600": [("RB-10M", 28, "KG"), ("GSK-0600", 1, "EA"), ("FRM-OIL", 0.3, "L")],
    # KNOWN ISSUE: rebar entered in pounds (LB) instead of KG
    "PC-RCP-0900": [("RB-15M", 121, "LB"), ("GSK-0900", 1, "EA"), ("FRM-OIL", 0.4, "L")],
    "PC-RCP-1200": [("RB-15M", 95, "KG"), ("GSK-1200", 1, "EA"), ("LFT-ANC", 2, "EA"), ("FRM-OIL", 0.5, "L")],
    "PC-BOX-1812": [("RB-20M", 260, "KG"), ("RB-15M", 80, "KG"), ("LFT-ANC", 4, "EA"),
                    ("SEAL-BUT", 8, "M"), ("FRM-OIL", 1.0, "L")],
    "PC-MH-1050B": [("RB-10M", 30, "KG"), ("LFT-ANC", 2, "EA"), ("FRM-OIL", 0.4, "L")],
    # KNOWN ISSUE: RB-12M is not a valid component code
    "PC-MH-1200B": [("RB-12M", 40, "KG"), ("LFT-ANC", 2, "EA"), ("FRM-OIL", 0.5, "L")],
    "PC-MH-1200R": [("WWM-152", 3.2, "M2"), ("SEAL-BUT", 3.8, "M"), ("FRM-OIL", 0.3, "L")],
    # lower-case code on purpose: shows why keys are normalised (TRIM + UPPER)
    "PC-CB-0600": [("RB-10M", 18, "KG"), ("cst-frm", 1, "EA"), ("FRM-OIL", 0.3, "L")],
    "PC-BAR-F38": [("RB-15M", 70, "KG"), ("LFT-ANC", 2, "EA"), ("FRM-OIL", 0.6, "L")],
    "PC-BAR-SS60": [("RB-15M", 120, "KG"), ("RB-20M", 40, "KG"), ("LFT-ANC", 2, "EA"),
                    ("SEAL-BUT", 2, "M"), ("FRM-OIL", 0.9, "L")],
    "PC-HC-200": [("STR-127", 3.8, "KG"), ("FRM-OIL", 0.05, "L")],
    "PC-HC-250": [("STR-127", 4.6, "KG"), ("FRM-OIL", 0.06, "L")],
}

CUSTOMERS = [
    # customer_id, ship_to_region, ship_to_country
    ("CUST-001", "AB", "CA"), ("CUST-002", "AB", "CA"), ("CUST-003", "SK", "CA"),
    ("CUST-004", "BC", "CA"), ("CUST-005", "MB", "CA"), ("CUST-006", "AB", "CA"),
    ("CUST-007", "MT", "US"), ("CUST-008", "ND", "US"), ("CUST-009", "SK", "CA"),
    ("CUST-010", "BC", "CA"),
]

# Typical order quantity range per category
QTY_RANGE = {
    "Pipe": (4, 30), "Box Culvert": (2, 10), "Manhole": (1, 8), "Catch Basin": (2, 12),
    "Barrier": (5, 40), "Slab": (40, 400),
}


def write_csv(name: str, header: list[str], rows: list[tuple]) -> None:
    path = OUT_DIR / name
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        for row in rows:
            writer.writerow(["" if v is None else v for v in row])
    print(f"wrote {path.relative_to(OUT_DIR.parent.parent)}  ({len(rows)} rows)")


def build_bom() -> list[tuple]:
    rows = []
    for item_no, *_rest in ITEMS:
        volume = _rest[-1]
        line = 10
        for comp, qty_per_m3, uom in MIX_PER_M3:
            rows.append((item_no, line, comp, round(qty_per_m3 * volume, 4), uom))
            line += 10
        for comp, qty, uom in PARTS[item_no]:
            rows.append((item_no, line, comp, qty, uom))
            line += 10
    return rows


def build_sales(rng: random.Random) -> list[tuple]:
    rows = []
    start = date(2025, 1, 6)
    item_lookup = {i[0]: i for i in ITEMS}
    invoice = 25000
    for _ in range(260):
        invoice += 1
        day = start + timedelta(days=rng.randint(0, 350))
        cust_id, region, country = rng.choice(CUSTOMERS)
        item_no = rng.choice(list(item_lookup))
        category, list_price = item_lookup[item_no][2], item_lookup[item_no][5]
        lo, hi = QTY_RANGE[category]
        qty = rng.randint(lo, hi)
        price = round(list_price * rng.uniform(0.90, 1.04), 2)
        rows.append((f"INV-{invoice}", day.isoformat(), cust_id, region, country, item_no, qty, price))

    # KNOWN ISSUES in sales
    rows[17] = rows[17][:7] + (None,)                    # missing unit price
    rows[88] = rows[88][:7] + (None,)                    # missing unit price
    r = rows[45]; rows[45] = r[:7] + (round(item_lookup[r[5]][5] * 0.62, 2),)   # deep discount
    r = rows[150]; rows[150] = r[:7] + (round(item_lookup[r[5]][5] * 0.70, 2),)  # deep discount
    r = rows[201]; rows[201] = r[:5] + ("PC-RCP-0750",) + r[6:]                 # unknown item
    rows.sort(key=lambda x: (x[1], x[0]))
    return rows


def main() -> None:
    rng = random.Random(SEED)
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    write_csv("items.csv",
              ["item_no", "description", "category", "uom", "hs_code",
               "list_price_cad", "conversion_cost_cad"],
              [i[:7] for i in ITEMS])
    write_csv("suppliers.csv",
              ["supplier_id", "supplier_name", "country", "currency", "lead_time_days"],
              SUPPLIERS)
    write_csv("components.csv",
              ["comp_code", "description", "uom", "unit_cost", "currency",
               "supplier_id", "origin_country", "hs_code"],
              COMPONENTS)
    write_csv("bom.csv",
              ["parent_item", "line_no", "comp_code", "qty_per", "uom"],
              build_bom())
    write_csv("sales_2025.csv",
              ["invoice_no", "invoice_date", "customer_id", "ship_to_region",
               "ship_to_country", "item_no", "qty", "unit_price_cad"],
              build_sales(rng))


if __name__ == "__main__":
    main()
