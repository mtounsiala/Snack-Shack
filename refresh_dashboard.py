"""Refresh the dashboard HTML from the Excel workbook.

Usage:  python refresh_dashboard.py "Financial_Document.xlsx"
Needs:  pip install openpyxl
Reads the values Excel last saved (so save the workbook first), then rewrites
the <DATA> block inside October_Business_Dashboard.html (same folder).
"""
import json, re, sys, datetime
from pathlib import Path
from openpyxl import load_workbook

XLSX = sys.argv[1] if len(sys.argv) > 1 else "Financial_Document.xlsx"
HTML = Path(__file__).with_name("October_Business_Dashboard.html")

CATEGORY = {  # new products fall back to "Other"
    "Mi goreng": "Noodles",
    "Fanta": "Drinks", "Sprite": "Drinks", "Pepsi 330": "Drinks", "Pepsi 500": "Drinks",
    "Cadbury Astro": "Choc", "Cadbury 5 stars": "Choc", "Cadbury White Biscuits": "Choc",
    "Kitkat": "Choc", "Snickers": "Choc", "Kitkat white": "Choc", "Cadbury chocolate biscuits": "Choc",
    "Peanuts": "Snacks", "Cheas Naks": "Snacks", "Lays yellow": "Snacks", "Fritos": "Snacks", "Diddle daddles": "Snacks",
    "Sour Jelly": "Sweets", "Jelly Jerseys": "Sweets", "Jelly Babies": "Sweets", "Fruit Jubes": "Sweets", "Jelly funky flavour": "Sweets",
    "Oreo Original": "Biscuits", "Oreo Golden": "Biscuits",
}
num = lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

wb = load_workbook(XLSX, data_only=True)
dash, rp, pm, bg = (wb[n] for n in ("Dashboard", "Revenues&Profit", "Product management", "Boys&Girls side"))

pm_rows = {r[0]: r for r in pm.iter_rows(min_row=2, values_only=True) if isinstance(r[0], str)}
bg_rows = {r[0]: r for r in bg.iter_rows(min_row=2, values_only=True) if isinstance(r[0], str)}

products = []
for r in rp.iter_rows(min_row=2, max_row=36, values_only=True):
    name = r[0]
    if not isinstance(name, str) or name.startswith("#") or not num(r[1]):
        continue                                   # skip blank / #VALUE! placeholder rows
    p, b = pm_rows.get(name), bg_rows.get(name)
    products.append(dict(
        name=name, qty=p[1], sold=p[2], left=p[3], daysLeft=num(p[6]),
        profit=round(r[5], 2), margin=round(r[6], 2),
        revenue=round(r[3] * r[1], 2),             # revenue per unit x quantity (NOT the 90-day projection)
        boysUnits=num(b[3]) or 0, girlsUnits=num(b[6]) or 0,
        category=CATEGORY.get(name, "Other"), tracked=bool(r[13]),
    ))

d = lambda c: dash[c].value
summary = dict(
    revenue=d("B3"), spent=d("D3"), profit=d("F3"), margin=d("H3"), variance=d("J3"),
    net=d("L3"), daily=d("N3"), units=d("P3"), days=d("R3"),
    projProfit=rp["L37"].value, projRevenue=rp["M37"].value, runRate=rp["K37"].value,
    updated=datetime.date.today().strftime("%d %b %Y"),
)

# sanity checks so a broken sheet never silently produces a wrong email
rev_sum = sum(p["revenue"] for p in products)
assert abs(rev_sum - summary["revenue"]) < 1, f"Product revenue {rev_sum} != sheet total {summary['revenue']}"
assert abs(sum(p["profit"] for p in products) - summary["profit"]) < 1, "Product profit != sheet total"
assert sum(p["qty"] for p in products) == summary["units"], "Units don't add up"

block = ("// <DATA>\nconst summary = " + json.dumps(summary) + ";\nconst products = [\n  "
         + ",\n  ".join(json.dumps(p) for p in products) + "\n];\n// </DATA>")
html = HTML.read_text(encoding="utf-8")
html = re.sub(r"// <DATA>.*?// </DATA>", lambda m: block, html, flags=re.S)
HTML.write_text(html, encoding="utf-8")
print(f"Updated {HTML.name}: {len(products)} products, revenue R{summary['revenue']:,.2f}, profit R{summary['profit']:,.2f}")
