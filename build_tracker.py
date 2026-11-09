"""Build Inventory Tracker.html from the Sage vItemQuantitiesbyBin export.

Usage:  python build_tracker.py [path\\to\\vItemQuantitiesbyBin.xlsx]
Writes "Inventory Tracker.html" next to this script, with the data embedded.
"""
import datetime
import json
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "vItemQuantitiesbyBin.xlsx")

wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
ws = wb["vItemQtybyBin"] if "vItemQtybyBin" in wb.sheetnames else wb.worksheets[0]
rows = ws.iter_rows(values_only=True)
header = [str(h).strip() if h is not None else "" for h in next(rows)]
col = {h: i for i, h in enumerate(header)}
need = ["WarehouseCode", "WarehouseName", "BinLocation", "ItemCode", "ItemCodeDesc",
        "QuantityOnHand", "QuantityCommitted", "Available", "LastTotalUnitCost"]
missing = [n for n in need if n not in col]
if missing:
    sys.exit(f"Missing columns in {src}: {missing}")

wh, bins, items, out = [], [], [], []
wi, bi, ii = {}, {}, {}


def s(v):
    return "" if v is None else str(v).strip()


def n(v):
    try:
        return round(float(v or 0), 4)
    except (TypeError, ValueError):
        return 0


for r in rows:
    code = s(r[col["ItemCode"]])
    if not code:
        continue
    w = s(r[col["WarehouseCode"]]).zfill(3)
    if w not in wi:
        wi[w] = len(wh)
        wh.append([w, s(r[col["WarehouseName"]])])
    elif not wh[wi[w]][1]:
        wh[wi[w]][1] = s(r[col["WarehouseName"]])
    b = s(r[col["BinLocation"]])
    if b not in bi:
        bi[b] = len(bins)
        bins.append(b)
    if code not in ii:
        ii[code] = len(items)
        items.append([code, s(r[col["ItemCodeDesc"]])])
    vals = [n(r[col[k]]) for k in ("QuantityOnHand", "QuantityCommitted", "Available", "LastTotalUnitCost")]
    out.append([wi[w], bi[b], ii[code]] + [int(v) if v == int(v) else v for v in vals])

mtime = datetime.datetime.fromtimestamp(os.path.getmtime(src))
data = {
    "source": os.path.basename(src),
    "asof": mtime.strftime("%b %d, %Y").replace(" 0", " "),
    "wh": wh, "bins": bins, "items": items, "rows": out,
}

with open(os.path.join(HERE, "template.html"), encoding="utf-8") as f:
    html = f.read()
payload = json.dumps(data, separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")
html = html.replace("/*__DATA__*/null", payload, 1)
dest = os.path.join(HERE, "Inventory Tracker.html")
with open(dest, "w", encoding="utf-8") as f:
    f.write(html)
print(f"Wrote {dest}: {len(out):,} bin records, {len(items):,} parts, {len(wh)} warehouses, "
      f"{os.path.getsize(dest) / 1e6:.1f} MB")
