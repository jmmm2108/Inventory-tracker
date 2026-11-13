"""Build the inventory tracker (index.html) from the Sage vItemQuantitiesbyBin export.

Usage:  python build_tracker.py [path\\to\\vItemQuantitiesbyBin.xlsx]
Writes index.html (the tracker) next to this script, from template.html, with the data embedded.
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


# Warehouses with "Site" in the name (job-site / site-material warehouses) are left out of everything.
def is_site(name):
    return "site" in name.lower()


SITE_WH = {}  # code -> name of the warehouses left out
rows = list(rows)
for r in rows:
    if is_site(s(r[col["WarehouseName"]])):
        SITE_WH[s(r[col["WarehouseCode"]]).zfill(3)] = s(r[col["WarehouseName"]])

for r in rows:
    code = s(r[col["ItemCode"]])
    if not code:
        continue
    w = s(r[col["WarehouseCode"]]).zfill(3)
    if w in SITE_WH:
        continue
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


def asof(path):
    return datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%b %d, %Y").replace(" 0", " ")


# Item warehouse detail: one row per item + warehouse, non-zero rows only.
# Same order as the Sage Item Inquiry "Quantities" grid.
# Every part/warehouse in the detail is kept; one with no bin record gets a "No bin" record (bin "") so it shows
# in the By part tab, the totals and the kit check.
DET_COLS = ["QuantityOnHand", "QuantityOnPurchaseOrder", "QuantityOnSalesOrder", "QuantityOnBackOrder", "Committed",
            "QuantityOnWorkOrder", "QuantityRequiredForWO", "Available", "QuantityInShipping", "OnHandLessShipping"]
det_src = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "vItemWarehouseDetail.xlsx")
det = None
if os.path.exists(det_src):
    dws = openpyxl.load_workbook(det_src, read_only=True, data_only=True)
    dws = dws["vItemWarehouseDetail"] if "vItemWarehouseDetail" in dws.sheetnames else dws.worksheets[0]
    drows = dws.iter_rows(values_only=True)
    dh = [str(h).strip() if h is not None else "" for h in next(drows)]
    dc = {h: i for i, h in enumerate(dh)}
    dmiss = [c for c in ["ItemCode", "WarehouseCode"] + DET_COLS if c not in dc]
    if dmiss:
        sys.exit(f"Missing columns in {det_src}: {dmiss}")
    have = {(r[0], r[2]) for r in out}  # (warehouse idx, item idx) pairs with a bin record
    cost = {}
    for r in out:
        if r[6] and r[2] not in cost:
            cost[r[2]] = r[6]
    if "" not in bi:
        bi[""] = len(bins)
        bins.append("")
    det_rows, added, new_items, new_wh, zero = [], 0, 0, 0, {}
    for r in drows:
        code = s(r[dc["ItemCode"]])
        w = s(r[dc["WarehouseCode"]]).zfill(3)
        vals = [n(r[dc[k]]) for k in DET_COLS]
        if not code:
            continue
        wname = s(r[dc["WarehouseName"]]) if "WarehouseName" in dc else ""
        if w in SITE_WH or is_site(wname):
            SITE_WH.setdefault(w, wname)
            continue
        if not any(vals):  # all-zero row: keep the part listed (as a zero "No bin" record) only if nothing else has it
            zero.setdefault(code, (w, s(r[dc["WarehouseName"]]) if "WarehouseName" in dc else "",
                                   s(r[dc["ItemCodeDesc"]]) if "ItemCodeDesc" in dc else ""))
            continue
        if w not in wi:
            wi[w] = len(wh)
            wh.append([w, s(r[dc["WarehouseName"]]) if "WarehouseName" in dc else ""])
            new_wh += 1
        if code not in ii:
            ii[code] = len(items)
            items.append([code, s(r[dc["ItemCodeDesc"]]) if "ItemCodeDesc" in dc else ""])
            new_items += 1
        vals = [int(v) if v == int(v) else v for v in vals]
        det_rows.append([ii[code], wi[w]] + vals)
        if (wi[w], ii[code]) not in have:
            qoh, com, avail = vals[0], vals[4], vals[7]
            out.append([wi[w], bi[""], ii[code], qoh, com, avail, cost.get(ii[code], 0)])
            added += 1
    zero_added = 0
    for code, (w, wname, desc) in zero.items():
        if code in ii:
            continue
        if w not in wi:
            wi[w] = len(wh)
            wh.append([w, wname])
            new_wh += 1
        ii[code] = len(items)
        items.append([code, desc])
        out.append([wi[w], bi[""], ii[code], 0, 0, 0, 0])
        new_items += 1
        zero_added += 1
    print(f"Warehouse detail: {zero_added:,} parts with all-zero quantities listed so they can be found")
    det = {"source": os.path.basename(det_src), "asof": asof(det_src), "rows": det_rows}
    print(f"Warehouse detail: {len(det_rows):,} item/warehouse rows; added {added:,} 'No bin' records "
          f"({new_items:,} parts and {new_wh} warehouses not in the bin report)")

data = {
    "source": os.path.basename(src),
    "asof": asof(src),
    "wh": wh, "bins": bins, "items": items, "rows": out, "det": det,
}

with open(os.path.join(HERE, "template.html"), encoding="utf-8") as f:
    html = f.read()
payload = json.dumps(data, separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")
html = html.replace("/*__DATA__*/null", payload, 1)
dest = os.path.join(HERE, "index.html")
with open(dest, "w", encoding="utf-8") as f:
    f.write(html)
print(f"Left out {len(SITE_WH)} 'Site' warehouses: " + ", ".join(f"{c} {nm}" for c, nm in sorted(SITE_WH.items())))
print(f"Wrote {dest}: {len(out):,} bin records, {len(items):,} parts, {len(wh)} warehouses, "
      f"{os.path.getsize(dest) / 1e6:.1f} MB")
