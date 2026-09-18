"""Add 'Sample Invoices' and 'Sample POs' after 'Queries' with a first-cut,
deterministic, materiality-based selection per GSTIN (B2B only; B2C and
advances excluded, per the CA).  Rule per GSTIN:
  - top 3 invoices by taxable value           -> 'Top-3 by value'
  - 2 mid-band invoices (nearest 50th / 25th pct) -> 'Mid-band'
  - 1 smallest positive invoice               -> 'Smallest'
  - top 2 credit notes by absolute value, plus any CN >= Rs 1 crore -> 'Large credit note'
Every line of a selected document is marked, so filtering the column shows whole documents.
"""
import os, collections, warnings; warnings.filterwarnings("ignore")
import openpyxl
from copy import copy

P = r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
R0, R1 = 5, 27006
def s(x): return "" if x is None else str(x).strip()

wb = openpyxl.load_workbook(P)
ws = wb["SR_2025-26"]
hdr = {s(ws.cell(4, c).value): c for c in range(1, ws.max_column + 1)}
cG, cF, cH, cI, cR, cQ = hdr["My GSTIN"], hdr["Document Number"], hdr["Document Type Code"], hdr["Supply Type Code"], hdr["Taxable Value"], hdr["Queries"]

# document-level taxable per (gstin, docno, type)
docs = collections.defaultdict(float); rows_of = collections.defaultdict(list)
for i in range(R0, R1 + 1):
    g, d, t, sup = s(ws.cell(i, cG).value), s(ws.cell(i, cF).value), s(ws.cell(i, cH).value), s(ws.cell(i, cI).value)
    if t.startswith("MOB") or (t == "INV" and sup == "B2C"):
        continue
    v = ws.cell(i, cR).value or 0
    docs[(g, d, t)] += float(v); rows_of[(g, d, t)].append(i)

picked = {}
by_g = collections.defaultdict(list)
for k, v in docs.items(): by_g[k[0]].append((k, v))
for g, lst in by_g.items():
    inv = sorted([(k, v) for k, v in lst if k[2] == "INV" and v > 0], key=lambda x: -x[1])
    crn = sorted([(k, v) for k, v in lst if k[2] == "CRN"], key=lambda x: abs(x[1]), reverse=True)
    for k, v in inv[:3]: picked.setdefault(k, "Top-3 by value")
    if len(inv) > 5:
        asc = sorted(inv, key=lambda x: x[1])
        for q in (0.50, 0.25):
            k, v = asc[int(len(asc) * q)]
            picked.setdefault(k, "Mid-band")
        picked.setdefault(asc[0][0], "Smallest")
    for k, v in crn[:2]: picked.setdefault(k, "Large credit note")
    for k, v in crn[2:]:
        if abs(v) >= 1e7: picked.setdefault(k, "Large credit note")

# insert the two columns after Queries
ws.insert_cols(cQ + 1, 2)
for off, name in ((1, "Sample Invoices"), (2, "Sample POs")):
    h = ws.cell(4, cQ + off, name); src = ws.cell(4, cQ)
    h.font = copy(src.font); h.fill = copy(src.fill); h.alignment = copy(src.alignment); h.border = copy(src.border)
    ws.column_dimensions[openpyxl.utils.get_column_letter(cQ + off)].width = 20
n_rows = 0
for k, reason in picked.items():
    for i in rows_of[k]:
        ws.cell(i, cQ + 1).value = reason
        ws.cell(i, cQ + 2).value = ("Pull PO / contract" if k[2] == "INV" else "Pull original invoice + CN approval")
        n_rows += 1
wb.save(P)
per_state = collections.Counter(k[0] for k in picked)
print("documents selected:", len(picked), "| register rows marked:", n_rows)
print("by reason:", dict(collections.Counter(picked.values())))
print("per GSTIN (min/max):", min(per_state.values()), "/", max(per_state.values()), "| states:", len(per_state))
