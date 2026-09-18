"""Sample Invoices / Sample POs, per the CA's on-screen method (session 2):
  B2B only (B2C + advances removed). Per GSTIN:
    - every invoice with taxable >= Rs 5 crore
    - if fewer than 3 such, top up to 3 by value
    - 2 small invoices (smallest positive, and ~25th percentile)
    - the 2 largest credit notes by absolute value, plus any CN with |taxable| >= Rs 1 crore
  Cell values: 'Sample Invoices' = the document number to pull (all lines of that document).
               'Sample POs'      = PO / contract reference from source data (filled by fill_po.py).
"""
import collections, warnings; warnings.filterwarnings("ignore")
import openpyxl
P = r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
R0, R1 = 5, 27006
def s(x): return "" if x is None else str(x).strip()
wb = openpyxl.load_workbook(P); ws = wb["SR_2025-26"]
hdr = {s(ws.cell(4, c).value): c for c in range(1, ws.max_column + 1)}
cG, cF, cH, cI, cR = hdr["My GSTIN"], hdr["Document Number"], hdr["Document Type Code"], hdr["Supply Type Code"], hdr["Taxable Value"]
cSI, cSP = hdr["Sample Invoices"], hdr["Sample POs"]
docs = collections.defaultdict(float); rows_of = collections.defaultdict(list)
for i in range(R0, R1 + 1):
    ws.cell(i, cSI).value = None; ws.cell(i, cSP).value = None      # clear the first-cut labels
    g, d, t, sup = s(ws.cell(i, cG).value), s(ws.cell(i, cF).value), s(ws.cell(i, cH).value), s(ws.cell(i, cI).value)
    if t.startswith("MOB") or (t == "INV" and sup == "B2C"):
        continue
    docs[(g, d, t)] += float(ws.cell(i, cR).value or 0); rows_of[(g, d, t)].append(i)
picked = set(); by_g = collections.defaultdict(list)
for k, v in docs.items(): by_g[k[0]].append((k, v))
for g, lst in by_g.items():
    inv = sorted([(k, v) for k, v in lst if k[2] == "INV" and v > 0], key=lambda x: -x[1])
    crn = sorted([(k, v) for k, v in lst if k[2] == "CRN"], key=lambda x: abs(x[1]), reverse=True)
    big = [k for k, v in inv if v >= 5e7]
    picked.update(big if len(big) >= 3 else [k for k, v in inv[:3]])
    if len(inv) > 4:
        asc = sorted(inv, key=lambda x: x[1])
        picked.add(asc[0][0]); picked.add(asc[len(asc) // 4][0])
    picked.update(k for k, v in crn[:2])
    picked.update(k for k, v in crn[2:] if abs(v) >= 1e7)
n = 0
for k in picked:
    for i in rows_of[k]:
        ws.cell(i, cSI).value = k[1]; n += 1
wb.save(P)
per = collections.Counter(k[0] for k in picked)
print("documents selected:", len(picked), "| invoices:", sum(1 for k in picked if k[2] == "INV"),
      "| credit notes:", sum(1 for k in picked if k[2] == "CRN"), "| rows marked:", n)
print("per GSTIN min/max:", min(per.values()), "/", max(per.values()))
print(">=5cr invoices selected:", sum(1 for k in picked if k[2] == "INV" and docs[k] >= 5e7),
      "| >=1cr credit notes:", sum(1 for k in picked if k[2] == "CRN" and abs(docs[k]) >= 1e7))
