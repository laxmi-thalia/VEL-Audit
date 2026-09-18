"""Sample Invoice / Sample PO (CA clarification, session 2 follow-up).

Cell values are DOCUMENT REFERENCES, never flags:
  Sample Invoices = the invoice / credit-note number selected for detailed testing
  Sample POs      = the related PO (SAP customer Contract No) and/or Sales Order, blank if none
Selection (B2B only; B2C + advances excluded), per GSTIN:
  materiality  : every invoice >= Rs 5 crore (min 3 per state by value)
  sampling     : smallest positive invoice + ~25th-percentile invoice
  unusual      : 2 largest credit notes + every credit note >= Rs 1 crore
  risk         : every document 'Not in GSTR-1'; up to 3 documents with an off-rate GST check
The basis of each pick is recorded on the 'Sample Selection' sheet.
"""
import os, collections, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill
SP = os.path.dirname(os.path.abspath(__file__))
src = open(os.path.join(SP, "sr_compare2.py"), encoding="utf-8").read().split("a,b=int")[0].replace("//server/", "//192.168.1.69/")
exec(src)   # BASE, SAPF, pick(), col()
def S(v): return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
def num(v):
    v = S(v)
    try: return str(int(float(v)))
    except Exception: return v

P = r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
R0, R1 = 5, 27006
# cached VALUES (rate check is a formula) from a data_only read
wv = openpyxl.load_workbook(P, read_only=True, data_only=True)["SR_2025-26"]
cached = {i: row for i, row in enumerate(wv.iter_rows(min_row=R0, max_row=R1, values_only=True), start=R0)}
wb = openpyxl.load_workbook(P); ws = wb["SR_2025-26"]
hdr = {S(ws.cell(4, c).value): c for c in range(1, ws.max_column + 1)}
cG, cB, cF, cH, cI, cR = hdr["My GSTIN"], hdr["My State"], hdr["Document Number"], hdr["Document Type Code"], hdr["Supply Type Code"], hdr["Taxable Value"]
cK, cRC, cMS, cSI, cSP = hdr["Recipient Legal Name"], hdr["GST rate check"], hdr["Matched with GSTR-1"], hdr["Sample Invoices"], hdr["Sample POs"]

# cached values for formula columns (rate check) come from the last Excel save
docs = collections.defaultdict(float); rows_of = collections.defaultdict(list); meta = {}
offrate = set(); notg1 = set()
for i in range(R0, R1 + 1):
    ws.cell(i, cSI).value = None; ws.cell(i, cSP).value = None
    g, d, t, sup = S(ws.cell(i, cG).value), S(ws.cell(i, cF).value), S(ws.cell(i, cH).value), S(ws.cell(i, cI).value)
    if t.startswith("MOB") or (t == "INV" and sup == "B2C"): continue
    k = (g, d, t)
    docs[k] += float(ws.cell(i, cR).value or 0); rows_of[k].append(i)
    meta.setdefault(k, (S(ws.cell(i, cB).value), S(ws.cell(i, cK).value)))
    if cached[i][cRC - 1] == "CHECK": offrate.add(k)
    if cached[i][cMS - 1] == "Not in GSTR-1": notg1.add(k)

basis = {}
def add(k, why): basis.setdefault(k, why)
by_g = collections.defaultdict(list)
for k, v in docs.items(): by_g[k[0]].append((k, v))
for g, lst in by_g.items():
    inv = sorted([(k, v) for k, v in lst if k[2] == "INV" and v > 0], key=lambda x: -x[1])
    crn = sorted([(k, v) for k, v in lst if k[2] == "CRN"], key=lambda x: abs(x[1]), reverse=True)
    big = [k for k, v in inv if v >= 5e7]
    for k in (big if len(big) >= 3 else [k for k, v in inv[:3]]): add(k, "Materiality: invoice >= Rs 5 crore" if docs[k] >= 5e7 else "Materiality: top-3 by value in state")
    if len(inv) > 4:
        asc = sorted(inv, key=lambda x: x[1])
        add(asc[0][0], "Sampling: smallest invoice in state"); add(asc[len(asc) // 4][0], "Sampling: small-value invoice")
    for k, v in crn[:2]: add(k, "Unusual: largest credit notes in state")
    for k, v in crn[2:]:
        if abs(v) >= 1e7: add(k, "Unusual: credit note >= Rs 1 crore")
    for k in sorted([k for k in offrate if k[0] == g], key=lambda k: -abs(docs[k]))[:3]: add(k, "Risk: GST rate charged differs from stated rate")
for k in notg1: add(k, "Risk: document not found in GSTR-1")

# PO / SO references from SAP, keyed by ODN
need = {k[1] for k in basis}
ref = collections.defaultdict(lambda: {"contract": set(), "so": set()})
for lbl, d, f in SAPF:
    got = pick(os.path.join(BASE, d, f))
    if not got: continue
    sh, hr, sap = got
    odn = col(sap, "ODN"); cn = col(sap, "Contract No"); so = col(sap, "Sales Doc Number")
    sub = sap[sap[odn].map(S).isin(need)]
    for _, r in sub.iterrows():
        k = S(r[odn])
        if cn is not None and S(r[cn]): ref[k]["contract"].add(num(r[cn]))
        if so is not None and S(r[so]): ref[k]["so"].add(num(r[so]))
def po_text(d):
    r = ref.get(d)
    if not r: return None
    parts = []
    if r["contract"]: parts.append("PO/Contract " + "/".join(sorted(r["contract"])))
    if r["so"]: parts.append("SO " + "/".join(sorted(r["so"])))
    return "; ".join(parts) or None

n = 0
for k in basis:
    po = po_text(k[1])
    for i in rows_of[k]:
        ws.cell(i, cSI).value = k[1]; ws.cell(i, cSP).value = po; n += 1

# Sample Selection sheet
if "Sample Selection" in wb.sheetnames: del wb["Sample Selection"]
sh = wb.create_sheet("Sample Selection")
sh.append(["State", "My GSTIN", "Sample Invoice", "Type", "Recipient", "Taxable Value", "Sample PO", "Basis of selection", "Document obtained? (CA)"])
for c in sh[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F4E79")
for k in sorted(basis, key=lambda k: (meta[k][0], -abs(docs[k]))):
    sh.append([meta[k][0], k[0], k[1], k[2], meta[k][1], round(docs[k], 2), po_text(k[1]), basis[k], ""])
for c_, w in zip("ABCDEFGHI", [20, 18, 18, 6, 40, 18, 34, 44, 22]): sh.column_dimensions[c_].width = w
sh.freeze_panes = "A2"
wb.save(P)
print("documents selected:", len(basis), "| rows marked:", n, "| with PO/SO ref:", sum(1 for k in basis if po_text(k[1])))
print("by basis:", dict(collections.Counter(basis.values())))
per = collections.Counter(k[0] for k in basis); print("per GSTIN min/max:", min(per.values()), "/", max(per.values()))
