"""Append July-25 RCM rows (missing from Conso RCM) to 'RCM Register' from the 15 monthly
'GSTR-3B <State> July 2025.xlsx' -> sheet 'RCM'. Posting grain like the rest of the register:
a monthly row with G/L '2610080300-301' (CGST+SGST combined) becomes TWO rows (CGST 2610080300,
SGST 2610080301), taxable value split 50/50 so state-month SUMIFS tie; IGST rows stay single.
Every formula column re-applied; dependents' ranges extended to the new last row; new rows get a
'Source' tag. Rows inserted in 3B-month order (after the Jun-25 block) so filters read naturally."""
import os, re, pickle, warnings, datetime as dt, collections
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.utils import get_column_letter as L
from copy import copy
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v)
    except Exception: return 0.0
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
SP = os.path.dirname(os.path.abspath(__file__))
f = open(P, "r+b"); f.close()
pk = pickle.load(open(os.path.join(SP, "rcm_monthly.pkl"), "rb")); D = pk["data"]
G2ST = {"01AAECR0503Q1ZM": "Jammu & Kashmir", "03AAECR0503Q1ZI": "Punjab", "06AAECR0503Q1ZC": "Haryana", "08AAECR0503Q1Z8": "Rajasthan",
        "10AAECR0503Q1ZN": "Bihar", "12AAECR0503Q1ZJ": "Arunachal Pradesh", "18AAECR0503Q1Z7": "Assam", "19AAECR0503Q1Z5": "West Bengal",
        "20AAECR0503Q1ZM": "Jharkhand", "23AAECR0503Q1ZG": "Madhya Pradesh", "24AAECR0503Q1ZE": "Gujarat", "27AAECR0503Q1Z8": "Maharashtra",
        "32AAECR0503Q1ZH": "Kerala", "33AAECR0503Q1ZF": "Tamil Nadu", "36AAECR0503Q1Z9": "Telangana", "37AAECR0503Q1Z7": "Andhra Pradesh",
        "09AAECR0503Q1Z6": "Uttar Pradesh", "22AAECR0503Q1ZI": "Chhattisgarh", "29AAECR0503Q1Z4": "Karnataka"}
GLN = {"2610080300": "CGST Output RCM", "2610080301": "SGST Output RCM", "2610080302": "IGST Output RCM"}
def fy(d):
    if not isinstance(d, dt.datetime): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
# ---- build new rows (posting grain)
new = []
for (g, m), rec in sorted(D.items()):
    if m != "Jul-25": continue
    hi = {}
    for i, h in enumerate(rec["header"]):
        if h and h not in hi: hi[h] = i
    for r in rec["rows"]:
        d = lambda h: r[hi[h]] if h in hi and hi[h] < len(r) else None
        gl = S(d("G/L Account")); tv = num(d("Taxable Value")); ig, cg, sg = num(d("IGST")), num(d("CGST")), num(d("SGST"))
        base = {
            "3B Month": dt.datetime(2025, 7, 1), "Business place": S(d("Business place")), "MY GSTN": g, "State": G2ST[g],
            "Fiscal Year": d("Fiscal Year"), "Year/Month": S(d("Year/Month")), "Document type": S(d("Document type")),
            "Document Number": d("Document Number"), "Posting Date": d("Posting Date"), "Posting Year": fy(d("Posting Date")),
            "Assignment": S(d("Assignment")), "Reference": S(d("Reference")), "Document Date": d("Document Date"), "Doc year": fy(d("Document Date")),
            "Vendor Code": d("Vendor Code"), "GSTN": S(d("Vendor GSTN")), "Vendor Name": S(d("Vendor Name")), "Posting Key": d("Posting Key"),
            "Local Currency": S(d("Local Currency")) or "INR", "Tax Code": S(d("Tax Code")), "Clearing Document": d("Clearing Document"),
            "Profit Center": d("Profit Center"), "Text": S(d("Text")), "Offsetting Account": d("Offsetting Account"),
            "GST Rate": num(d("Rate")) if S(d("Rate")) else (round((ig + cg + sg) / tv * 100, 2) if tv else None),
            "Nature of Services": S(d("Category")), "GSTR 3B Claim month": "05 Aug 2025",
            "Source": "Monthly working Jul-25 (RCM sheet) - not in Conso RCM",
        }
        if gl.startswith("2610080302") or (ig and not cg):
            row = dict(base); row.update({"G/L Account": "2610080302", "GL Name": GLN["2610080302"], "Amount in Local Currency": -ig,
                                          "Taxable Value as per SAP": tv, "IGST AS PER SAP": ig, "CGST AS PER SAP": 0.0, "SGST AS PER SAP": 0.0})
            new.append(row)
        else:
            for acct, amt, tag in (("2610080300", cg, "C"), ("2610080301", sg, "S")):
                row = dict(base); row.update({"G/L Account": acct, "GL Name": GLN[acct], "Amount in Local Currency": -amt,
                                              "Taxable Value as per SAP": tv / 2, "IGST AS PER SAP": 0.0,
                                              "CGST AS PER SAP": cg if tag == "C" else 0.0, "SGST AS PER SAP": sg if tag == "S" else 0.0})
                new.append(row)
print("new July posting rows:", len(new), "| taxable %.2f | IGST %.2f CGST %.2f SGST %.2f" % (
    sum(r["Taxable Value as per SAP"] for r in new), sum(r["IGST AS PER SAP"] for r in new), sum(r["CGST AS PER SAP"] for r in new), sum(r["SGST AS PER SAP"] for r in new)))
# ---- open master
wb = openpyxl.load_workbook(P)
rr = wb["RCM Register"]
RH = {}
for c in range(1, rr.max_column + 1):
    h = S(rr.cell(5, c).value)
    if h and h not in RH: RH[h] = c
NC = max(RH.values()); R0 = 6
OLDN = max(r for r in range(R0, rr.max_row + 1) if rr.cell(r, RH["MY GSTN"]).value or rr.cell(r, RH["Business place"]).value)
if "Source" not in RH:
    NC += 1; RH["Source"] = NC
    x = rr.cell(5, NC, "Source"); x.font = copy(rr.cell(5, NC - 1).font); x.fill = copy(rr.cell(5, NC - 1).fill); rr.column_dimensions[L(NC)].width = 40
    for r in range(R0, OLDN + 1): rr.cell(r, NC).value = "Conso RCM FY 25-26 (client working)"
# insert point: first row whose 3B Month > Jul-25
ins = None
for r in range(R0, OLDN + 1):
    v = rr.cell(r, RH["3B Month"]).value
    if isinstance(v, dt.datetime) and v > dt.datetime(2025, 7, 1): ins = r; break
ins = ins or OLDN + 1
K = len(new)
print("inserting %d rows at %d (old last %d)" % (K, ins, OLDN))
# formula templates from row R0 (formula columns) - rebuilt after insert for ALL rows to keep refs consistent
col_formula = {c: rr.cell(R0, c).value for c in range(1, NC + 1) if isinstance(rr.cell(R0, c).value, str) and str(rr.cell(R0, c).value).startswith("=")}
style_src = {c: copy(rr.cell(R0, c)._style) for c in range(1, NC + 1)}
# ---- do the insert via openpyxl: shift values manually (openpyxl insert_rows does not adjust formulas -> we re-write every formula anyway)
rr.insert_rows(ins, K)
NEWN = OLDN + K
for i, row in enumerate(new):
    r = ins + i
    for h, c in RH.items():
        if c in col_formula: continue
        rr.cell(r, c).value = row.get(h)
    for c in range(1, NC + 1): rr.cell(r, c)._style = copy(style_src[c])
# re-write formula columns for every data row (row-relative refs)
def rel(fx, r):
    return re.sub(r"(?<![A-Z!])(\$?[A-Z]{1,3}\$?)%d(?!\d)" % R0, lambda m: m.group(1) + str(r), fx)
for r in range(R0, NEWN + 1):
    for c, fx in col_formula.items(): rr.cell(r, c).value = rel(fx, r)
# subtotal row 4
for c in range(1, NC + 1):
    v = rr.cell(4, c).value
    if isinstance(v, str) and "SUBTOTAL" in v:
        rr.cell(4, c).value = re.sub(r"(\$?[A-Z]{1,3}\$?)%d\)" % OLDN, lambda m: "%s%d)" % (m.group(1), NEWN), v)
rr.auto_filter.ref = "A5:%s%d" % (L(NC), NEWN)
# ---- repoint dependents ($X$6:$X$OLDN -> NEWN)
pat = re.compile(r"('RCM Register'!\$[A-Z]{1,3}\$)(\d+)(:\$[A-Z]{1,3}\$)(\d+)")
n = 0
for nm in wb.sheetnames:
    if nm == "RCM Register": continue
    for rowc in wb[nm].iter_rows():
        for cell in rowc:
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "RCM Register" in v:
                nv = pat.sub(lambda m: "%s%s%s%d" % (m.group(1), m.group(2), m.group(3), NEWN) if int(m.group(4)) == OLDN else m.group(0), v)
                if nv != v: cell.value = nv; n += 1
print("dependent formulas repointed:", n, "| register now rows %d..%d" % (R0, NEWN))
# TB scrutiny COUNTIF ranges + any other hard OLDN refs handled by the same regex (they use $..$6:$..$OLDN)
wb.save(P); print("saved")
