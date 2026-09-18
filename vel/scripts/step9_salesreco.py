"""Q12 - rebuild the FS reco in the CA's own 'Sales Reco' layout (decoded from the FY 24-25
workbook): states across columns, FS block / GSTR-9 block / Difference (A) / two reasons blocks
with GSTR-9C Table 5 codes (5B/5C/5H/5I/5O) / Total (B) / Net Difference. Judgment rows amber.
Also fills register column AV 'Matched with FS - revenue' (state-level reco marker)."""
import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
SP = os.path.dirname(os.path.abspath(__file__))
def S(v): return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79")
AMB = PatternFill("solid", fgColor="FFF2CC"); TOT = Font(bold=True); GF = PatternFill("solid", fgColor="DDEBF7")
GSTINS = [("37AAECR0503Q1Z7", "Andhra Pradesh"), ("12AAECR0503Q1ZJ", "Arunachal Pradesh"), ("18AAECR0503Q1Z7", "Assam"),
 ("10AAECR0503Q1ZN", "Bihar"), ("22AAECR0503Q1ZI", "Chhattisgarh"), ("24AAECR0503Q1ZE", "Gujarat"),
 ("06AAECR0503Q1ZC", "Haryana"), ("01AAECR0503Q1ZM", "Jammu & Kashmir"), ("20AAECR0503Q1ZM", "Jharkhand"),
 ("29AAECR0503Q1Z4", "Karnataka"), ("32AAECR0503Q1ZH", "Kerala"), ("23AAECR0503Q1ZG", "Madhya Pradesh"),
 ("27AAECR0503Q1Z8", "Maharashtra"), ("03AAECR0503Q1ZI", "Punjab"), ("08AAECR0503Q1Z8", "Rajasthan"),
 ("33AAECR0503Q1ZF", "Tamil Nadu"), ("36AAECR0503Q1Z9", "Telangana"), ("09AAECR0503Q1Z6", "Uttar Pradesh"),
 ("19AAECR0503Q1Z5", "West Bengal")]   # last year's column order, alphabetical-ish
NST = len(GSTINS)
R = pd.read_pickle(os.path.join(SP, "register.pkl"))
R["st"] = R["state"].map(S)
def bucket(r):
    d = S(r["dtc"])
    if d == "INV": return "Sales"
    if d == "CRN": return "Credit Notes"
    if d == "DBN": return "Debit Notes"
    if d == "MOB ADV REC": return "Advance Received"
    if d == "MOB ADV ADJ": return "Advance Adjusted"
    return "Advance Received" if (r["tax"] or 0) > 0 else "Advance Adjusted"
R["bk"] = R.apply(bucket, axis=1)
comp = R.groupby(["st", "bk"])["tax"].sum().round(2)
OPEN = {"Arunachal Pradesh": 12398533.90, "Bihar": 729161.86, "Gujarat": 52303577.12, "Madhya Pradesh": 132470378.81}
CLOSE = {}  # closing = opening + rec - |adj| per state (from Step 6 logic)
for g, st in GSTINS:
    rec = float(comp.get((st, "Advance Received"), 0)); adj = float(comp.get((st, "Advance Adjusted"), 0))
    CLOSE[st] = round(OPEN.get(st, 0.0) + rec + adj, 2)   # adj is negative in data
P = r"C:\Users\pawar\Downloads\VEL_Step9_FS_vs_GST_DRAFT.xlsx"
wb = openpyxl.load_workbook(P)
if "Sales Reco" in wb.sheetnames: del wb["Sales Reco"]
ws = wb.create_sheet("Sales Reco", 1)
NF = wb["FS Revenue Data"].max_row
FD = lambda c: "'FS Revenue Data'!$%s$2:$%s$%d" % (c, c, NF)
TCOL = NST + 2   # Total column
CREF = NST + 3   # 9C table ref column
ws["A1"] = "Vikran Engineering Limited"; ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "Reconciliation Statement — FY 2025-26 (last year's 'Sales Reco' layout; column %s maps each reason line to GSTR-9C Table 5)" % L(CREF)
ws["A4"] = "Reconciliation of turnover declared in audited Annual Financial Statement with turnover declared in GST returns (GSTR-9C Pt.II)"
ws["A4"].font = Font(bold=True)
ws.cell(5, 1, "Particulars"); ws.cell(6, 1, "GSTN")
for j, (g, st) in enumerate(GSTINS, start=2):
    ws.cell(5, j, st); ws.cell(6, j, g)
ws.cell(5, TCOL, "Total"); ws.cell(6, TCOL, "All 19"); ws.cell(5, CREF, "9C Table ref")
for rr in (5, 6):
    for c in range(1, CREF + 1):
        x = ws.cell(rr, c); x.font = HF; x.fill = HB; x.alignment = Alignment(horizontal="center", wrap_text=True)
row = {}
def line(r, label, valuefn=None, ref="", amber=False, bold=False):
    ws.cell(r, 1, label)
    if bold: ws.cell(r, 1).font = TOT
    for j, (g, st) in enumerate(GSTINS, start=2):
        v = valuefn(st, j) if valuefn else None
        if v is not None: ws.cell(r, j, v)
        if amber: ws.cell(r, j).fill = AMB
        ws.cell(r, j).number_format = "#,##0"
    ws.cell(r, TCOL, "=SUM(B%d:%s%d)" % (r, L(NST + 1), r)); ws.cell(r, TCOL).number_format = "#,##0"; ws.cell(r, TCOL).font = TOT
    if ref: ws.cell(r, CREF, ref)
    row[label] = r
r = 8
ws.cell(r, 1, "Details as per Financial Statements").font = TOT; r += 1
line(r, "Revenue from Operations", lambda st, j: "=SUMIFS(%s,%s,\"%s\")" % (FD("C"), FD("B"), st)); r += 1
line(r, "Other Income", lambda st, j: 0); ws.cell(r, TCOL).value = 169648973.83; ws.cell(r, TCOL).number_format = "#,##0"; row["Other Income"] = r; r += 1
line(r, "Total Income", lambda st, j: "=%s%d+%s%d" % (L(j), r - 2, L(j), r - 1), bold=True); r += 2
ws.cell(r, 1, "Details as per Form GSTR-9").font = TOT; r += 1
def cval(name):
    return lambda st, j: float(comp.get((st, name), 0.0))
line(r, "Sales", cval("Sales")); r += 1
line(r, "Credit Notes", cval("Credit Notes")); r += 1
line(r, "Advance Received", cval("Advance Received")); r += 1
line(r, "Debit Notes", cval("Debit Notes")); r += 1
line(r, "Scrap Sales", lambda st, j: 0.0); r += 1
line(r, "Advance Adjusted", cval("Advance Adjusted")); r += 1
line(r, "Sales STO", lambda st, j: 0.0, amber=True); r += 1
first_g = row["Sales"]
line(r, "Total Sales as per GST Return", lambda st, j: "=SUM(%s%d:%s%d)" % (L(j), first_g, L(j), r - 1), bold=True); r += 2
line(r, "Difference (A)", lambda st, j: "=%s%d-%s%d" % (L(j), row["Total Income"], L(j), row["Total Sales as per GST Return"]), bold=True); r += 2
ws.cell(r, 1, "Reasons for Differences:").font = TOT; r += 1
ws.cell(r, 1, "In Financials not in GST returns").font = TOT; r += 1
start_a = r
line(r, "Discount Given", lambda st, j: None, ref="5O", amber=True); r += 1
line(r, "Foreseeable reversal", lambda st, j: None, amber=True); r += 1
line(r, "Not considered in return / Uncertified revenue", lambda st, j: None, ref="5H", amber=True); r += 1
line(r, "Sales (AS-7 / unbilled impact)", lambda st, j: None, ref="5B/5H", amber=True); r += 1
line(r, "Other Incomes (non-GST)", lambda st, j: None, amber=True); r += 1
line(r, "Total (Financials-side reasons)", lambda st, j: "=SUM(%s%d:%s%d)" % (L(j), start_a, L(j), r - 1), bold=True); r += 2
ws.cell(r, 1, "In GST returns not in Financials").font = TOT; r += 1
start_b = r
line(r, "Unadjusted advance at the beginning of FY (-)", lambda st, j: -OPEN.get(st, 0.0), ref="5I"); r += 1
line(r, "Unadjusted advance at the end of FY (+)", lambda st, j: CLOSE.get(st, 0.0), ref="5C"); r += 1
line(r, "Scrap sales reduced from cost", lambda st, j: None, amber=True); r += 1
line(r, "Sales STO", lambda st, j: None, amber=True); r += 1
line(r, "Total (GST-side reasons)", lambda st, j: "=SUM(%s%d:%s%d)" % (L(j), start_b, L(j), r - 1), bold=True); r += 2
tb1 = row["Total (Financials-side reasons)"]; tb2 = row["Total (GST-side reasons)"]
line(r, "Total (B)", lambda st, j: "=-%s%d+%s%d" % (L(j), tb1, L(j), tb2), bold=True); r += 1
line(r, "Net Difference (A) - explained", lambda st, j: "=%s%d-(%s%d-%s%d)" % (L(j), row["Difference (A)"], L(j), tb1, L(j), tb2), bold=True)
net_r = r
ws.cell(net_r, 1).value = "Net unexplained (A minus reasons)"
ws.column_dimensions["A"].width = 44
for c in range(2, CREF + 1): ws.column_dimensions[L(c)].width = 16
ws.column_dimensions[L(CREF)].width = 12
ws.freeze_panes = "B7"
wb.save(P); print("Sales Reco sheet built: rows 5-%d, %d state cols + Total + 9C ref" % (net_r, NST))

# ---- register AV fill
REG = r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
try:
    f = open(REG, "r+b"); f.close()
    wb2 = openpyxl.load_workbook(REG); ws2 = wb2["SR_2025-26"]
    H = {S(ws2.cell(4, c).value): c for c in range(1, ws2.max_column + 1)}
    cAV = H["Matched with FS - revenue"]
    n = 0
    for i in range(5, 27007):
        ws2.cell(i, cAV).value = "FS reco at state level - see Step 9 'Sales Reco'"
        n += 1
    wb2.save(REG); print("register AV filled on %d rows" % n)
except Exception as e:
    print("register LOCKED - AV fill pending:", type(e).__name__)
