"""MASTER part 2: S4 GL grid, S5 HSN/Rate, S6 Advances, S7 CN time-bar, S9 Sales Reco, merged Open Points."""
import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
SP = os.path.dirname(os.path.abspath(__file__))
def S(v): return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79")
TOT = Font(bold=True); AMB = PatternFill("solid", fgColor="FFF2CC")
OUT = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
D = r"C:\Users\pawar\Downloads"
wb = openpyxl.load_workbook(OUT)
ws0 = wb["SR_2025-26"]
H = {S(ws0.cell(4, c).value): L(c) for c in range(1, ws0.max_column + 1)}
R0, R1 = 5, 27006
SR = lambda name: "'SR_2025-26'!$%s$%d:$%s$%d" % (H[name], R0, H[name], R1)
NG = wb["GSTR-1 Data"].max_row; NGL = wb["GL Data"].max_row; NFS = wb["FS Revenue Data"].max_row
G1 = lambda c: "'GSTR-1 Data'!$%s$2:$%s$%d" % (c, c, NG)
GD = lambda c: "'GL Data'!$%s$2:$%s$%d" % (c, c, NGL)
FD = lambda c: "'FS Revenue Data'!$%s$2:$%s$%d" % (c, c, NFS)
GSTINS = [("01AAECR0503Q1ZM", "Jammu & Kashmir"), ("03AAECR0503Q1ZI", "Punjab"), ("06AAECR0503Q1ZC", "Haryana"),
 ("08AAECR0503Q1Z8", "Rajasthan"), ("10AAECR0503Q1ZN", "Bihar"), ("12AAECR0503Q1ZJ", "Arunachal Pradesh"),
 ("18AAECR0503Q1Z7", "Assam"), ("19AAECR0503Q1Z5", "West Bengal"), ("20AAECR0503Q1ZM", "Jharkhand"),
 ("23AAECR0503Q1ZG", "Madhya Pradesh"), ("24AAECR0503Q1ZE", "Gujarat"), ("27AAECR0503Q1Z8", "Maharashtra"),
 ("32AAECR0503Q1ZH", "Kerala"), ("33AAECR0503Q1ZF", "TamilNadu"), ("36AAECR0503Q1Z9", "Telangana"),
 ("37AAECR0503Q1Z7", "Andhra Pradesh"), ("09AAECR0503Q1Z6", "Uttar Pradesh"), ("22AAECR0503Q1ZI", "Chhattisgarh"),
 ("29AAECR0503Q1Z4", "Karnataka")]
for nm in ["S4 GL vs SR", "S4 Exceptions", "S5 HSN Summary", "S5 Rate-wise", "S6 Advances Control", "S6 GL Advance Check",
           "S7 CN Time-bar", "S9 Sales Reco", "Open Points (all)"]:
    if nm in wb.sheetnames: del wb[nm]

# ---------- S4 ----------
ws = wb.create_sheet("S4 GL vs SR")
ws["A1"] = "S4 - Output GL vs Sales Register (live). GL = 'GL Data' (sales-origin rows); SR = SR_2025-26."
ws["A1"].font = Font(bold=True, size=12)
ws.append([]); ws.append(["Business place", "State", "GL CGST", "SR CGST", "Diff CGST", "GL IGST", "SR IGST", "Diff IGST",
                          "Explained (S4 Exceptions)", "Residual"])
for c in ws[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
sw = openpyxl.load_workbook(D + r"\VEL_Step4_Reco_GL_vs_SR_DRAFT.xlsx", read_only=True)
ex4 = wb.create_sheet("S4 Exceptions")
for row in sw["Exceptions"].iter_rows(values_only=True): ex4.append(list(row))
sw.close()
for c in ex4[1]: c.font = HF; c.fill = HB
NE4 = ex4.max_row
for c_, w in zip("ABCDEFGH", [26, 18, 44, 10, 9, 16, 16, 20]): ex4.column_dimensions[c_].width = w
BPMAP = {"Jammu & Kashmir": "JK01", "Punjab": "PB01", "Haryana": "HR01", "Rajasthan": "RJ01", "Bihar": "BR01",
         "Arunachal Pradesh": "AR01", "Assam": "AS01", "West Bengal": "WB01", "Jharkhand": "JH01", "Madhya Pradesh": "MP01",
         "Gujarat": "GU01", "Maharashtra": "MH01", "Kerala": "KL01", "TamilNadu": "TN01", "Telangana": "TG01",
         "Andhra Pradesh": "AP01", "Uttar Pradesh": "UP01", "Chhattisgarh": "CG01", "Karnataka": "KA01"}
r = 3
for g_, st in GSTINS:
    r += 1
    ws.cell(r, 1, BPMAP.get(st, "")); ws.cell(r, 2, st)
    ws.cell(r, 3, '=SUMIFS(%s,%s,$B%d,%s,"CGST",%s,"Y")' % (GD("J"), GD("A"), r, GD("I"), GD("F")))
    ws.cell(r, 4, "=SUMIFS(%s,%s,$B%d)" % (SR("CGST Amount"), SR("My State"), r))
    ws.cell(r, 5, "=C%d-D%d" % (r, r))
    ws.cell(r, 6, '=SUMIFS(%s,%s,$B%d,%s,"IGST",%s,"Y")' % (GD("J"), GD("A"), r, GD("I"), GD("F")))
    ws.cell(r, 7, "=SUMIFS(%s,%s,$B%d)" % (SR("IGST Amount"), SR("My State"), r))
    ws.cell(r, 8, "=F%d-G%d" % (r, r))
    ws.cell(r, 9, "=SUMIFS('S4 Exceptions'!$H$2:$H$%d,'S4 Exceptions'!$B$2:$B$%d,$B%d)" % (NE4, NE4, r))
    ws.cell(r, 10, "=E%d-I%d" % (r, r))
    for c in range(3, 11): ws.cell(r, c).number_format = "#,##0.00"
tr = r + 1
ws.cell(tr, 2, "Total").font = TOT
for c in range(3, 11):
    ws.cell(tr, c, "=SUM(%s4:%s%d)" % (L(c), L(c), r)); ws.cell(tr, c).font = TOT; ws.cell(tr, c).number_format = "#,##0.00"
for c_, w in zip("ABCDEFGHIJ", [13, 20] + [16] * 8): ws.column_dimensions[c_].width = w

# ---------- S5 ----------
Rp = pd.read_pickle(os.path.join(SP, "register.pkl"))
Rp["advf"] = Rp["dtc"].astype(str).str.startswith("MOB")
it = Rp[~Rp["advf"]]
combos = it.groupby([it["gstin"].map(S), it["state"].map(S)]).size().reset_index()
h5 = wb.create_sheet("S5 HSN Summary")
h5["A1"] = "S5 - HSN summary (Table 17 shape), live over SR_2025-26. Blank UQC matched via '=' criterion."
h5["A1"].font = Font(bold=True)
h5.append([]); h5.append(["State", "My GSTIN", "HSN/SAC", "UQC (as in register; blank shown '-')", "Rate %", "Quantity", "Taxable", "IGST", "CGST", "SGST"])
for c in h5[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
wv = openpyxl.load_workbook(D + r"\VEL_Step5_HSN_Rate_Summary_DRAFT.xlsx", read_only=True)
si = wv["SR Items"]
cb = collections.Counter()
for row in si.iter_rows(min_row=2, values_only=True):
    cb[(S(row[0]), S(row[1]), S(row[3]), S(row[5]), row[6] or 0)] += 1
wv.close()
r = 3
for (g_, st, hsn, uqc, rate), _n in sorted(cb.items(), key=lambda x: (x[0][1], x[0][2], x[0][4])):
    r += 1
    h5.cell(r, 1, st); h5.cell(r, 2, g_); h5.cell(r, 3, hsn); h5.cell(r, 4, uqc); h5.cell(r, 5, rate)
    crit = ('%s,$B{r},%s,$C{r},%s,"="&IF($D{r}="-","",$D{r}),%s,$E{r},%s,"<>MOB*"'
            % (SR("My GSTIN"), SR("HSN or SAC Code"), SR("Unit of Measurement for uploading"), SR("GST Rate"), SR("Document Type Code"))).replace("{r}", str(r))
    for j, name in enumerate(["Quantity for uploading", "Taxable Value", "IGST Amount", "CGST Amount", "SGST Amount"], start=6):
        h5.cell(r, j, "=SUMIFS(%s,%s)" % (SR(name), crit)); h5.cell(r, j).number_format = "#,##0.00"
NH5 = r
r += 1
h5.cell(r, 1, "Total").font = TOT
for j in range(6, 11):
    h5.cell(r, j, "=SUM(%s4:%s%d)" % (L(j), L(j), NH5)); h5.cell(r, j).font = TOT; h5.cell(r, j).number_format = "#,##0.00"
h5.auto_filter.ref = "A3:J%d" % NH5
for c_, w in zip("ABCDEFGHIJ", [18, 18, 12, 16, 8, 13, 16, 13, 13, 13]): h5.column_dimensions[c_].width = w
r5 = wb.create_sheet("S5 Rate-wise")
r5["A1"] = "S5 - Rate-wise summary, live over SR_2025-26 (advances excluded)."; r5["A1"].font = Font(bold=True)
r5.append([]); r5.append(["State", "My GSTIN", "Rate %", "Taxable", "IGST", "CGST", "SGST", "Actual rate %"])
for c in r5[3]: c.font = HF; c.fill = HB
rates = sorted({(S(a), S(b), c_ or 0) for a, b, c_ in zip(it["gstin"], it["state"], it.get("rate", pd.Series([None] * len(it))))} if "rate" in it else set())
# rate not in pickle: derive combos from HSN combos' rates
rc = sorted({(g_, st, rate) for (g_, st, hsn, uqc, rate) in cb})
r = 3
for g_, st, rate in rc:
    r += 1
    r5.cell(r, 1, st); r5.cell(r, 2, g_); r5.cell(r, 3, rate)
    crit = ('%s,$B{r},%s,$C{r},%s,"<>MOB*"' % (SR("My GSTIN"), SR("GST Rate"), SR("Document Type Code"))).replace("{r}", str(r))
    for j, name in enumerate(["Taxable Value", "IGST Amount", "CGST Amount", "SGST Amount"], start=4):
        r5.cell(r, j, "=SUMIFS(%s,%s)" % (SR(name), crit)); r5.cell(r, j).number_format = "#,##0.00"
    r5.cell(r, 8, '=IF(N(D%d)=0,"",ROUND((E%d+F%d+G%d)/D%d*100,2))' % (r, r, r, r, r))
NR5 = r
r += 1
r5.cell(r, 1, "Total").font = TOT
for j in range(4, 8):
    r5.cell(r, j, "=SUM(%s4:%s%d)" % (L(j), L(j), NR5)); r5.cell(r, j).font = TOT; r5.cell(r, j).number_format = "#,##0.00"
for c_, w in zip("ABCDEFGH", [18, 18, 8, 16, 13, 13, 13, 12]): r5.column_dimensions[c_].width = w

# ---------- S6 ----------
a6 = wb.create_sheet("S6 Advances Control")
a6["A1"] = "S6 - Advances control account (live over SR_2025-26 via 'Adv Bucket (helper)'). Openings = audited FY 24-25 closings."
a6["A1"].font = Font(bold=True, size=12)
a6.append([]); a6.append(["State", "GSTIN", "Opening 01.04.25", "Received (books)", "Adjusted (books, |sum|)", "Closing",
                          "GSTR-1 net advances", "Books net", "Books vs GSTR-1", "Remark"])
for c in a6[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
OPEN = {"Arunachal Pradesh": 12398533.90, "Bihar": 729161.86, "Gujarat": 52303577.12, "Madhya Pradesh": 132470378.81}
r = 3
for g_, st in GSTINS:
    r += 1
    a6.cell(r, 1, st); a6.cell(r, 2, g_); a6.cell(r, 3, OPEN.get(st, 0.0))
    a6.cell(r, 4, '=SUMIFS(%s,%s,$B%d,%s,"Received")' % (SR("Taxable Value"), SR("My GSTIN"), r, SR("Adv Bucket (helper)")))
    a6.cell(r, 5, '=-SUMIFS(%s,%s,$B%d,%s,"Adjusted")' % (SR("Taxable Value"), SR("My GSTIN"), r, SR("Adv Bucket (helper)")))
    a6.cell(r, 6, "=C%d+D%d-E%d" % (r, r, r))
    a6.cell(r, 7, '=SUMIFS(%s,%s,$B%d,%s,"SalesSummary-Net (summary)",%s,"Advance*")' % (G1("F"), G1("A"), r, G1("C"), G1("J")))
    a6.cell(r, 8, "=D%d-E%d" % (r, r)); a6.cell(r, 9, "=H%d-G%d" % (r, r))
    a6.cell(r, 10, '=IF(F%d<-1,"OVER-ADJUSTED - check","")' % r)
    for c in range(3, 10): a6.cell(r, c).number_format = "#,##0.00"
tr6 = r + 1
a6.cell(tr6, 1, "Total").font = TOT
for c in range(3, 10):
    a6.cell(tr6, c, "=SUM(%s4:%s%d)" % (L(c), L(c), r)); a6.cell(tr6, c).font = TOT; a6.cell(tr6, c).number_format = "#,##0.00"
for c_, w in zip("ABCDEFGHIJ", [18, 18, 16, 16, 16, 16, 16, 16, 14, 26]): a6.column_dimensions[c_].width = w
sw = openpyxl.load_workbook(D + r"\VEL_Step6_Advances_Control_DRAFT.xlsx", read_only=True)
gc = wb.create_sheet("S6 GL Advance Check")
for row in sw["GL Advance Check"].iter_rows(values_only=True): gc.append(list(row))
sw.close()
for c_, w in zip("ABCDE", [14, 20, 26, 26, 22]): gc.column_dimensions[c_].width = w

# ---------- S7 ----------
s7 = wb.create_sheet("S7 CN Time-bar")
s7["A1"] = ("S7 - Credit-note Sec 34(2) time-bar. Original invoice no/date PULLED LIVE from SR_2025-26 cols %s/%s - "
            "fill them in the register and this sheet retests itself." % (H["Original Invoice Number"], H["Original Invoice Date"]))
s7["A1"].font = Font(bold=True)
s7.append([]); s7.append(["State", "CN Number", "CN Date", "Taxable", "Original Invoice No (live)", "Original Invoice Date (live)",
                          "Declare-by (30 Nov)", "Status (live)"])
for c in s7[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
cns = pd.read_pickle(os.path.join(SP, "cn_rows.pkl"))
docs = cns.groupby(["gstin", "docno"]).agg(state=("state", "first"), cndate=("cndate", "first"), tax=("tax", "sum")).reset_index()
r = 3
for _, d_ in docs.sort_values(["state", "docno"]).iterrows():
    r += 1
    s7.cell(r, 1, S(d_["state"])); s7.cell(r, 2, S(d_["docno"])); s7.cell(r, 3, d_["cndate"]); s7.cell(r, 4, round(float(d_["tax"]), 2))
    s7.cell(r, 5, "=IFERROR(INDEX('SR_2025-26'!$%s:$%s,MATCH($B%d,'SR_2025-26'!$%s:$%s,0)),\"\")" % (
        H["Original Invoice Number"], H["Original Invoice Number"], r, H["Document Number"], H["Document Number"]))
    s7.cell(r, 6, "=IFERROR(INDEX('SR_2025-26'!$%s:$%s,MATCH($B%d,'SR_2025-26'!$%s:$%s,0)),\"\")" % (
        H["Original Invoice Date"], H["Original Invoice Date"], r, H["Document Number"], H["Document Number"]))
    s7.cell(r, 7, '=IF(F%d="","",DATE(IF(MONTH(F%d)>=4,YEAR(F%d)+1,YEAR(F%d)),11,30))' % (r, r, r, r))
    s7.cell(r, 8, ('=IF(F{0}="","Original ref not captured - cannot test",IF(F{0}>=DATE(2025,4,1),"OK - original in FY 25-26",'
                   'IF(C{0}>G{0},"BEYOND Sec 34(2) WINDOW","OK - within window")))').format(r))
    s7.cell(r, 3).number_format = "dd-mmm-yy"; s7.cell(r, 6).number_format = "dd-mmm-yy"; s7.cell(r, 7).number_format = "dd-mmm-yy"
    s7.cell(r, 4).number_format = "#,##0.00"
for c_, w in zip("ABCDEFGH", [18, 16, 11, 14, 20, 14, 13, 44]): s7.column_dimensions[c_].width = w
s7.freeze_panes = "A4"

# ---------- S9 Sales Reco (port; FS live, GST live over register) ----------
s9 = wb.create_sheet("S9 Sales Reco")
s9["A1"] = "S9 - Reconciliation of turnover per audited FS with GST returns (last year's layout; 9C Table 5 refs). Amber rows = CA."
s9["A1"].font = Font(bold=True, size=12)
ORDER = [("37AAECR0503Q1Z7", "Andhra Pradesh"), ("12AAECR0503Q1ZJ", "Arunachal Pradesh"), ("18AAECR0503Q1Z7", "Assam"),
 ("10AAECR0503Q1ZN", "Bihar"), ("22AAECR0503Q1ZI", "Chhattisgarh"), ("24AAECR0503Q1ZE", "Gujarat"), ("06AAECR0503Q1ZC", "Haryana"),
 ("01AAECR0503Q1ZM", "Jammu & Kashmir"), ("20AAECR0503Q1ZM", "Jharkhand"), ("29AAECR0503Q1Z4", "Karnataka"),
 ("32AAECR0503Q1ZH", "Kerala"), ("23AAECR0503Q1ZG", "Madhya Pradesh"), ("27AAECR0503Q1Z8", "Maharashtra"),
 ("03AAECR0503Q1ZI", "Punjab"), ("08AAECR0503Q1Z8", "Rajasthan"), ("33AAECR0503Q1ZF", "Tamil Nadu"),
 ("36AAECR0503Q1Z9", "Telangana"), ("09AAECR0503Q1Z6", "Uttar Pradesh"), ("19AAECR0503Q1Z5", "West Bengal")]
NST = len(ORDER); TC = NST + 2; CR = NST + 3
s9.cell(3, 1, "Particulars"); s9.cell(4, 1, "GSTN")
for j, (g_, st) in enumerate(ORDER, start=2):
    s9.cell(3, j, st); s9.cell(4, j, g_)
s9.cell(3, TC, "Total"); s9.cell(3, CR, "9C ref")
for rr in (3, 4):
    for c in range(1, CR + 1):
        x = s9.cell(rr, c); x.font = HF; x.fill = HB; x.alignment = Alignment(horizontal="center", wrap_text=True)
row = {}
def line9(r, label, cellfn=None, ref="", amber=False, bold=False):
    s9.cell(r, 1, label)
    if bold: s9.cell(r, 1).font = TOT
    for j, (g_, st) in enumerate(ORDER, start=2):
        v = cellfn(g_, st, j, r) if cellfn else None
        if v is not None: s9.cell(r, j, v)
        if amber: s9.cell(r, j).fill = AMB
        s9.cell(r, j).number_format = "#,##0"
    s9.cell(r, TC, "=SUM(B%d:%s%d)" % (r, L(NST + 1), r)); s9.cell(r, TC).font = TOT; s9.cell(r, TC).number_format = "#,##0"
    if ref: s9.cell(r, CR, ref)
    row[label] = r
FIX9 = {"Chhatisgarh": "Chhattisgarh"}
r = 6
line9(r, "Revenue from Operations (FS)", lambda g_, st, j, r_: "=SUMIFS(%s,%s,%s)" % (FD("C"), FD("B"), '"%s"' % st)); r += 1
line9(r, "Total GST turnover (register: billed + advances net)", lambda g_, st, j, r_: "=SUMIFS(%s,%s,\"%s\")" % (SR("Taxable Value"), SR("My GSTIN"), g_), bold=True); r += 1
line9(r, "Difference (A)", lambda g_, st, j, r_: "=%s%d-%s%d" % (L(j), r_ - 2, L(j), r_ - 1), bold=True); r += 2
s9.cell(r, 1, "Reasons (CA to fill; 9C Table 5 refs at right)").font = TOT; r += 1
start = r
line9(r, "Unbilled / uncertified revenue movement", None, ref="5B/5H", amber=True); r += 1
line9(r, "Discounts not permissible under GST", None, ref="5O", amber=True); r += 1
line9(r, "Other income / non-GST items", None, amber=True); r += 1
line9(r, "Unadjusted advances - opening (-) / closing (+)", lambda g_, st, j, r_: {"Arunachal Pradesh": -12398533.90 + 8591511.68, "Bihar": -729161.86 - 4.81, "Gujarat": -52303577.12 + 1233643.83, "Madhya Pradesh": -132470378.81 + 213995089.93}.get(st), ref="5C/5I"); r += 1
line9(r, "Others", None, amber=True); r += 1
line9(r, "Total reasons (B)", lambda g_, st, j, r_: "=SUM(%s%d:%s%d)" % (L(j), start, L(j), r_ - 1), bold=True); r += 1
line9(r, "Net unexplained (A - B)", lambda g_, st, j, r_: "=%s%d-%s%d" % (L(j), row["Difference (A)"], L(j), row["Total reasons (B)"]), bold=True)
s9.column_dimensions["A"].width = 46
for c in range(2, CR + 1): s9.column_dimensions[L(c)].width = 15
s9.freeze_panes = "B5"

# ---------- merged Open Points ----------
op = wb.create_sheet("Open Points (all)")
op.append(["Step", "#", "Owner", "Point", "Detail", "Resolved?"])
for c in op[1]: c.font = HF; c.fill = HB
FILES = [("S1", "VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"), ("S2", "VEL_Step2_Reco_GSTR1_DRAFT.xlsx"),
         ("S3", "VEL_Step3_Reco_GSTR1_vs_3B_DRAFT.xlsx"), ("S4", "VEL_Step4_Reco_GL_vs_SR_DRAFT.xlsx"),
         ("S5", "VEL_Step5_HSN_Rate_Summary_DRAFT.xlsx"), ("S6", "VEL_Step6_Advances_Control_DRAFT.xlsx"),
         ("S7", "VEL_Step7_CN_TimeBar_DRAFT.xlsx"), ("S9", "VEL_Step9_FS_vs_GST_DRAFT.xlsx")]
n = 0
for tag, fn in FILES:
    try:
        sw = openpyxl.load_workbook(D + "\\" + fn, read_only=True)
        if "Open Points" in sw.sheetnames:
            for row in sw["Open Points"].iter_rows(min_row=5, values_only=True):
                if row and row[0] is not None:
                    op.append([tag] + list(row[:5])); n += 1
        sw.close()
    except Exception as e:
        print("open points skip", fn, type(e).__name__)
for c_, w in zip("ABCDEF", [6, 4, 9, 50, 90, 18]): op.column_dimensions[c_].width = w
for rr in op.iter_rows(min_row=2):
    for c in rr: c.alignment = Alignment(wrap_text=True, vertical="top")
op.freeze_panes = "A2"
wb.save(OUT)
print("part 2 done: S4/S5/S6/S7/S9 + Open Points (all) with %d points | sheets: %d" % (n, len(wb.sheetnames)))
