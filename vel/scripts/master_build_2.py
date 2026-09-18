"""MASTER part 1b: S2 + S3 suites, single-source (register) SUMIFS."""
import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as L
SP = os.path.dirname(os.path.abspath(__file__))
def S(v): return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); GF = PatternFill("solid", fgColor="DDEBF7")
TOT = Font(bold=True); AMB = PatternFill("solid", fgColor="FFF2CC"); RED = PatternFill("solid", fgColor="FFC7CE")
thin = Side(style="thin", color="B0B0B0"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
OUT = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
M = pd.read_pickle(os.path.join(SP, "step2_match2.pkl"))
wb = openpyxl.load_workbook(OUT)
ws0 = wb["SR_2025-26"]
H = {S(ws0.cell(4, c).value): L(c) for c in range(1, ws0.max_column + 1)}
R0, R1 = 5, 27006
SR = lambda name: "'SR_2025-26'!$%s$%d:$%s$%d" % (H[name], R0, H[name], R1)
NG = wb["GSTR-1 Data"].max_row; N3 = wb["3B Data"].max_row
G1 = lambda c: "'GSTR-1 Data'!$%s$2:$%s$%d" % (c, c, NG)
B3 = lambda c: "'3B Data'!$%s$2:$%s$%d" % (c, c, N3)
G1m = {0: "F", 1: "G", 2: "H", 3: "I"}; B3m = {0: "D", 1: "E", 2: "F", 3: "G"}
SRm = {0: "Taxable Value", 1: "IGST Amount", 2: "CGST Amount", 3: "SGST Amount"}
GSTINS = [("01AAECR0503Q1ZM", "Jammu & Kashmir"), ("03AAECR0503Q1ZI", "Punjab"), ("06AAECR0503Q1ZC", "Haryana"),
 ("08AAECR0503Q1Z8", "Rajasthan"), ("10AAECR0503Q1ZN", "Bihar"), ("12AAECR0503Q1ZJ", "Arunachal Pradesh"),
 ("18AAECR0503Q1Z7", "Assam"), ("19AAECR0503Q1Z5", "West Bengal"), ("20AAECR0503Q1ZM", "Jharkhand"),
 ("23AAECR0503Q1ZG", "Madhya Pradesh"), ("24AAECR0503Q1ZE", "Gujarat"), ("27AAECR0503Q1Z8", "Maharashtra"),
 ("32AAECR0503Q1ZH", "Kerala"), ("33AAECR0503Q1ZF", "TamilNadu"), ("36AAECR0503Q1Z9", "Telangana"),
 ("37AAECR0503Q1Z7", "Andhra Pradesh"), ("09AAECR0503Q1Z6", "Uttar Pradesh"), ("22AAECR0503Q1ZI", "Chhattisgarh"),
 ("29AAECR0503Q1Z4", "Karnataka")]
MONTHS = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26"]
GROUPS = ["B2B", "B2C", "Credit note", "Debit note", "Advance received", "Advance adjusted", "Total liability"]
MEAS = ["Taxable value", "IGST", "CGST", "SGST"]
def books(mi, gcell, gn, extra=""):
    m = SR(SRm[mi]); base = "SUMIFS(%s,%s,%s%s" % (m, SR("My GSTIN"), gcell, extra)
    if gn == "B2B": return base + ',%s,"INV",%s,"<>B2C")' % (SR("Document Type Code"), SR("Supply Type Code"))
    if gn == "B2C": return base + ',%s,"INV",%s,"B2C")' % (SR("Document Type Code"), SR("Supply Type Code"))
    if gn == "Credit note": return base + ',%s,"CRN")' % SR("Document Type Code")
    if gn == "Debit note": return base + ',%s,"DBN")' % SR("Document Type Code")
    if gn == "Advance received": return base + ',%s,"Received")' % SR("Adv Bucket (helper)")
    if gn == "Advance adjusted": return base + ',%s,"Adjusted")' % SR("Adv Bucket (helper)")
def g1sum(mi, gcell, gn, extra=""):
    return 'SUMIFS(%s,%s,%s,%s,"%s"%s)' % (G1(G1m[mi]), G1("A"), gcell, G1("J"), gn, extra)
for nm in ["S2 Reco (CA format)", "S2 Month-on-Month", "S2 Exceptions", "S3 1 vs 3B", "S3 Month-on-Month", "S3 SR vs 3B", "S3 Amendment Check"]:
    if nm in wb.sheetnames: del wb[nm]

# ---------- S2 Month-on-Month (long; auto-explain; typed remark) ----------
mm = wb.create_sheet("S2 Month-on-Month")
mm["A1"] = "S2 - Sales Register vs GSTR-1, State | Month | Type (live over SR_2025-26 / GSTR-1 Data). Type remarks in the amber column."
mm["A1"].font = Font(bold=True)
mm.append([]); mm.append(["State", "GSTIN", "Month", "GSTR-1 Type", "Books Taxable", "GSTR-1 Taxable", "Diff Taxable",
                          "Books CGST", "GSTR-1 CGST", "Diff CGST", "Auto: books not in GSTR-1", "Auto: GSTR-1 not in books",
                          "UNEXPLAINED residual", "Remark (type here)", "key"])
for c in mm[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
AUTO2 = {("Bihar", "Jan-26", "Credit note"): "34 no-IRN credit notes in GSTR-1 only - not in books, not in 3B",
         ("Telangana", "May-25", "B2B"): "e-invoiced docs not uploaded to GSTR-1; tax paid via 3B",
         ("Telangana", "Jun-25", "B2B"): "e-invoiced docs not uploaded to GSTR-1; tax paid via 3B",
         ("Telangana", "Aug-25", "B2B"): "e-invoiced docs not uploaded to GSTR-1; tax paid via 3B",
         ("Telangana", "Aug-25", "Credit note"): "e-invoiced docs not uploaded to GSTR-1; tax paid via 3B",
         ("Madhya Pradesh", "Dec-25", "Advance adjusted"): "Rs 9 advance rounding"}
r = 3
for g_, st in GSTINS:
    for mo in MONTHS:
        for gn in GROUPS[:6]:
            r += 1
            mm.cell(r, 1, st); mm.cell(r, 2, g_); mm.cell(r, 3, mo); mm.cell(r, 4, gn)
            mm.cell(r, 5, "=" + books(0, "$B%d" % r, gn, extra=",%s,$C%d" % (SR("Month"), r)))
            mm.cell(r, 6, "=" + g1sum(0, "$B%d" % r, gn, extra=",%s,$C%d" % (G1("K"), r)))
            mm.cell(r, 7, "=E%d-F%d" % (r, r))
            mm.cell(r, 8, "=" + books(2, "$B%d" % r, gn, extra=",%s,$C%d" % (SR("Month"), r)))
            mm.cell(r, 9, "=" + g1sum(2, "$B%d" % r, gn, extra=",%s,$C%d" % (G1("K"), r)))
            mm.cell(r, 10, "=H%d-I%d" % (r, r))
            mm.cell(r, 11, '=SUMIFS(%s,%s,$B%d,%s,$C%d,%s,"Not in GSTR-1")' % (SR("Taxable Value"), SR("My GSTIN"), r, SR("Month"), r, SR("Matched with GSTR-1")))
            mm.cell(r, 12, '=-SUMIFS(%s,%s,$B%d,%s,$C%d,%s,"Not in books")' % (G1("F"), G1("A"), r, G1("K"), r, G1("N")))
            mm.cell(r, 13, "=G%d-K%d-L%d" % (r, r, r))
            mm.cell(r, 14).value = AUTO2.get((st, mo, gn)); mm.cell(r, 14).fill = AMB
            mm.cell(r, 15, '=A%d&"|"&C%d' % (r, r))
            for c in range(5, 14): mm.cell(r, c).number_format = "#,##0.00"
last2 = r
mm.conditional_formatting.add("G4:G%d" % last2, CellIsRule(operator="notBetween", formula=["-1", "1"], fill=RED))
mm.auto_filter.ref = "A3:O%d" % last2
for c_, w in zip("ABCDEFGHIJKLMNO", [18, 18, 9, 14] + [14] * 9 + [42, 4]): mm.column_dimensions[c_].width = w
mm.column_dimensions["O"].hidden = True
mm.freeze_panes = "E4"
def concat2(state_cell):
    terms = []
    for mo in MONTHS:
        # first non-empty remark for that state+month across the 6 type rows: use MATCH on key then scan 6 rows via IFERROR chain is heavy;
        # keys repeat 6x per state-month; INDEX/MATCH returns the FIRST row - remarks are typed on any of the 6; take first match.
        idx = "INDEX('S2 Month-on-Month'!$N$4:$N$%d,MATCH(%s&\"|%s\",'S2 Month-on-Month'!$O$4:$O$%d,0))" % (last2, state_cell, mo, last2)
        terms.append('IFERROR(IF(LEN(%s)>0,"%s: "&%s&"; ",""),"")' % (idx, mo, idx))
    return "=TRIM(" + "&".join(terms) + ")"

# ---------- S2 Reco (CA format) ----------
ws = wb.create_sheet("S2 Reco (CA format)")
ws["A1"] = "Vikran Engineering Limited"; ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "S2 - Sales Register vs GSTR-1 (all cells live over SR_2025-26 / GSTR-1 Data; advances by the sign rule via the register helper column)"
def header(r0, title):
    ws.cell(r0, 1, title).font = Font(bold=True, size=12)
    for gi, gn in enumerate(GROUPS):
        c = 3 + gi * 4; ws.cell(r0 + 1, c, gn).font = Font(bold=True); ws.cell(r0 + 1, c).alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=r0 + 1, start_column=c, end_row=r0 + 1, end_column=c + 3); ws.cell(r0 + 1, c).fill = GF
    ws.cell(r0 + 2, 1, "GSTIN"); ws.cell(r0 + 2, 2, "State")
    for gi in range(7):
        for mi, mn in enumerate(MEAS): ws.cell(r0 + 2, 3 + gi * 4 + mi, mn)
    for c in range(1, 31):
        x = ws.cell(r0 + 2, c); x.font = HF; x.fill = HB; x.alignment = Alignment(horizontal="center", wrap_text=True)
def block(r0, title, kind):
    header(r0, title); first = r0 + 3
    for i2, (g_, st) in enumerate(GSTINS):
        r = first + i2; ws.cell(r, 1, g_); ws.cell(r, 2, st)
        for gi, gn in enumerate(GROUPS):
            for mi in range(4):
                c = 3 + gi * 4 + mi
                if gn == "Total liability": f_ = "=" + "+".join("%s%d" % (L(3 + k * 4 + mi), r) for k in range(6))
                elif kind == "G1": f_ = "=" + g1sum(mi, "$A%d" % r, gn)
                else: f_ = "=" + books(mi, "$A%d" % r, gn)
                ws.cell(r, c, f_); ws.cell(r, c).number_format = "#,##0.00"; ws.cell(r, c).border = BD
        ws.cell(r, 1).border = BD; ws.cell(r, 2).border = BD
    tr = first + len(GSTINS); ws.cell(tr, 1, "Total").font = TOT
    for c in range(3, 31):
        ws.cell(tr, c, "=SUM(%s%d:%s%d)" % (L(c), first, L(c), tr - 1)); ws.cell(tr, c).font = TOT
        ws.cell(tr, c).number_format = "#,##0.00"; ws.cell(tr, c).border = BD
    return first, tr
f1, t1 = block(4, "As per GSTR-1", "G1"); f2, t2 = block(t1 + 2, "Sales Register", "SR")
r0 = t2 + 2; header(r0, "Difference (Sales Register - GSTR-1)"); f3 = r0 + 3
for i2, (g_, st) in enumerate(GSTINS):
    r = f3 + i2; ws.cell(r, 1, g_); ws.cell(r, 2, st)
    for c in range(3, 31):
        ws.cell(r, c, "=%s%d-%s%d" % (L(c), f2 + i2, L(c), f1 + i2)); ws.cell(r, c).number_format = "#,##0.00"; ws.cell(r, c).border = BD
    ws.cell(r, 31, concat2("$B%d" % r)); ws.cell(r, 31).alignment = Alignment(wrap_text=True)
t3 = f3 + len(GSTINS); ws.cell(t3, 1, "Total").font = TOT
for c in range(3, 31):
    ws.cell(t3, c, "=SUM(%s%d:%s%d)" % (L(c), f3, L(c), t3 - 1)); ws.cell(t3, c).font = TOT; ws.cell(t3, c).number_format = "#,##0.00"
ws.cell(f3 - 1, 31, "Remarks (auto from S2 Month-on-Month)").font = HF; ws.cell(f3 - 1, 31).fill = HB
ws.column_dimensions["A"].width = 20; ws.column_dimensions["B"].width = 20
for c in range(3, 31): ws.column_dimensions[L(c)].width = 15
ws.column_dimensions["AE"].width = 50
ws.freeze_panes = "C7"

# ---------- S2 Exceptions (values) ----------
ex = wb.create_sheet("S2 Exceptions")
ex.append(["Status", "State", "My GSTIN", "Books Doc No", "Books type", "Books taxable", "GSTR-1 Doc No", "GSTR-1 type", "GSTR-1 taxable", "Assessment"])
for c in ex[1]: c.font = HF; c.fill = HB
for _, rr in M[M.status != "MATCHED"].sort_values(["status", "state"]).iterrows():
    zero = abs(float(rr["g1_taxable"] or 0)) < 0.005
    stt = "NIL DUPLICATE IN GSTR-1" if (rr["status"] == "IN GSTR-1 ONLY" and zero) else rr["status"]
    a = ("Zero-value second entry of a matched CN - no liability effect; client to confirm why re-reported at nil" if stt == "NIL DUPLICATE IN GSTR-1"
         else "Credit note in GSTR-1 with no IRN; absent from books AND 3B - client to explain" if rr["status"] == "IN GSTR-1 ONLY"
         else "E-invoiced (valid IRN) but never reported in GSTR-1; in 3B, tax paid")
    ex.append([stt, rr["state"], rr["gstin"], rr["docno"], rr["reg_type"], rr["reg_taxable"], rr["g1_docno"], rr["g1_type"], rr["g1_taxable"], a])
for c_, w in zip("ABCDEFGHIJ", [24, 18, 20, 18, 11, 16, 18, 11, 16, 58]): ex.column_dimensions[c_].width = w
ex.freeze_panes = "A2"

# ---------- S3 sheets ----------
mm3 = wb.create_sheet("S3 Month-on-Month")
mm3["A1"] = "S3 - GSTR-1 vs GSTR-3B, State | Month (live). Type remarks in the amber column."; mm3["A1"].font = Font(bold=True)
mm3.append([]); mm3.append(["State", "GSTIN", "Month", "GSTR-1 Taxable", "3B Taxable", "Diff Taxable", "GSTR-1 CGST", "3B CGST", "Diff CGST", "Remark (type here)", "key"])
for c in mm3[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
AUTO3 = {("Bihar", "Aug-25"): "3B taxable mis-keyed (5,73,80,517 as 57,38,017); tax correct - see S3 evidence in Step 3 file",
         ("Bihar", "Jan-26"): "34 no-IRN GSTR-1-only credit notes not taken in 3B",
         ("Telangana", "May-25"): "e-invoiced docs missing from GSTR-1; in 3B, tax paid",
         ("Telangana", "Jun-25"): "e-invoiced docs missing from GSTR-1; in 3B, tax paid",
         ("Telangana", "Aug-25"): "e-invoiced docs missing from GSTR-1; in 3B, tax paid",
         ("Punjab", "May-25"): "Re 1 CGST rounding"}
r = 3
for g_, st in GSTINS:
    for mo in MONTHS:
        r += 1
        mm3.cell(r, 1, st); mm3.cell(r, 2, g_); mm3.cell(r, 3, mo)
        mm3.cell(r, 4, "=SUMIFS(%s,%s,$B%d,%s,$C%d)" % (G1("F"), G1("A"), r, G1("K"), r))
        mm3.cell(r, 5, "=SUMIFS(%s,%s,$B%d,%s,$C%d)" % (B3("D"), B3("B"), r, B3("C"), r))
        mm3.cell(r, 6, "=D%d-E%d" % (r, r))
        mm3.cell(r, 7, "=SUMIFS(%s,%s,$B%d,%s,$C%d)" % (G1("H"), G1("A"), r, G1("K"), r))
        mm3.cell(r, 8, "=SUMIFS(%s,%s,$B%d,%s,$C%d)" % (B3("F"), B3("B"), r, B3("C"), r))
        mm3.cell(r, 9, "=G%d-H%d" % (r, r))
        mm3.cell(r, 10).value = AUTO3.get((st, mo)); mm3.cell(r, 10).fill = AMB
        mm3.cell(r, 11, '=A%d&"|"&C%d' % (r, r))
        for c in range(4, 10): mm3.cell(r, c).number_format = "#,##0.00"
last3 = r
mm3.conditional_formatting.add("F4:F%d" % last3, CellIsRule(operator="notBetween", formula=["-1", "1"], fill=RED))
mm3.auto_filter.ref = "A3:K%d" % last3
for c_, w in zip("ABCDEFGHIJK", [18, 18, 9] + [14] * 6 + [46, 4]): mm3.column_dimensions[c_].width = w
mm3.column_dimensions["K"].hidden = True
mm3.freeze_panes = "D4"
def concat3(state_cell):
    terms = []
    for mo in MONTHS:
        idx = "INDEX('S3 Month-on-Month'!$J$4:$J$%d,MATCH(%s&\"|%s\",'S3 Month-on-Month'!$K$4:$K$%d,0))" % (last3, state_cell, mo, last3)
        terms.append('IFERROR(IF(LEN(%s)>0,"%s: "&%s&"; ",""),"")' % (idx, mo, idx))
    return "=TRIM(" + "&".join(terms) + ")"
g3 = wb.create_sheet("S3 1 vs 3B")
g3["A1"] = "S3 - GSTR-1 vs GSTR-3B 3.1(a), annual by state (live)"; g3["A1"].font = Font(bold=True, size=12)
g3.append([]); g3.append(["GSTIN", "State", "GSTR-1 Taxable", "3B Taxable", "Diff Taxable", "GSTR-1 CGST", "3B CGST", "Diff CGST",
                          "GSTR-1 IGST", "3B IGST", "Diff IGST", "Remarks (auto from S3 Month-on-Month)"])
for c in g3[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
r = 3
for g_, st in GSTINS:
    r += 1
    g3.cell(r, 1, g_); g3.cell(r, 2, st)
    g3.cell(r, 3, "=SUMIFS(%s,%s,$A%d)" % (G1("F"), G1("A"), r)); g3.cell(r, 4, "=SUMIFS(%s,%s,$A%d)" % (B3("D"), B3("B"), r))
    g3.cell(r, 5, "=C%d-D%d" % (r, r))
    g3.cell(r, 6, "=SUMIFS(%s,%s,$A%d)" % (G1("H"), G1("A"), r)); g3.cell(r, 7, "=SUMIFS(%s,%s,$A%d)" % (B3("F"), B3("B"), r))
    g3.cell(r, 8, "=F%d-G%d" % (r, r))
    g3.cell(r, 9, "=SUMIFS(%s,%s,$A%d)" % (G1("G"), G1("A"), r)); g3.cell(r, 10, "=SUMIFS(%s,%s,$A%d)" % (B3("E"), B3("B"), r))
    g3.cell(r, 11, "=I%d-J%d" % (r, r))
    g3.cell(r, 12, concat3("$B%d" % r)); g3.cell(r, 12).alignment = Alignment(wrap_text=True)
    for c in range(3, 12): g3.cell(r, c).number_format = "#,##0.00"
tr3 = r + 1
g3.cell(tr3, 2, "Total").font = TOT
for c in range(3, 12):
    g3.cell(tr3, c, "=SUM(%s4:%s%d)" % (L(c), L(c), r)); g3.cell(tr3, c).font = TOT; g3.cell(tr3, c).number_format = "#,##0.00"
for c_, w in zip("ABCDEFGHIJKL", [20, 18] + [15] * 9 + [52]): g3.column_dimensions[c_].width = w
g3.freeze_panes = "C4"
# SR vs 3B
sv = wb.create_sheet("S3 SR vs 3B")
sv["A1"] = "S3 - Sales Register (books incl advances) vs GSTR-3B 3.1(a) direct, by state (live over SR_2025-26)"
sv["A1"].font = Font(bold=True, size=12)
sv.append([]); sv.append(["GSTIN", "State", "SR Taxable (incl adv)", "3B Taxable", "Diff Taxable", "SR CGST", "3B CGST", "Diff CGST"])
for c in sv[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
r = 3
for g_, st in GSTINS:
    r += 1
    sv.cell(r, 1, g_); sv.cell(r, 2, st)
    sv.cell(r, 3, "=SUMIFS(%s,%s,$A%d)" % (SR("Taxable Value"), SR("My GSTIN"), r))
    sv.cell(r, 4, "=SUMIFS(%s,%s,$A%d)" % (B3("D"), B3("B"), r)); sv.cell(r, 5, "=C%d-D%d" % (r, r))
    sv.cell(r, 6, "=SUMIFS(%s,%s,$A%d)" % (SR("CGST Amount"), SR("My GSTIN"), r))
    sv.cell(r, 7, "=SUMIFS(%s,%s,$A%d)" % (B3("F"), B3("B"), r)); sv.cell(r, 8, "=F%d-G%d" % (r, r))
    for c in range(3, 9): sv.cell(r, c).number_format = "#,##0.00"
trv = r + 1
sv.cell(trv, 2, "Total").font = TOT
for c in range(3, 9):
    sv.cell(trv, c, "=SUM(%s4:%s%d)" % (L(c), L(c), r)); sv.cell(trv, c).font = TOT; sv.cell(trv, c).number_format = "#,##0.00"
for c_, w in zip("ABCDEFGH", [20, 18] + [16] * 6): sv.column_dimensions[c_].width = w
# Amendment check
am = wb.create_sheet("S3 Amendment Check")
am["A1"] = "Amendments"; am["A1"].font = Font(bold=True, size=12)
am["A3"] = "FY 2025-26 GSTR-1 rows flagged 'Is Amendment = Yes'"; am["B3"] = '=COUNTIF(%s,"Yes")' % G1("M")
am["A4"] = "FY 2026-27 GSTR-1 - amendments pointing at FY 25-26 documents"
am["B4"] = "PENDING - 26-27 exports not yet in Portal Reports\\GSTR-1; method: Is Amendment=Yes, join Original/Revised Doc No to SR_2025-26"
am.column_dimensions["A"].width = 60; am.column_dimensions["B"].width = 90
wb.save(OUT)
print("part 1b done: S2 grid rows %d-%d-%d | S2 MoM rows %d | S3 MoM %d | S3 grids added" % (t1, t2, t3, last2 - 3, last3 - 3))
