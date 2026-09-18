"""Batch-2 step E (last year's formula logic, our sheets):
 - hidden 'LY 24-25 claims' = last year's 'ITC Register 2025-26' rows (FY 24-25 invoices claimed in 25-26; Table 13 / 12C detail)
 - register cols GSTR 9C_Reporting / Reasons / Matching of 12B-12C -> live formulas
 - 'ITC Summary' (79 cols, 19 GSTINs, all SUMIFS; 6A gross blank per Priyesh; 8A from 2B until system-generated GSTR-9)
 - 'Table 13 & 6A1 differences', 'T12B_T12C differences', 'Table 8C vs 13-12 (FY 25-26)'"""
import os, re, warnings, datetime as dt, pickle
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
def S(v): return "" if v is None else str(v).strip()
def norm(s): return re.sub(r"[ \-/.'_]", "", S(s).upper())
SP = os.path.dirname(os.path.abspath(__file__))
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
open(P, "r+b").close()
NF = "#,##0.00"; HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); H2 = PatternFill("solid", fgColor="305496"); H3 = PatternFill("solid", fgColor="8EA9DB")
TOT = Font(bold=True); AMB = PatternFill("solid", fgColor="FFF2CC"); GRY = PatternFill("solid", fgColor="D9D9D9")
thin = Side(style="thin", color="B0B0B0"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
STATES = [("Jammu & Kashmir", "01", "01AAECR0503Q1ZM"), ("Punjab", "03", "03AAECR0503Q1ZI"), ("Haryana", "06", "06AAECR0503Q1ZC"), ("Rajasthan", "08", "08AAECR0503Q1Z8"),
          ("Uttar Pradesh", "09", "09AAECR0503Q1Z6"), ("Bihar", "10", "10AAECR0503Q1ZN"), ("Arunachal Pradesh", "12", "12AAECR0503Q1ZJ"), ("Assam", "18", "18AAECR0503Q1Z7"),
          ("West Bengal", "19", "19AAECR0503Q1Z5"), ("Jharkhand", "20", "20AAECR0503Q1ZM"), ("Chhattisgarh", "22", "22AAECR0503Q1ZI"), ("Madhya Pradesh", "23", "23AAECR0503Q1ZG"),
          ("Gujarat", "24", "24AAECR0503Q1ZE"), ("Maharashtra", "27", "27AAECR0503Q1Z8"), ("Karnataka", "29", "29AAECR0503Q1Z4"), ("Kerala", "32", "32AAECR0503Q1ZH"),
          ("Tamil Nadu", "33", "33AAECR0503Q1ZF"), ("Telangana", "36", "36AAECR0503Q1Z9"), ("Andhra Pradesh", "37", "37AAECR0503Q1Z7")]
wb = openpyxl.load_workbook(P)
# ---------- hidden LY sheet ----------
ly = pickle.load(open(os.path.join(SP, "ly_claims.pkl"), "rb"))
if "LY 24-25 claims" in wb.sheetnames: del wb["LY 24-25 claims"]
ls = wb.create_sheet("LY 24-25 claims"); ls.row_dimensions[1].height = 21
LH = ly["H"] + ["KEY"]
for c, h in enumerate(LH, 1): ls.cell(2, c, h).font = Font(bold=True)
for i, r in enumerate(ly["rows"], 3):
    for c, v in enumerate(r, 1):
        cell = ls.cell(i, c, v)
        if isinstance(v, dt.datetime): cell.number_format = "DD-MM-YYYY"
    ls.cell(i, len(LH), S(r[1]).upper() + norm(r[2]))
NL = 2 + len(ly["rows"]); ls.sheet_state = "hidden"
LY = lambda h: "'LY 24-25 claims'!$%s$3:$%s$%d" % (L(LH.index(h) + 1), L(LH.index(h) + 1), NL)
ls.cell(1, 2, "Last year's 'ITC Register 2025-26' sheet (VEL_GSTR 9_9C FY 24-25.xlsb): FY 24-25 invoices claimed in FY 25-26 = Table 13 (GSTR 9_Reporting=13) and Table 12C (GSTR 9C_Reporting=12C) detail. Tie to the filed GSTR-9/9C PDFs when all 19 are available.").font = Font(italic=True, size=9, color="808080")
print("LY sheet rows:", len(ly["rows"]))
# ---------- register 9C reporting columns ----------
reg = wb["ITC Register 2025-26"]; RH = {S(reg.cell(5, c).value): c for c in range(1, reg.max_column + 1)}
RN = max(r for r in range(6, reg.max_row + 1) if reg.cell(r, 4).value)
rc = lambda h: L(RH[h])
RG = lambda h: "'ITC Register 2025-26'!$%s$6:$%s$%d" % (rc(h), rc(h), RN)
for r in range(6, RN + 1):
    reg.cell(r, RH["GSTR 9C_Reporting"]).value = '=IF($%s%d="2024-25","12B",IF($%s%d="2025-26","NA","Check - dated "&$%s%d))' % (rc("Invoice Year"), r, rc("Invoice Year"), r, rc("Invoice Year"), r)
    reg.cell(r, RH["Reasons"]).value = '=IF($%s%d="12B","ITC booked in 2024-25 claimed in 2025-26",IF($%s%d="NA","ITC booked in 2025-26 claimed in 2025-26","Prior-year dated - time-bar review"))' % (rc("GSTR 9C_Reporting"), r, rc("GSTR 9C_Reporting"), r)
    reg.cell(r, RH["Matching of 12B of FY 25-26 and 12C of 24-25"]).value = '=IF($%s%d<>"12B","NA",IF(COUNTIFS(%s,$%s%d,%s,"12C")>0,"Matched with 12C of FY 24-25","Not in 12C of FY 24-25 - review"))' % (rc("GSTR 9C_Reporting"), r, LY("KEY"), rc("KEY"), r, LY("GSTR 9C_Reporting"))
print("register 9C reporting formulas written")
# ---------- helpers for other sheets ----------
b3 = wb["3B Data"]; B3H = {S(b3.cell(2, c).value): c for c in range(1, b3.max_column + 1)}; N3 = max(r for r in range(3, b3.max_row + 1) if b3.cell(r, 2).value)
def b3col(pre, tail): return L(next(cc for h, cc in B3H.items() if h.startswith(pre) and h.endswith(tail)))
B3 = lambda col: "'3B Data'!$%s$3:$%s$%d" % (col, col, N3)
b2 = wb["GSTR-2B Apr25-Aug26"]; BH = {S(b2.cell(2, c).value): c for c in range(1, b2.max_column + 1)}; NB = max(r for r in range(3, b2.max_row + 1) if b2.cell(r, 1).value)
B2 = lambda h: "'GSTR-2B Apr25-Aug26'!$%s$3:$%s$%d" % (L(BH[h]), L(BH[h]), NB)
ex = wb["T6A1 Extract - 24-25"]; EN = max(r for r in range(5, ex.max_row + 1) if ex.cell(r, 1).value)
EX = lambda col: "'T6A1 Extract - 24-25'!$%s$5:$%s$%d" % (col, col, EN)
tc = wb["Tax comp report"]; TN = max(r for r in range(8, tc.max_row + 1) if tc.cell(r, 2).value)
TC = lambda col: "'Tax comp report'!$%s$8:$%s$%d" % (col, col, TN)
rcm = wb["RCM Register"]; XH = {S(rcm.cell(5, c).value): c for c in range(1, rcm.max_column + 1)}; XN = max(r for r in range(6, rcm.max_row + 1) if rcm.cell(r, XH["MY GSTN"]).value or rcm.cell(r, XH["Business place"]).value)
XR = lambda h: "'RCM Register'!$%s$6:$%s$%d" % (L(XH[h]), L(XH[h]), XN)
HEADS = ("IGST", "CGST", "SGST"); EXCOL = {"IGST": "P", "CGST": "Q", "SGST": "R"}; REGCOL = {"IGST": "IGST", "CGST": "CGST", "SGST": "SGST"}
B2COL = {"IGST": "IGST (Net)", "CGST": "CGST (Net)", "SGST": "SGST (Net)"}
def new_sheet(nm, after):
    if nm in wb.sheetnames: del wb[nm]
    ws = wb.create_sheet(nm, wb.sheetnames.index(after) + 1); ws.row_dimensions[1].height = 21; return ws
# ---------- ITC Summary ----------
ws = new_sheet("ITC Summary", "T6A1 Extract - 24-25")
ws.cell(2, 1, "ITC Master Summary - FY 2025-26 (all live; last year's block logic)").font = TOT
BLOCKS = [  # (group row2, table row3, kind)
    ("Auto-populate (kept BLANK per Priyesh 17-09)", "GSTR 3B Gross ITC 6A", "blank"),
    ("2024-25 Invoices reflecting in 2024-25 2B", "Table 6A1-Claimed", ("ex", "ITC dated 24-25 in 2B of 24-25 availed in 25-26")),
    ("2024-25 Invoices reflecting in 2025-26 2B", "Table 6A1-Claimed", ("ex", "Table 6A1 of GSTR-9 - Claimed")),
    ("2024-25 Invoices_ Correction Entries", "Table 6A1-Claimed", ("ex", "Correction Entries - ITC dated 24-25 booked/reversed in 25-26 (not in 2B)")),
    ("2024-25 Invoices reflecting in 2025-26 2B", "Table 6A1- Unclaimed", ("ex", "Table 6A1 of GSTR-9 - Unclaimed")),
    ("RCM paid in Mar-25 availed in Apr-25", "Table 6A1- RCM paid in Mar-25 availed in Apr-25", ("ex", "RCM paid in Mar-25 availed in Apr-25")),
    ("Total", "ITC Available Table 6A1", "tot6a1"),
    ("4A(5) - 6H - 6A(1) + RCM 6A1", "ITC Available Table 6B", "6b"),
    ("4A(3) GSTR 3B", "ITC Available Table 6D", "6d"),
    ("4A(4) GSTR 3B", "ITC Available Table 6G", "6g"),
    ("4D(1) - 6A(1)", "ITC Reclaimed Table 6H", "6h"),
    ("Difference", "Table 6J", "6j"),
    ("4B(1) + 4B(2)", "ITC Reversal-Table 7H", "7h"),
    ("Formula", "Net ITC available for Utilization-Table 7J", "7j"),
    ("As per Filed GSTR 3B", "Net ITC as per GSTR 3B", "net3b"),
    ("Formula", "Difference", "diff1"),
    ("From GSTR-2B (system-generated GSTR-9 not yet available)", "Total ITC Available Table 8A", "8a"),
    ("Auto captured from T6B of GSTR 9", "Table 8B- ITC availed in 25-26", "8b"),
    ("2025-26 Invoices reflecting in 2026-27 2B", "Table 8C- ITC availed in 26-27", "8c"),
    ("Formula", "Table 8D - Difference (8A-(8B+8C))", "8d"),
    ("GST Portal", "As per Tax Comparison Report", "tcr"),
    ("After considering tax comparison report", "Difference", "diff2"),
    ("Reasons for differences in T8D", "Difference in GSTR-2B Amounts", ("tc", ("V", "W", "X"))),
    ("Reasons for differences in T8D", "(Short)/ Excess reported in 4A5 - Other than CN", ("tc", ("Y", "Z", "AA"))),
    ("Reasons for differences in T8D", "Short/ (Excess) reporting CN in 4A5", ("tc", ("AB", "AC", "AD"))),
]
col = 4; POS = []
for g2, g3, kind in BLOCKS:
    POS.append((kind, col))
    ws.cell(2, col, g2); ws.cell(3, col, g3)
    ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 2); ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 2)
    for j, h in enumerate(HEADS): ws.cell(4, col + j, ("ITC Reversal " + h) if kind == "7h" else h)
    col += 3
    if kind == "diff2": ws.cell(4, col, "Total"); ws.cell(3, col, "Difference"); col += 1
NC = col - 1
for c in range(1, NC + 1):
    for r, fill in ((2, H2), (3, H3), (4, HB)):
        x = ws.cell(r, c); x.font = HF; x.fill = fill; x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); x.border = BD
ws.cell(4, 1, "Locations"); ws.cell(4, 2, "State Code"); ws.cell(4, 3, "GSTN")
ws.row_dimensions[2].height = 30; ws.row_dimensions[3].height = 44
R0 = 5
for i, (st, code, g) in enumerate(STATES):
    r = R0 + i
    ws.cell(r, 1, st); ws.cell(r, 2, code); ws.cell(r, 3, g)
    for kind, c0 in POS:
        for j, h in enumerate(HEADS):
            cell = ws.cell(r, c0 + j)
            if kind == "blank": cell.fill = GRY
            elif isinstance(kind, tuple) and kind[0] == "ex": cell.value = '=SUMIFS(%s,%s,$C%d,%s,"%s")' % (EX(EXCOL[h]), EX("I"), r, EX("G"), kind[1])
            elif isinstance(kind, tuple) and kind[0] == "tc": cell.value = "=SUMIF(%s,$C%d,%s)" % (TC("B"), r, TC(kind[1][j]))
    # explicit positional blocks (the three 'Claimed' blocks share a title)
    c1, c2, c3, c4, c5 = 7, 10, 13, 16, 19       # claimed-PY2B, claimed-CY2B, correction, unclaimed, RCM
    for j, h in enumerate(HEADS):
        a = lambda c0: "%s%d" % (L(c0 + j), r)
        ws.cell(r, 22 + j, "=%s+%s+%s+%s+%s" % (a(c1), a(c2), a(c3), a(c4), a(c5)))                       # Total 6A1
        ws.cell(r, 25 + j, "=SUMIFS(%s,%s,$C%d)-%s-%s+%s" % (B3(b3col("4A(5)", h)), B3("B"), r, a(34), a(22), a(c5)))   # 6B
        ws.cell(r, 28 + j, "=SUMIFS(%s,%s,$C%d)-%s" % (B3(b3col("4A(3)", h)), B3("B"), r, a(c5)))               # 6D
        ws.cell(r, 31 + j, "=SUMIFS(%s,%s,$C%d)" % (B3(b3col("4A(4)", h)), B3("B"), r))                        # 6G
        ws.cell(r, 34 + j, "=SUMIFS(%s,%s,$C%d)-%s-%s" % (B3(b3col("4D(1)", h)), B3("B"), r, a(c1), a(c3)))    # 6H
        ws.cell(r, 37 + j, "=%s-%s-%s-%s-%s-%s" % (a(4), a(22), a(25), a(28), a(31), a(34)))                    # 6J
        ws.cell(r, 40 + j, "=-(SUMIFS(%s,%s,$C%d)+SUMIFS(%s,%s,$C%d))" % (B3(b3col("4B(1)", h)), B3("B"), r, B3(b3col("4B(2)", h)), B3("B"), r))  # 7H (Octa 4B negative -> shown positive)
        ws.cell(r, 43 + j, "=%s+%s+%s+%s-%s" % (a(25), a(28), a(31), a(34), a(40)))                            # 7J
        ws.cell(r, 46 + j, "=SUMIFS(%s,%s,$C%d)" % (B3(b3col("4(C)", h)), B3("B"), r))                         # Net ITC 3B
        ws.cell(r, 49 + j, "=%s-%s+%s" % (a(43), a(46), a(22)))                                                 # Difference
        ws.cell(r, 52 + j, '=SUMIFS(%s,%s,$C%d,%s,"Yes")' % (B2(B2COL[h]), B2("Company GSTIN"), r, B2("Table 8A")))   # 8A
        ws.cell(r, 55 + j, "=%s" % a(25))                                                                       # 8B
        ws.cell(r, 58 + j, '=SUMIFS(%s,%s,$C%d,%s,"Table 8C of GSTR-9")' % (B2(B2COL[h]), B2("Company GSTIN"), r, B2("GSTR-9/9C")))  # 8C
        ws.cell(r, 61 + j, "=%s-%s-%s" % (a(52), a(55), a(58)))                                                 # 8D
        ws.cell(r, 64 + j, "=SUMIF(%s,$C%d,%s)" % (TC("B"), r, TC(("AE", "AF", "AG")[j])))                      # Tax comp difference
        ws.cell(r, 67 + j, "=%s+%s" % (a(61), a(64)))                                                           # Diff after TCR
    ws.cell(r, 70, "=SUM(%s%d:%s%d)" % (L(67), r, L(69), r))
    for c in range(4, NC + 1): ws.cell(r, c).number_format = NF; ws.cell(r, c).border = BD
RT = R0 + len(STATES)
ws.cell(RT, 1, "Total").font = TOT
for c in range(4, NC + 1):
    ws.cell(RT, c, "=SUM(%s%d:%s%d)" % (L(c), R0, L(c), RT - 1)); ws.cell(RT, c).font = TOT; ws.cell(RT, c).number_format = NF; ws.cell(RT, c).border = BD
ws.cell(RT + 2, 1, "Notes: 6A gross (auto-populate) left blank per Priyesh 17-09 - fill from the filed GSTR-9 when available. 7H shown positive (Octa reports 4B negative). 8A from the merged 2B (Table 8A = Yes) until the system-generated GSTR-9 is downloaded. 'As per Tax Comparison Report' = the residual Difference block (AE:AG) of the Tax comp report per GSTIN.").font = Font(italic=True, size=9, color="808080")
ws.column_dimensions["A"].width = 20; ws.column_dimensions["B"].width = 7; ws.column_dimensions["C"].width = 17
for c in range(4, NC + 1): ws.column_dimensions[L(c)].width = 13
ws.freeze_panes = "D5"
print("ITC Summary: %d cols" % NC)
SM = lambda c0, j, r: "'ITC Summary'!%s%d" % (L(c0 + j), r)
# ---------- Table 13 & 6A1 differences ----------
ws = new_sheet("Table 13 & 6A1 differences", "ITC Summary")
ws.cell(2, 1, "VIKRAN ENGINEERING LIMITED").font = TOT; ws.cell(3, 1, "Table 13 (FY 24-25 GSTR-9) vs Table 6A1 (FY 25-26) - live").font = TOT
groups = [(3, "Table 13 - FY 24-25"), (6, "Total Table 6A1 - FY 25-26"), (9, "Table 6A1 - 24-25 in 25-26 - Unclaimed"), (12, "Difference = Total Table 6A1 of FY 25-26 (-) Table 6A1 Unclaimed (-) Table 13 of FY 24-25"),
          (16, "GSTR-9C"), (17, "Remarks"), (19, "24-25 Reflecting in 2025-26 2B- Claim"), (22, "Accounting entries - 24-25 dt in 25-26 (Refer ITC reg 25-26)"), (25, "Amount paid via DRC-03 wrongly reduced from Table 13"), (28, "RCM paid Mar-25 availed Apr-25")]
for c0, t in groups:
    ws.cell(4, c0, t); w = 3 if c0 not in (16, 17, 25) else 1
    if w > 1: ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + w - 1)
hdr6 = {1: "GSTIN", 2: "States", 15: "Total", 16: "Note", 17: "Description", 25: "IGST"}
for c0 in (3, 6, 9, 12, 19, 22, 28):
    for j, h in enumerate(HEADS): hdr6[c0 + j] = h
for c, h in hdr6.items(): ws.cell(6, c, h)
for c in range(1, 31):
    for r in (4, 6):
        x = ws.cell(r, c); x.font = HF; x.fill = HB if r == 6 else H2; x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); x.border = BD
ws.row_dimensions[4].height = 46
for i, (st, code, g) in enumerate(STATES):
    r = 7 + i; sr = R0 + i
    ws.cell(r, 1, g); ws.cell(r, 2, st)
    for j, h in enumerate(HEADS):
        ws.cell(r, 3 + j, '=SUMIFS(%s,%s,$A%d,%s,"13")' % (LY(h), LY("VEL GSTIN"), r, LY("GSTR 9_Reporting")))
        ws.cell(r, 6 + j, "=%s" % SM(22, j, sr)); ws.cell(r, 9 + j, "=%s" % SM(16, j, sr))
        ws.cell(r, 12 + j, "=(%s%d-%s%d)-%s%d" % (L(6 + j), r, L(9 + j), r, L(3 + j), r))
        ws.cell(r, 19 + j, "=%s" % SM(10, j, sr)); ws.cell(r, 22 + j, "=%s" % SM(13, j, sr)); ws.cell(r, 28 + j, "=%s" % SM(19, j, sr))
    ws.cell(r, 15, "=ROUND(SUM(L%d:N%d),0)" % (r, r))
    def chunk(t, n=200): return "&".join('"%s"' % t[i:i + n] for i in range(0, len(t), n))
    A = "The ITC pertaining to invoices of FY 2024-25 has been claimed within the permissible time limit in FY 2025-26. However, it appears that a reporting discrepancy may have occurred while filing GSTR-9 for FY 2024-25, resulting in an under-reporting of ITC in Table 13 amounting to Rs. "
    B = "/-. However, ITC has been correctly availed within the specified time limit."; C = "ITC of Rs. "
    D = "/- reported in Table 13 of GSTR-9 for FY 2024-25 is in excess of the ITC of FY 2024-25 availed in FY 2025-26 as per Table 6A1; the difference is explained by the columns S to Y."
    ws.cell(r, 16, '=IF(ABS(O%d)<1,"",IF(O%d>0,%s&TEXT(O%d,"#,##0")&%s,%s&TEXT(-O%d,"#,##0")&%s))' % (r, r, chunk(A), r, chunk(B), chunk(C), r, chunk(D)))
    ws.cell(r, 17, '=IF(ABS(O%d)<1,"Matched",IF(ABS(O%d-ROUND(SUM(V%d:X%d)+SUM(AB%d:AD%d)+Y%d,0))<1,"Reasons identified","Reasons not identifiable"))' % (r, r, r, r, r, r, r))
    ws.cell(r, 25).fill = AMB
    for c in list(range(3, 16)) + list(range(19, 31)): ws.cell(r, c).number_format = NF
    for c in range(1, 31): ws.cell(r, c).border = BD
RT2 = 7 + len(STATES); ws.cell(RT2, 2, "Total").font = TOT
for c in list(range(3, 16)) + list(range(19, 31)):
    ws.cell(RT2, c, "=SUM(%s7:%s%d)" % (L(c), L(c), RT2 - 1)); ws.cell(RT2, c).font = TOT; ws.cell(RT2, c).number_format = NF
ws.column_dimensions["A"].width = 17; ws.column_dimensions["B"].width = 18; ws.column_dimensions["P"].width = 60; ws.column_dimensions["Q"].width = 22
for c in list(range(3, 16)) + list(range(19, 31)): ws.column_dimensions[L(c)].width = 13
ws.freeze_panes = "C7"
# ---------- T12B_T12C ----------
ws = new_sheet("T12B_T12C differences", "Table 13 & 6A1 differences")
ws.cell(2, 1, "VIKRAN ENGINEERING LIMITED").font = TOT; ws.cell(3, 1, "GSTR-9C Table 12C of FY 24-25 vs Table 12B of FY 25-26 - live").font = TOT
for c, h in enumerate(["GSTIN", "States", "GSTR-9C of PY", "Total GST", "GSTR-9C of CY", "Total GST", "Differences", "Reasons for Differences", "GSTR-9C Notes", "Remarks"], 1):
    x = ws.cell(5, c, h); x.font = HF; x.fill = HB; x.border = BD
for i, (st, code, g) in enumerate(STATES):
    r = 6 + i
    ws.cell(r, 1, g); ws.cell(r, 2, st); ws.cell(r, 3, "Table 12C of FY 2024-25")
    ws.cell(r, 4, '=SUMIFS(%s,%s,$A%d,%s,"12C")' % (LY("Total GST"), LY("VEL GSTIN"), r, LY("GSTR 9C_Reporting")))
    ws.cell(r, 5, "Table 12B of FY 25-26")
    ws.cell(r, 6, '=SUMIFS(%s,%s,$A%d,%s,"12B")' % (RG("Total GST"), RG("VEL GSTIN"), r, RG("GSTR 9C_Reporting")))
    ws.cell(r, 7, "=ROUND(D%d-F%d,0)" % (r, r))
    ws.cell(r, 8, '=IF(G%d=0,"Matched","Review - ITC booked in 24-25 and claimed in 25-26 (12B) differs from last year\'s 12C; see register column \'Matching of 12B of FY 25-26 and 12C of 24-25\'")' % r)
    ws.cell(r, 9, '=IF(G%d=0,"NA","")' % r); ws.cell(r, 10, '=IF(G%d=0,"NA","")' % r)
    for c in (4, 6, 7): ws.cell(r, c).number_format = NF
    for c in range(1, 11): ws.cell(r, c).border = BD
    ws.cell(r, 9).fill = AMB; ws.cell(r, 10).fill = AMB
ws.cell(4, 4, "=SUM(D6:D%d)" % (5 + len(STATES))).number_format = NF; ws.cell(4, 6, "=SUM(F6:F%d)" % (5 + len(STATES))).number_format = NF
for c, w in zip(range(1, 11), [17, 18, 22, 14, 22, 14, 13, 60, 14, 14]): ws.column_dimensions[L(c)].width = w
ws.freeze_panes = "C6"
# ---------- Table 8C vs 13-12 (FY 25-26) ----------
ws = new_sheet("Table 8C vs 13-12 (FY 25-26)", "T12B_T12C differences")
ws.cell(2, 1, "Table 8C (FY 25-26 GSTR-9) vs Net ITC claimed in FY 26-27 (T13 - T12) - PROVISIONAL: T13-T12 needs the FY 26-27 ITC register; until then it mirrors the 2B Table 13 marks.").font = Font(bold=True, color="C00000")
grp = [(2, "Table 8C- ITC availed in 26-27"), (5, "Net ITC claimed T13-T12 (PROVISIONAL)"), (8, "Difference"), (11, "Unclaimed ITC in Table 8C"), (14, "ITC of 25-26 in 2B 25-26"), (17, "Correction entries"), (20, "Not matched with 2B but in ITC reg/ Matched (M) - Considered in Table 13"), (23, "RCM of Mar-26 claimed in Apr-26"), (26, "Net difference"), (29, "Remarks")]
for c0, t in grp:
    ws.cell(3, c0, t)
    if c0 < 29: ws.merge_cells(start_row=3, start_column=c0, end_row=3, end_column=c0 + 2)
ws.cell(4, 1, "GSTN")
for c0, t in grp[:-1]:
    for j, h in enumerate(HEADS): ws.cell(4, c0 + j, h)
for c in range(1, 30):
    for r in (3, 4):
        x = ws.cell(r, c); x.font = HF; x.fill = HB if r == 4 else H2; x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); x.border = BD
ws.row_dimensions[3].height = 46
for i, (st, code, g) in enumerate(STATES):
    r = 5 + i; sr = R0 + i
    ws.cell(r, 1, g)
    for j, h in enumerate(HEADS):
        ws.cell(r, 2 + j, "=%s" % SM(58, j, sr))
        ws.cell(r, 5 + j, '=SUMIFS(%s,%s,$A%d,%s,"<>")' % (B2(B2COL[h]), B2("Company GSTIN"), r, B2("Table 13")))
        ws.cell(r, 8 + j, "=%s%d-%s%d" % (L(5 + j), r, L(2 + j), r))
        for c0 in (11, 14, 17, 20): ws.cell(r, c0 + j, 0).fill = AMB
        ws.cell(r, 23 + j, "=SUMIFS(%s,%s,$A%d,%s,DATE(2026,3,1))" % (XR({"IGST": "IGST AS PER SAP", "CGST": "CGST AS PER SAP", "SGST": "SGST AS PER SAP"}[h]), XR("MY GSTN"), r, XR("Final 3B Month")))
        ws.cell(r, 26 + j, "=%s%d+%s%d-%s%d-%s%d-%s%d-%s%d" % (L(8 + j), r, L(11 + j), r, L(14 + j), r, L(17 + j), r, L(20 + j), r, L(23 + j), r))
    ws.cell(r, 29).fill = AMB
    for c in range(2, 29): ws.cell(r, c).number_format = NF
    for c in range(1, 30): ws.cell(r, c).border = BD
RT3 = 5 + len(STATES); ws.cell(RT3, 1, "Total").font = TOT
for c in range(2, 29): ws.cell(RT3, c, "=SUM(%s5:%s%d)" % (L(c), L(c), RT3 - 1)); ws.cell(RT3, c).font = TOT; ws.cell(RT3, c).number_format = NF
ws.column_dimensions["A"].width = 17
for c in range(2, 29): ws.column_dimensions[L(c)].width = 13
ws.column_dimensions["AC"].width = 30; ws.freeze_panes = "B5"
# ---------- INDEX ----------
ix = wb["INDEX"]; have = {S(ix.cell(rr, 2).value) for rr in range(5, ix.max_row + 1)}
for desc, sheet, lbl in (("ITC Master Summary - Tables 6A1/6B/6D/6G/6H/6J/7H/7J/8A-8D per GSTIN (live)", "ITC Summary", "ITC Summary"),
                         ("Table 13 (FY 24-25) vs Table 6A1 (FY 25-26) differences + 9C note", "Table 13 & 6A1 differences", "Table 13 & 6A1"),
                         ("GSTR-9C Table 12C (FY 24-25) vs Table 12B (FY 25-26)", "T12B_T12C differences", "T12B_T12C"),
                         ("Table 8C (FY 25-26) vs Table 13-12 of FY 26-27 (provisional)", "Table 8C vs 13-12 (FY 25-26)", "8C vs 13-12")):
    if desc in have: continue
    rr = ix.max_row + 1; ix.cell(rr, 1, "ITC"); ix.cell(rr, 2, desc)
    x = ix.cell(rr, 7); x.value = '=HYPERLINK("#\'%s\'!A1","%s")' % (sheet, lbl); x.font = Font(color="0563C1", underline="single")
    for c in range(1, 11): ix.cell(rr, c).border = BD
wb.save(P); print("saved step E")
