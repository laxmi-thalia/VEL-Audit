"""Standalone Step 2 workbook — Sales Register vs GSTR-1, fully formula-driven.
Sheets: Summary (live) | SR vs GSTR-1 (CA 3-block, live SUMIFS) | Month-on-Month (bird's-eye:
state -> 12 months down, Books & GSTR-1 taxable+I/C/S, live) | Exceptions (listing) |
SR Data (register extract) | GSTR-1 Data (portal extract)."""
import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
SP = os.path.dirname(os.path.abspath(__file__))
R = pd.read_pickle(os.path.join(SP, "register.pkl")); G = pd.read_pickle(os.path.join(SP, "gstr1.pkl"))
Su = pd.read_pickle(os.path.join(SP, "g1_summary.pkl")); M = pd.read_pickle(os.path.join(SP, "step2_match2.pkl"))
def s(x): return "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip()
for c_ in ["Taxable Value (Net)", "IGST (Net)", "CGST (Net)", "SGST (Net)"]:
    G[c_] = pd.to_numeric(G[c_], errors="coerce").fillna(0); Su[c_] = pd.to_numeric(Su[c_], errors="coerce").fillna(0)
GSTINS = [("01AAECR0503Q1ZM", "Jammu & Kashmir"), ("03AAECR0503Q1ZI", "Punjab"), ("06AAECR0503Q1ZC", "Haryana"),
 ("08AAECR0503Q1Z8", "Rajasthan"), ("10AAECR0503Q1ZN", "Bihar"), ("12AAECR0503Q1ZJ", "Arunachal Pradesh"),
 ("18AAECR0503Q1Z7", "Assam"), ("19AAECR0503Q1Z5", "West Bengal"), ("20AAECR0503Q1ZM", "Jharkhand"),
 ("23AAECR0503Q1ZG", "Madhya Pradesh"), ("24AAECR0503Q1ZE", "Gujarat"), ("27AAECR0503Q1Z8", "Maharashtra"),
 ("32AAECR0503Q1ZH", "Kerala"), ("33AAECR0503Q1ZF", "TamilNadu"), ("36AAECR0503Q1Z9", "Telangana"),
 ("37AAECR0503Q1Z7", "Andhra Pradesh"), ("09AAECR0503Q1Z6", "Uttar Pradesh"), ("22AAECR0503Q1ZI", "Chhattisgarh"),
 ("29AAECR0503Q1Z4", "Karnataka")]
MONTHS = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26"]
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); GF = PatternFill("solid", fgColor="DDEBF7")
TOT = Font(bold=True); thin = Side(style="thin", color="B0B0B0"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
wb = Workbook()

# ---------- SR Data ----------
sr = wb.active; sr.title = "SR Data"
sr.append(["My GSTIN", "Month", "Document Number", "Doc Type Code", "Supply Type", "Taxable", "IGST", "CGST", "SGST", "Bucket (formula)"])
for c in sr[1]: c.font = HF; c.fill = HB
i = 1
for _, r in R.iterrows():
    i += 1
    mon = pd.to_datetime(r["ddate"]).strftime("%b-%y") if pd.notna(r["ddate"]) else s(r["src"])
    sr.append([s(r["gstin"]), mon, s(r["docno"]), s(r["dtc"]), s(r["sup"]), r["tax"] or 0, r["igst"] or 0, r["cgst"] or 0, r["sgst"] or 0,
               ('=IF(D{n}="INV",IF(E{n}="B2C","B2C","B2B"),IF(D{n}="CRN","Credit note",IF(D{n}="DBN","Debit note",'
                'IF(D{n}="MOB ADV REC","Advance received",IF(D{n}="MOB ADV ADJ","Advance adjusted",'
                'IF(F{n}>0,"Advance received","Advance adjusted"))))))').format(n=i)])
NSR = i
for c_, w in zip("ABCDEFGHIJ", [18, 9, 18, 14, 12, 14, 12, 12, 12, 18]): sr.column_dimensions[c_].width = w
sr.freeze_panes = "A2"
def SR(col): return "'SR Data'!$%s$2:$%s$%d" % (col, col, NSR)
SRm = {0: "F", 1: "G", 2: "H", 3: "I"}

# ---------- GSTR-1 Data ----------
g1 = wb.create_sheet("GSTR-1 Data")
g1.append(["My GSTIN", "Tax Period", "Source", "Raw Type", "Document Number", "Taxable (Net)", "IGST (Net)", "CGST (Net)", "SGST (Net)", "Bucket (formula)", "Month (formula)"])
for c in g1[1]: c.font = HF; c.fill = HB
BF = ('=IF(D{n}="Invoice","B2B",IF(D{n}="Credit Note","Credit note",IF(D{n}="Debit Note","Debit note",'
      'IF(D{n}="B2CS Sales","B2C",IF(D{n}="Advance Received","Advance received","Advance adjusted")))))')
j = 1
for _, r in G.iterrows():
    j += 1; g1.append([s(r["Company GSTIN"]), r["Tax Period"], "Sales-Net (document)", s(r["Doc Type"]), s(r["Doc No"]),
                       r["Taxable Value (Net)"], r["IGST (Net)"], r["CGST (Net)"], r["SGST (Net)"], BF.format(n=j), '=TEXT(B{0},"mmm-yy")'.format(j)])
for _, r in Su.iterrows():
    j += 1; g1.append([s(r["Company GSTIN"]), r["Tax Period"], "SalesSummary-Net (summary)", s(r["Summary Type"]), "",
                       r["Taxable Value (Net)"], r["IGST (Net)"], r["CGST (Net)"], r["SGST (Net)"], BF.format(n=j), '=TEXT(B{0},"mmm-yy")'.format(j)])
NG = j
for c_, w in zip("ABCDEFGHIJK", [18, 12, 24, 16, 18, 15, 12, 12, 12, 18, 12]): g1.column_dimensions[c_].width = w
g1.freeze_panes = "A2"
def G1(col): return "'GSTR-1 Data'!$%s$2:$%s$%d" % (col, col, NG)
G1m = {0: "F", 1: "G", 2: "H", 3: "I"}

# ---------- SR vs GSTR-1 (CA format) ----------
ws = wb.create_sheet("SR vs GSTR-1", 0)
ws["A1"] = "Vikran Engineering Limited"; ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "Sales Register vs GSTR-1"; ws["A2"].font = Font(bold=True, size=12)
ws["A3"] = "FY 2025-26 | every cell is a live SUMIFS over 'SR Data' / 'GSTR-1 Data'. Advance REV rows classify by sign (CA rule): positive = received, negative = adjusted."
GROUPS = ["B2B", "B2C", "Credit note", "Debit note", "Advance received", "Advance adjusted", "Total liability"]
MEAS = ["Taxable value", "IGST", "CGST", "SGST"]
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
    for i2, (g, st) in enumerate(GSTINS):
        r = first + i2; ws.cell(r, 1, g); ws.cell(r, 2, st)
        for gi, gn in enumerate(GROUPS):
            for mi in range(4):
                c = 3 + gi * 4 + mi
                if gn == "Total liability": f_ = "=" + "+".join("%s%d" % (L(3 + k * 4 + mi), r) for k in range(6))
                elif kind == "G1": f_ = '=SUMIFS(%s,%s,$A%d,%s,"%s")' % (G1(G1m[mi]), G1("A"), r, G1("J"), gn)
                else: f_ = '=SUMIFS(%s,%s,$A%d,%s,"%s")' % (SR(SRm[mi]), SR("A"), r, SR("J"), gn)
                ws.cell(r, c, f_); ws.cell(r, c).number_format = "#,##0.00"; ws.cell(r, c).border = BD
        ws.cell(r, 1).border = BD; ws.cell(r, 2).border = BD
    tr = first + len(GSTINS); ws.cell(tr, 1, "Total").font = TOT
    for c in range(3, 31):
        ws.cell(tr, c, "=SUM(%s%d:%s%d)" % (L(c), first, L(c), tr - 1)); ws.cell(tr, c).font = TOT; ws.cell(tr, c).number_format = "#,##0.00"; ws.cell(tr, c).border = BD
    return first, tr
f1, t1 = block(4, "As per GSTR-1", "G1"); f2, t2 = block(t1 + 2, "Sales Register", "SR")
r0 = t2 + 2; header(r0, "Difference (Sales Register - GSTR-1)"); f3 = r0 + 3
for i2, (g, st) in enumerate(GSTINS):
    r = f3 + i2; ws.cell(r, 1, g); ws.cell(r, 2, st)
    for c in range(3, 31):
        ws.cell(r, c, "=%s%d-%s%d" % (L(c), f2 + i2, L(c), f1 + i2)); ws.cell(r, c).number_format = "#,##0.00"; ws.cell(r, c).border = BD
t3 = f3 + len(GSTINS); ws.cell(t3, 1, "Total").font = TOT
for c in range(3, 31):
    ws.cell(t3, c, "=SUM(%s%d:%s%d)" % (L(c), f3, L(c), t3 - 1)); ws.cell(t3, c).font = TOT; ws.cell(t3, c).number_format = "#,##0.00"
ws.column_dimensions["A"].width = 20; ws.column_dimensions["B"].width = 20
for c in range(3, 31): ws.column_dimensions[L(c)].width = 15
ws.freeze_panes = "C7"

# ---------- Month-on-Month (bird's-eye) ----------
mm = wb.create_sheet("Month-on-Month", 1)
mm["A1"] = "Sales Register vs GSTR-1 — month on month, state by state (live). Books by document-date month; GSTR-1 by tax period."; mm["A1"].font = Font(bold=True)
mm.append([]); mm.append(["State", "GSTIN", "Month", "Books Taxable", "Books IGST", "Books CGST", "Books SGST",
                          "GSTR-1 Taxable", "GSTR-1 IGST", "GSTR-1 CGST", "GSTR-1 SGST", "Diff Taxable", "Diff IGST", "Diff CGST", "Diff SGST"])
for c in mm[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
r = 3
for g, st in GSTINS:
    first = r + 1
    for mo in MONTHS:
        r += 1; mm.cell(r, 1, st); mm.cell(r, 2, g); mm.cell(r, 3, mo)
        for mi in range(4):
            mm.cell(r, 4 + mi, "=SUMIFS(%s,%s,$B%d,%s,$C%d)" % (SR(SRm[mi]), SR("A"), r, SR("B"), r))
            mm.cell(r, 8 + mi, "=SUMIFS(%s,%s,$B%d,%s,$C%d)" % (G1(G1m[mi]), G1("A"), r, G1("K"), r))
            mm.cell(r, 12 + mi, "=%s%d-%s%d" % (L(4 + mi), r, L(8 + mi), r))
        for c in range(4, 16): mm.cell(r, c).number_format = "#,##0.00"
    r += 1; mm.cell(r, 1, st + " — Total").font = TOT
    for c in range(4, 16):
        mm.cell(r, c, "=SUM(%s%d:%s%d)" % (L(c), first, L(c), r - 1)); mm.cell(r, c).font = TOT; mm.cell(r, c).number_format = "#,##0.00"; mm.cell(r, c).fill = GF
    r += 1
for c_, w in zip("ABCDEFGHIJKLMNO", [20, 18, 9] + [15] * 12): mm.column_dimensions[c_].width = w
mm.freeze_panes = "D4"

# ---------- Exceptions ----------
ex = wb.create_sheet("Exceptions", 2)
ex.append(["Status", "State", "My GSTIN", "Books Doc No", "Books type", "Books taxable", "GSTR-1 Doc No", "GSTR-1 type", "GSTR-1 taxable", "Assessment"])
for c in ex[1]: c.font = HF; c.fill = HB
for _, rr in M[M.status != "MATCHED"].sort_values(["status", "state"]).iterrows():
    a = ("Credit note declared in GSTR-1 with no IRN; absent from books AND from 3B (see Step 3) - CA to explain"
         if rr["status"] == "IN GSTR-1 ONLY" else "E-invoiced (valid IRN) but never reported in GSTR-1; included in 3B, tax paid (see Step 3)")
    ex.append([rr["status"], rr["state"], rr["gstin"], rr["docno"], rr["reg_type"], rr["reg_taxable"], rr["g1_docno"], rr["g1_type"], rr["g1_taxable"], a])
NE = ex.max_row
for c_, w in zip("ABCDEFGHIJ", [18, 18, 20, 18, 11, 16, 18, 11, 16, 60]): ex.column_dimensions[c_].width = w
ex.freeze_panes = "A2"

# ---------- Summary ----------
su = wb.create_sheet("Summary", 0)
nm = int((M.status == "MATCHED").sum()); by = M[M.status == "MATCHED"]["matched_by"].value_counts().to_dict()
rows = [["STEP 2 — Sales Register vs GSTR-1 (FY 2025-26) — all figures live", "", ""], [],
 ["Total liability — Sales Register", "='SR vs GSTR-1'!AA%d" % t2, ""],
 ["Total liability — GSTR-1", "='SR vs GSTR-1'!AA%d" % t1, ""],
 ["Difference", "='SR vs GSTR-1'!AA%d" % t3, "-> 'SR vs GSTR-1' Difference block shows state and document group"], [],
 ["Documents matched (IRN / GSTIN+DocNo / Customer+Amount)", nm, "IRN %d, GSTIN+DocNo %d, Customer+Amount %d — per-row result is in the register's 'Matched with GSTR-1' column" % (by.get("IRN", 0), by.get("GSTIN+DocNo", 0), by.get("Customer+Amount", 0))],
 ["Documents in books only", '=COUNTIF(Exceptions!A2:A%d,"IN BOOKS ONLY")' % NE, '=SUMIFS(Exceptions!F2:F%d,Exceptions!A2:A%d,"IN BOOKS ONLY")' % (NE, NE)],
 ["Documents in GSTR-1 only", '=COUNTIF(Exceptions!A2:A%d,"IN GSTR-1 ONLY")' % NE, '=SUMIFS(Exceptions!I2:I%d,Exceptions!A2:A%d,"IN GSTR-1 ONLY")' % (NE, NE)],
 ["   listing of every document above", "'Exceptions' sheet — one row per document, with assessment", ""], [],
 ["NOTE", "advances and B2C reconcile at summary level only (no invoice-level entry in GSTR-1) — see 'SR vs GSTR-1' columns S-AB and 'Month-on-Month'", ""]]
for rr in rows: su.append(rr)
su["A1"].font = Font(bold=True, size=12)
for r_ in (3, 4, 5): su.cell(r_, 2).number_format = "#,##0.00"
su.cell(8, 3).number_format = "#,##0.00"; su.cell(9, 3).number_format = "#,##0.00"
su.column_dimensions["A"].width = 56; su.column_dimensions["B"].width = 26; su.column_dimensions["C"].width = 90
OUT = r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
wb.save(OUT); print("WROTE", OUT, "| sheets:", wb.sheetnames)
print("blocks %d-%d / %d-%d / %d-%d | SR Data %d | GSTR-1 Data %d | exceptions %d" % (f1, t1, f2, t2, f3, t3, NSR - 1, NG - 1, NE - 1))
