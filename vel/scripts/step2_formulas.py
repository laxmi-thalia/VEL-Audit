import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
SP = os.path.dirname(os.path.abspath(__file__))
R = pd.read_pickle(os.path.join(SP, "register.pkl"))
G = pd.read_pickle(os.path.join(SP, "gstr1.pkl"))
Su = pd.read_pickle(os.path.join(SP, "g1_summary.pkl"))
M = pd.read_pickle(os.path.join(SP, "step2_match2.pkl"))

def s(x):
    return "" if pd.isna(x) else str(x).strip()

GSTINS = [("01AAECR0503Q1ZM", "Jammu & Kashmir"), ("03AAECR0503Q1ZI", "Punjab"), ("06AAECR0503Q1ZC", "Haryana"),
 ("08AAECR0503Q1Z8", "Rajasthan"), ("10AAECR0503Q1ZN", "Bihar"), ("12AAECR0503Q1ZJ", "Arunachal Pradesh"),
 ("18AAECR0503Q1Z7", "Assam"), ("19AAECR0503Q1Z5", "West Bengal"), ("20AAECR0503Q1ZM", "Jharkhand"),
 ("23AAECR0503Q1ZG", "Madhya Pradesh"), ("24AAECR0503Q1ZE", "Gujarat"), ("27AAECR0503Q1Z8", "Maharashtra"),
 ("32AAECR0503Q1ZH", "Kerala"), ("33AAECR0503Q1ZF", "TamilNadu"), ("36AAECR0503Q1Z9", "Telangana"),
 ("37AAECR0503Q1Z7", "Andhra Pradesh"), ("09AAECR0503Q1Z6", "Uttar Pradesh"), ("22AAECR0503Q1ZI", "Chhattisgarh"),
 ("29AAECR0503Q1Z4", "Karnataka")]
GROUPS = ["B2B", "B2C", "Credit note", "Debit note", "Advance received", "Advance adjusted", "Total liability"]
MEAS = ["Taxable value", "IGST", "CGST", "SGST"]
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79")
GF = PatternFill("solid", fgColor="DDEBF7"); TOT = Font(bold=True)
thin = Side(style="thin", color="B0B0B0"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
wb = Workbook()

# ============ SR Data (register, item level) ============
sr = wb.active; sr.title = "SR Data"
sr.append(["My GSTIN", "Document Date", "Source Month (ISO)", "Document Number", "Doc Type Code", "Supply Type",
           "Taxable", "IGST", "CGST", "SGST", "Bucket (formula)", "GSTR Month (formula)"])
for c in sr[1]:
    c.font = HF; c.fill = HB
srciso = {"Apr-25": "2025-04", "May-25": "2025-05", "Jun-25": "2025-06", "Jul-25": "2025-07", "Aug-25": "2025-08",
          "Sep-25": "2025-09", "Oct-25": "2025-10", "Nov-25": "2025-11", "Dec-25": "2025-12", "Jan-26": "2026-01",
          "Feb-26": "2026-02", "Mar-26": "2026-03"}
i = 1
for _, r in R.iterrows():
    i += 1
    dd = r["ddate"] if pd.notna(r["ddate"]) else None
    sr.append([s(r["gstin"]), dd, srciso.get(s(r["src"]), ""), s(r["docno"]), s(r["dtc"]), s(r["sup"]),
               r["tax"] or 0, r["igst"] or 0, r["cgst"] or 0, r["sgst"] or 0,
               ('=IF(E{n}="INV",IF(F{n}="B2C","B2C","B2B"),IF(E{n}="CRN","Credit note",'
                'IF(E{n}="DBN","Debit note",IF(E{n}="MOB ADV REC","Advance received",'
                'IF(E{n}="MOB ADV ADJ","Advance adjusted",'
                'IF(G{n}>0,"Advance received","Advance adjusted"))))))').format(n=i),
               '=IF(B{n}="",C{n},TEXT(B{n},"yyyy-mm"))'.format(n=i)])
NSR = i
for col, w in zip("ABCDEFGHIJKL", [18, 12, 14, 18, 14, 12, 14, 12, 12, 12, 18, 14]):
    sr.column_dimensions[col].width = w
sr.freeze_panes = "A2"

# ============ GSTR-1 Data (document level + summary level) ============
g1 = wb.create_sheet("GSTR-1 Data")
g1.append(["My GSTIN", "Tax Period", "Source", "Raw Type", "Document Number",
           "Taxable (Net)", "IGST (Net)", "CGST (Net)", "SGST (Net)",
           "Bucket (formula)", "GSTR Month (formula)"])
for c in g1[1]:
    c.font = HF; c.fill = HB
for c_ in ["Taxable Value (Net)", "IGST (Net)", "CGST (Net)", "SGST (Net)"]:
    G[c_] = pd.to_numeric(G[c_], errors="coerce").fillna(0)
    Su[c_] = pd.to_numeric(Su[c_], errors="coerce").fillna(0)
BUCKET_F = ('=IF(D{n}="Invoice","B2B",IF(D{n}="Credit Note","Credit note",IF(D{n}="Debit Note","Debit note",'
            'IF(D{n}="B2CS Sales","B2C",IF(D{n}="Advance Received","Advance received","Advance adjusted")))))')
j = 1
for _, r in G.iterrows():
    j += 1
    g1.append([s(r["Company GSTIN"]), r["Tax Period"], "Sales-Net (document)", s(r["Doc Type"]), s(r["Doc No"]),
               r["Taxable Value (Net)"], r["IGST (Net)"], r["CGST (Net)"], r["SGST (Net)"],
               BUCKET_F.format(n=j), '=TEXT(B{n},"yyyy-mm")'.format(n=j)])
for _, r in Su.iterrows():
    j += 1
    g1.append([s(r["Company GSTIN"]), r["Tax Period"], "SalesSummary-Net (summary)", s(r["Summary Type"]), "",
               r["Taxable Value (Net)"], r["IGST (Net)"], r["CGST (Net)"], r["SGST (Net)"],
               BUCKET_F.format(n=j), '=TEXT(B{n},"yyyy-mm")'.format(n=j)])
NG1 = j
for col, w in zip("ABCDEFGHIJK", [18, 12, 24, 16, 18, 15, 12, 12, 12, 18, 14]):
    g1.column_dimensions[col].width = w
g1.freeze_panes = "A2"

# ============ CA-format sheet, all three blocks LIVE ============
ws = wb.create_sheet("SR vs GSTR-1 (CA format)", 0)
ws["A1"] = "Vikran Engineering Limited"; ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "Sales Register vs GSTR-1"; ws["A2"].font = Font(bold=True, size=12)
ws["A3"] = ("FY 2025-26 | EVERY cell is a live SUMIFS over the 'SR Data' / 'GSTR-1 Data' sheets in this workbook. "
            "Advance REV rows classify by sign (CA rule): positive = received, negative = adjusted.")
MCOL = {0: "F", 1: "G", 2: "H", 3: "I"}   # GSTR-1 Data measure columns
SCOL = {0: "G", 1: "H", 2: "I", 3: "J"}   # SR Data measure columns

def header(r0, title):
    ws.cell(r0, 1, title).font = Font(bold=True, size=12)
    for gi, gn in enumerate(GROUPS):
        c = 3 + gi * 4
        ws.cell(r0 + 1, c, gn).font = Font(bold=True)
        ws.cell(r0 + 1, c).alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=r0 + 1, start_column=c, end_row=r0 + 1, end_column=c + 3)
        ws.cell(r0 + 1, c).fill = GF
    ws.cell(r0 + 2, 1, "GSTIN"); ws.cell(r0 + 2, 2, "State")
    for gi in range(len(GROUPS)):
        for mi, mn in enumerate(MEAS):
            ws.cell(r0 + 2, 3 + gi * 4 + mi, mn)
    for c in range(1, 31):
        x = ws.cell(r0 + 2, c)
        x.font = HF; x.fill = HB
        x.alignment = Alignment(horizontal="center", wrap_text=True)

def block(r0, title, kind):
    header(r0, title); first = r0 + 3
    for i2, (g, st) in enumerate(GSTINS):
        r = first + i2
        ws.cell(r, 1, g); ws.cell(r, 2, st)
        for gi, gn in enumerate(GROUPS):
            for mi in range(4):
                c = 3 + gi * 4 + mi
                if gn == "Total liability":
                    parts = "+".join("{0}{1}".format(get_column_letter(3 + k * 4 + mi), r) for k in range(6))
                    f_ = "=" + parts
                elif kind == "G1":
                    f_ = ("=SUMIFS('GSTR-1 Data'!${m}$2:${m}${N},"
                          "'GSTR-1 Data'!$A$2:$A${N},$A{r},"
                          "'GSTR-1 Data'!$J$2:$J${N},\"{g}\")").format(m=MCOL[mi], N=NG1, r=r, g=gn)
                else:  # SR
                    f_ = ("=SUMIFS('SR Data'!${m}$2:${m}${N},"
                          "'SR Data'!$A$2:$A${N},$A{r},"
                          "'SR Data'!$K$2:$K${N},\"{g}\")").format(m=SCOL[mi], N=NSR, r=r, g=gn)
                ws.cell(r, c, f_)
                ws.cell(r, c).number_format = "#,##0.00"; ws.cell(r, c).border = BD
        ws.cell(r, 1).border = BD; ws.cell(r, 2).border = BD
    tr = first + len(GSTINS)
    ws.cell(tr, 1, "Total").font = TOT
    for c in range(3, 31):
        L = get_column_letter(c)
        ws.cell(tr, c, "=SUM({0}{1}:{0}{2})".format(L, first, tr - 1))
        ws.cell(tr, c).font = TOT; ws.cell(tr, c).number_format = "#,##0.00"; ws.cell(tr, c).border = BD
    return first, tr

f1, t1 = block(4, "As per GSTR-1", "G1")
f2, t2 = block(t1 + 2, "Sales Register", "SR")
r0 = t2 + 2
header(r0, "Difference (Sales Register - GSTR-1)")
f3 = r0 + 3
for i2, (g, st) in enumerate(GSTINS):
    r = f3 + i2
    ws.cell(r, 1, g); ws.cell(r, 2, st)
    for gi in range(len(GROUPS)):
        for mi in range(4):
            c = 3 + gi * 4 + mi; L = get_column_letter(c)
            ws.cell(r, c, "={0}{1}-{0}{2}".format(L, f2 + i2, f1 + i2))
            ws.cell(r, c).number_format = "#,##0.00"; ws.cell(r, c).border = BD
    ws.cell(r, 1).border = BD; ws.cell(r, 2).border = BD
t3 = f3 + len(GSTINS)
ws.cell(t3, 1, "Total").font = TOT
for c in range(3, 31):
    L = get_column_letter(c)
    ws.cell(t3, c, "=SUM({0}{1}:{0}{2})".format(L, f3, t3 - 1))
    ws.cell(t3, c).font = TOT; ws.cell(t3, c).number_format = "#,##0.00"; ws.cell(t3, c).border = BD
ws.column_dimensions["A"].width = 20; ws.column_dimensions["B"].width = 20
for c in range(3, 31):
    ws.column_dimensions[get_column_letter(c)].width = 15
ws.freeze_panes = "C7"

# ============ Month-on-Month, live SUMIFS ============
mm = wb.create_sheet("Month-on-Month")
mm.append(["State", "My GSTIN", "GSTR-1 Month", "Doc Bucket",
           "Taxable Books", "Taxable GSTR-1", "Taxable DIFF", "Note"])
for c in mm[1]:
    c.font = HF; c.fill = HB
piv = pd.read_pickle(os.path.join(SP, "mom2.pkl"))
lab = {g: stt for g, stt in GSTINS}
BK = {"INV": "B2B", "CRN": "Credit note", "DBN": "Debit note"}
k = 1
for _, r in piv.iterrows():
    k += 1
    bucket = BK.get(r["dtc"], r["dtc"])
    if bucket in ("B2B", "B2C", "Credit note", "Debit note"):
        f_books = ("=SUMIFS('SR Data'!$G$2:$G${N},'SR Data'!$A$2:$A${N},$B{k},"
                   "'SR Data'!$L$2:$L${N},$C{k},'SR Data'!$K$2:$K${N},$D{k})").format(N=NSR, k=k)
    else:
        f_books = ("=SUMIFS('SR Data'!$G$2:$G${N},'SR Data'!$A$2:$A${N},$B{k},"
                   "'SR Data'!$L$2:$L${N},$C{k},'SR Data'!$E$2:$E${N},$D{k})").format(N=NSR, k=k)
    f_g1 = ("=SUMIFS('GSTR-1 Data'!$F$2:$F${N},'GSTR-1 Data'!$A$2:$A${N},$B{k},"
            "'GSTR-1 Data'!$K$2:$K${N},$C{k},'GSTR-1 Data'!$J$2:$J${N},$D{k})").format(N=NG1, k=k)
    mm.append([lab.get(s(r["g"]), s(r["g"])), s(r["g"]), r["month"], bucket,
               f_books, f_g1, "=E{0}-F{0}".format(k), r.get("note", "")])
for col, w in zip("ABCDEFGH", [20, 20, 13, 14, 18, 18, 16, 50]):
    mm.column_dimensions[col].width = w
mm.freeze_panes = "A2"

# ============ Matched / Exceptions listings ============
mt = wb.create_sheet("Matched")
mt.append(["State", "My GSTIN", "Books Doc No", "GSTR-1 Doc No", "Type",
           "Books taxable", "GSTR-1 taxable", "Matched by", "Value check (formula)"])
for c in mt[1]:
    c.font = HF; c.fill = HB
k = 1
for _, r in M[M.status.str.startswith("MATCHED")].iterrows():
    k += 1
    mt.append([r["state"], r["gstin"], r["docno"], r["g1_docno"], r["reg_type"],
               r["reg_taxable"], r["g1_taxable"], r["matched_by"],
               '=IF(ABS(F{0}-G{0})<1,"OK","DIFF")'.format(k)])
for col, w in zip("ABCDEFGHI", [20, 20, 20, 20, 8, 16, 16, 14, 16]):
    mt.column_dimensions[col].width = w
mt.freeze_panes = "A2"

ex = wb.create_sheet("Exceptions")
ex.append(["Status", "State", "My GSTIN", "Books Doc No", "Books type", "Books taxable",
           "GSTR-1 Doc No", "GSTR-1 type", "GSTR-1 taxable", "Assessment"])
for c in ex[1]:
    c.font = HF; c.fill = HB
for _, r in M[M.status != "MATCHED"].sort_values(["status", "state"]).iterrows():
    a = ("Credit note declared in GSTR-1, no IRN, absent from books AND from 3B (see Step 3) - CA to explain"
         if r["status"] == "IN GSTR-1 ONLY" else
         "E-invoiced (valid IRN) but never reported in GSTR-1; included in 3B, tax paid (see Step 3)")
    ex.append([r["status"], r["state"], r["gstin"], r["docno"], r["reg_type"], r["reg_taxable"],
               r["g1_docno"], r["g1_type"], r["g1_taxable"], a])
for col, w in zip("ABCDEFGHIJ", [18, 18, 20, 18, 11, 16, 18, 11, 16, 58]):
    ex.column_dimensions[col].width = w
ex.freeze_panes = "A2"

# ============ Reconciliation Summary, live over the listings ============
NM = mt.max_row; NE = ex.max_row
su = wb.create_sheet("Reconciliation Summary", 1)
su.append(["STEP 2 - Sales Register vs GSTR-1 (FY 2025-26) - every figure below is a live formula"])
su["A1"].font = Font(bold=True, size=12)
b_parts = "+".join('SUMIFS(\'SR Data\'!$G$2:$G${N},\'SR Data\'!$K$2:$K${N},"{g}")'.format(N=NSR, g=g_)
                   for g_ in ["B2B", "B2C", "Credit note", "Debit note"])
rows = [[],
 ["Books taxable (excl. advances) - from SR Data", "=" + b_parts],
 ["GSTR-1 taxable (documents, Sales-Net)",
  '=SUMIFS(\'GSTR-1 Data\'!$F$2:$F${N},\'GSTR-1 Data\'!$C$2:$C${N},"Sales-Net (document)")'.format(N=NG1)],
 ["DIFFERENCE", "=B3-B4"],
 [],
 ["Matched documents", "=COUNTA(Matched!A2:A{0})".format(NM)],
 ["   of which value-OK", '=COUNTIF(Matched!I2:I{0},"OK")'.format(NM)],
 ["   matched by IRN", '=COUNTIF(Matched!H2:H{0},"IRN")'.format(NM)],
 ["   matched by GSTIN+DocNo", '=COUNTIF(Matched!H2:H{0},"GSTIN+DocNo")'.format(NM)],
 ["   matched by Customer+Amount", '=COUNTIF(Matched!H2:H{0},"Customer+Amount")'.format(NM)],
 ["In books only (docs / taxable)", '=COUNTIF(Exceptions!A2:A{0},"IN BOOKS ONLY")'.format(NE),
  '=SUMIFS(Exceptions!F2:F{0},Exceptions!A2:A{0},"IN BOOKS ONLY")'.format(NE)],
 ["In GSTR-1 only (docs / taxable)", '=COUNTIF(Exceptions!A2:A{0},"IN GSTR-1 ONLY")'.format(NE),
  '=SUMIFS(Exceptions!I2:I{0},Exceptions!A2:A{0},"IN GSTR-1 ONLY")'.format(NE)],
 [],
 ["NOTE: advances & B2C reconcile at summary level - see the CA-format sheet columns S-AB."],
 ["NOTE: the pairing itself (which GSTR-1 row belongs to which books row) is done by IRN join;"],
 ["      every pair is listed on 'Matched' with both doc numbers and a live value check."],
]
for r in rows:
    su.append(r)
su.column_dimensions["A"].width = 48
su.column_dimensions["B"].width = 24
su.column_dimensions["C"].width = 22

OUT = os.path.join(SP, "VEL_Step2_FORMULA_BUILD.xlsx")
wb.save(OUT)
print("BUILT", OUT)
print("SR Data rows:", NSR - 1, "| GSTR-1 Data rows:", NG1 - 1,
      "| blocks:", f1, "-", t1, ",", f2, "-", t2, ",", f3, "-", t3)
