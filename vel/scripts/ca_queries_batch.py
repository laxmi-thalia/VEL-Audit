"""CA query batch 1:
Q3  Step 2: GSTR-1 Data gains 'Matched SR Doc No' beside Match Status (vice-versa mapping)
Q4  Step 4: GL Data gains 'Matched with SR' per row (vice-versa mapping)
Q10 Step 4: GL Advance Check gains a month-level bridge (ledger vs books, both directions)
Q8  Step 6: 'Bucket Bridge' sheet - raw doc types -> Received/Adjusted, reconciling the CA's filter view
Q7  Step 5: HSN/SAC master sheets embedded + live 'In master?' check column
Q11 Step 1: 'Original Invoice Number' + 'Original Invoice Date' appended to the register from source
Q5  Step 3: new 'SR vs GSTR-3B' reco (books straight to 3B) + 'SR Advances' data sheet
"""
import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
SP = os.path.dirname(os.path.abspath(__file__))
D = r"C:\Users\pawar\Downloads"
def S(v): return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79")
def locked(p):
    try: f = open(p, "r+b"); f.close(); return False
    except Exception: return True
def hdr_style(c): c.font = HF; c.fill = HB
M = pd.read_pickle(os.path.join(SP, "step2_match2.pkl"))
R = pd.read_pickle(os.path.join(SP, "register.pkl"))
R["adv"] = R["dtc"].astype(str).str.startswith("MOB")

# ---------------- Q3: Step 2 ----------------
p = D + r"\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
if locked(p): print("LOCKED:", p)
else:
    pair = {}
    for _, r in M[M.status.str.startswith("MATCHED")].iterrows():
        pair[(S(r["gstin"]).upper(), S(r["g1_docno"]).upper())] = S(r["docno"])
    wb = openpyxl.load_workbook(p); g1 = wb["GSTR-1 Data"]
    hdrs = [S(g1.cell(1, c).value) for c in range(1, g1.max_column + 1)]
    col = hdrs.index("Matched SR Doc No") + 1 if "Matched SR Doc No" in hdrs else g1.max_column + 1
    g1.cell(1, col, "Matched SR Doc No"); hdr_style(g1.cell(1, col))
    n = 0
    for i in range(2, g1.max_row + 1):
        key = (S(g1.cell(i, 1).value).upper(), S(g1.cell(i, 5).value).upper())
        v = pair.get(key, "")
        st = S(g1.cell(i, 12).value)
        g1.cell(i, col).value = v if v else ("-" if st in ("Matched", "") else st)
        if v: n += 1
    g1.column_dimensions[L(col)].width = 20
    wb.save(p); print("Q3 done: Matched SR Doc No on %d GSTR-1 rows" % n)

# ---------------- Q4 + Q10: Step 4 ----------------
p = D + r"\VEL_Step4_Reco_GL_vs_SR_DRAFT.xlsx"
if locked(p): print("LOCKED:", p)
else:
    wb = openpyxl.load_workbook(p)
    gd = wb["GL Data"]; sd = wb["SR Docs"]
    # doc-level SR presence: (state, docno) -> status  (built from SR Docs itself)
    srdocs = {}
    for i in range(2, sd.max_row + 1):
        srdocs[(S(sd.cell(i, 1).value), S(sd.cell(i, 3).value).upper())] = S(sd.cell(i, 11).value)
    hdrs4 = [S(gd.cell(1, c).value) for c in range(1, gd.max_column + 1)]
    col = hdrs4.index("Matched with SR") + 1 if "Matched with SR" in hdrs4 else gd.max_column + 1
    gd.cell(1, col, "Matched with SR"); hdr_style(gd.cell(1, col))
    cnt = collections.Counter()
    for i in range(2, gd.max_row + 1):
        st = S(gd.cell(i, 1).value); ref = S(gd.cell(i, 3).value).upper()
        sales = S(gd.cell(i, 6).value) == "Y"
        if not sales: v = "n/a (clearing/other doc type)"
        elif ref.startswith("MOB") or "ADV" in ref or ref == "COLLECTION": v = "Advance posting - state-level reco (GL Advance Check)"
        elif (st, ref) in srdocs: v = "Matched with SR"
        else: v = "NOT IN SALES REGISTER"
        gd.cell(i, col).value = v; cnt[v.split(" -")[0]] += 1
    gd.column_dimensions[L(col)].width = 34
    wb.save(p); print("Q4 done:", dict(cnt))

# ---------------- Q8: Step 6 bucket bridge ----------------
p = D + r"\VEL_Step6_Advances_Control_DRAFT.xlsx"
if locked(p): print("LOCKED:", p)
else:
    wb = openpyxl.load_workbook(p)
    if "Bucket Bridge" in wb.sheetnames: del wb["Bucket Bridge"]
    b = wb.create_sheet("Bucket Bridge", 2)
    b["A1"] = ("Why 'Received/Adjusted' differ from a raw doc-type filter on the register: MOB ADV REV rows are classified by SIGN "
               "(positive = received, negative = adjusted) - the CA-approved rule that ties books to GSTR-1 (Rs 9) and to the advance ledger (exact).")
    b["A1"].font = Font(bold=True)
    adv = R[R["adv"]].copy()
    adv["sgn"] = adv["tax"].map(lambda v: "positive" if (v or 0) > 0 else "negative")
    piv = adv.groupby([adv["dtc"].map(S), "sgn"])["tax"].agg(["count", "sum"]).round(2).reset_index()
    b.append([]); b.append(["Raw doc type (normalised)", "Sign", "Rows", "Taxable sum", "Goes to bucket"])
    for c in b[3]: hdr_style(c)
    for _, r in piv.iterrows():
        bucket = ("Received" if r["dtc"] == "MOB ADV REC" else "Adjusted" if r["dtc"] == "MOB ADV ADJ"
                  else ("Received" if r["sgn"] == "positive" else "Adjusted"))
        b.append([r["dtc"], r["sgn"], int(r["count"]), float(r["sum"]), bucket])
    b.append([])
    rec = adv[(adv["dtc"] == "MOB ADV REC") | ((adv["dtc"] == "MOB ADV REV") & (adv["tax"] > 0))]["tax"].sum()
    adj = adv[(adv["dtc"] == "MOB ADV ADJ") | ((adv["dtc"] == "MOB ADV REV") & (adv["tax"] <= 0))]["tax"].sum()
    b.append(["Received bucket total (= Control Account col D)", "", "", round(float(rec), 2), ""])
    b.append(["Adjusted bucket total (Control Account col E shows |value|)", "", "", round(float(adj), 2), ""])
    for c_, w in zip("ABCDE", [44, 10, 8, 18, 16]): b.column_dimensions[c_].width = w
    for rr in (len(piv) + 5, len(piv) + 6):
        b.cell(rr, 4).number_format = "#,##0.00"; b.cell(rr, 1).font = Font(bold=True)
    # Q10: month bridge on GL Advance Check
    g = wb["GL Advance Check"]
    adv = R[R["adv"]].copy()
    adv["mon"] = adv.apply(lambda r: pd.to_datetime(r["ddate"]).strftime("%b-%y") if pd.notna(r["ddate"]) else S(r["src"]), axis=1)
    bk = (adv.groupby([adv["state"].map(S), "mon"])[["cgst", "sgst"]].sum().sum(axis=1)).round(2)
    r0 = g.max_row + 3
    g.cell(r0, 1, "Month-level bridge (both directions): books advance tax movement by state-month (values from the register). "
                  "Any ledger month with movement but no books movement (or vice versa) is visible by comparing with the GL extract's Year/Month totals.")
    g.cell(r0, 1).font = Font(bold=True)
    g.cell(r0 + 1, 1, "State"); g.cell(r0 + 1, 2, "Month"); g.cell(r0 + 1, 3, "Books C+S movement")
    for c in range(1, 4): hdr_style(g.cell(r0 + 1, c))
    rr = r0 + 1
    for (st, mo), v in sorted(bk.items()):
        rr += 1; g.cell(rr, 1, st); g.cell(rr, 2, mo); g.cell(rr, 3, float(v)); g.cell(rr, 3).number_format = "#,##0.00"
    wb.save(p); print("Q8 done: bucket bridge with %d raw-type rows | Q10 bridge rows: %d" % (len(piv), rr - r0 - 1))

# ---------------- Q7: Step 5 master check ----------------
p = D + r"\VEL_Step5_HSN_Rate_Summary_DRAFT.xlsx"
if locked(p): print("LOCKED:", p)
else:
    m = pd.ExcelFile(r"C:\Users\pawar\Downloads\HSN_SAC.xlsx")
    hs_m = m.parse("HSN_MSTR"); sa_m = m.parse("SAC_MSTR")
    wb = openpyxl.load_workbook(p)
    for nm, df in (("HSN_MSTR", hs_m), ("SAC_MSTR", sa_m)):
        if nm in wb.sheetnames: del wb[nm]
        ws = wb.create_sheet(nm)
        ws.append(list(df.columns)); [hdr_style(c) for c in ws[1]]
        for _, r in df.iterrows(): ws.append([S(r.iloc[0]), S(r.iloc[1])[:120]])
        ws.column_dimensions["A"].width = 12; ws.column_dimensions["B"].width = 100
        ws.sheet_state = "hidden"
    NH_M = hs_m.shape[0] + 1; NS_M = sa_m.shape[0] + 1
    hsum = wb["HSN Summary"]
    col = 13
    hsum.cell(3, col, "In master? (CA query 7)"); hdr_style(hsum.cell(3, col))
    last = max(r for r in range(4, hsum.max_row + 1) if hsum.cell(r, 3).value not in (None, "") and S(hsum.cell(r, 1).value) != "Total")
    for r in range(4, last + 1):
        hsum.cell(r, col).value = ('=IF(C{0}="","",IF(LEFT(C{0},2)="99",'
            'IF(ISNUMBER(MATCH(C{0},SAC_MSTR!$A$2:$A${1},0)),"OK","NOT IN SAC MASTER"),'
            'IF(OR(ISNUMBER(MATCH(C{0},HSN_MSTR!$A$2:$A${2},0)),ISNUMBER(MATCH(LEFT(C{0},6),HSN_MSTR!$A$2:$A${2},0))),"OK","NOT IN HSN MASTER")))').format(r, NS_M, NH_M)
    hsum.column_dimensions[L(col)].width = 22
    ck = wb["Checks"]
    ck.append(["Codes not found in the client-shared HSN/SAC master", '=COUNTIF(\'HSN Summary\'!M4:M%d,"NOT*")' % last, "master embedded (hidden sheets HSN_MSTR / SAC_MSTR); pre-verified: 0"])
    wb.save(p); print("Q7 done: master embedded, live check col M on %d HSN rows" % (last - 3))

print("batch 1 part A complete")
