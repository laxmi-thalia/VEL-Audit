"""MASTER workbook part 1 (CA query 9: one workbook, one source).
Base = copy of the Step 1 register. Adds:
 - SR_2025-26 helper col 'Adv Bucket (helper)' (live; sign rule) so advance SUMIFS read the register
 - 'GSTR-1 Data' (canonical: doc + summary rows, bucket/month formulas, IRN, amendment, match status, matched SR doc)
 - '3B Data', 'GL Data', 'FS Revenue Data' (external sources, values)
 - S2 suite: 'S2 Reco (CA format)' + 'S2 Month-on-Month' + 'S2 Exceptions'  (all SUMIFS -> SR_2025-26)
 - S3 suite: 'S3 1 vs 3B' + 'S3 Month-on-Month' + 'S3 SR vs 3B' + 'S3 Amendment Check'
Every SR-side figure reads SR_2025-26 by header-resolved column letters."""
import os, shutil, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
SP = os.path.dirname(os.path.abspath(__file__))
def S(v): return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); GF = PatternFill("solid", fgColor="DDEBF7")
TOT = Font(bold=True); AMB = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="B0B0B0"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
SRC = r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
OUT = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
shutil.copyfile(SRC, OUT)
G = pd.read_pickle(os.path.join(SP, "gstr1.pkl")); Su = pd.read_pickle(os.path.join(SP, "g1_summary.pkl"))
M = pd.read_pickle(os.path.join(SP, "step2_match2.pkl")); T3 = pd.read_pickle(os.path.join(SP, "b3.pkl"))
for c in ["Taxable Value (Net)", "IGST (Net)", "CGST (Net)", "SGST (Net)"]:
    G[c] = pd.to_numeric(G[c], errors="coerce").fillna(0); Su[c] = pd.to_numeric(Su[c], errors="coerce").fillna(0)
wb = openpyxl.load_workbook(OUT)
ws = wb["SR_2025-26"]
H = {S(ws.cell(4, c).value): L(c) for c in range(1, ws.max_column + 1)}
R0, R1 = 5, 27006
SR = lambda name: "'SR_2025-26'!$%s$%d:$%s$%d" % (H[name], R0, H[name], R1)
# ---- helper col: Adv Bucket
if "Adv Bucket (helper)" not in H:
    c = ws.max_column + 1
    x = ws.cell(4, c, "Adv Bucket (helper)"); x.font = HF; x.fill = HB
    hh, tt = H["Document Type Code"], H["Taxable Value"]
    for i in range(R0, R1 + 1):
        ws.cell(i, c).value = ('=IF(LEFT(' + hh + str(i) + ',3)<>"MOB","",IF(' + hh + str(i) + '="MOB ADV REC","Received",IF(' +
                               hh + str(i) + '="MOB ADV ADJ","Adjusted",IF(' + tt + str(i) + '>0,"Received","Adjusted"))))')
    ws.column_dimensions[L(c)].width = 12
    H["Adv Bucket (helper)"] = L(c)
GSTINS = [("01AAECR0503Q1ZM", "Jammu & Kashmir"), ("03AAECR0503Q1ZI", "Punjab"), ("06AAECR0503Q1ZC", "Haryana"),
 ("08AAECR0503Q1Z8", "Rajasthan"), ("10AAECR0503Q1ZN", "Bihar"), ("12AAECR0503Q1ZJ", "Arunachal Pradesh"),
 ("18AAECR0503Q1Z7", "Assam"), ("19AAECR0503Q1Z5", "West Bengal"), ("20AAECR0503Q1ZM", "Jharkhand"),
 ("23AAECR0503Q1ZG", "Madhya Pradesh"), ("24AAECR0503Q1ZE", "Gujarat"), ("27AAECR0503Q1Z8", "Maharashtra"),
 ("32AAECR0503Q1ZH", "Kerala"), ("33AAECR0503Q1ZF", "TamilNadu"), ("36AAECR0503Q1Z9", "Telangana"),
 ("37AAECR0503Q1Z7", "Andhra Pradesh"), ("09AAECR0503Q1Z6", "Uttar Pradesh"), ("22AAECR0503Q1ZI", "Chhattisgarh"),
 ("29AAECR0503Q1Z4", "Karnataka")]
MONTHS = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26"]
# ---- GSTR-1 Data (canonical)
pair = {}
for _, r in M[M.status.str.startswith("MATCHED")].iterrows():
    pair[(S(r["gstin"]).upper(), S(r["g1_docno"]).upper())] = S(r["docno"])
matched_g1 = set(pair)
only = M[M.status == "IN GSTR-1 ONLY"]
val_only = {(S(r["gstin"]).upper(), S(r["g1_docno"]).upper()) for _, r in only[only["g1_taxable"].abs() >= 0.005].iterrows()}
for nm in ["GSTR-1 Data", "3B Data", "GL Data", "FS Revenue Data"]:
    if nm in wb.sheetnames: del wb[nm]
g1 = wb.create_sheet("GSTR-1 Data")
g1.append(["My GSTIN", "Tax Period", "Source", "Raw Type", "Document Number", "Taxable (Net)", "IGST (Net)", "CGST (Net)", "SGST (Net)",
           "Bucket (formula)", "Month (formula)", "IRN", "Is Amendment", "Match Status", "Matched SR Doc No"])
for c in g1[1]: c.font = HF; c.fill = HB
BF = ('=IF(D{n}="Invoice","B2B",IF(D{n}="Credit Note","Credit note",IF(D{n}="Debit Note","Debit note",'
      'IF(D{n}="B2CS Sales","B2C",IF(D{n}="Advance Received","Advance received","Advance adjusted")))))')
j = 1
for _, r in G.iterrows():
    j += 1
    key = (S(r["Company GSTIN"]).upper(), S(r["Doc No"]).upper())
    irn = S(r["IRN"]) if S(r["IRN"]).lower() != "nan" else ""
    tv = float(r["Taxable Value (Net)"])
    stt = ("Matched" if key in matched_g1 else "Not in books" if key in val_only
           else "Nil duplicate of a matched CN" if abs(tv) < 0.005 and not irn and S(r["Doc Type"]) == "Credit Note" else "Matched")
    g1.append([S(r["Company GSTIN"]), r["Tax Period"], "Sales-Net (document)", S(r["Doc Type"]), S(r["Doc No"]),
               tv, r["IGST (Net)"], r["CGST (Net)"], r["SGST (Net)"], BF.format(n=j), '=TEXT(B{0},"mmm-yy")'.format(j),
               irn, S(r["Is Amendment"]), stt, pair.get(key, "")])
for _, r in Su.iterrows():
    j += 1
    g1.append([S(r["Company GSTIN"]), r["Tax Period"], "SalesSummary-Net (summary)", S(r["Summary Type"]), "",
               r["Taxable Value (Net)"], r["IGST (Net)"], r["CGST (Net)"], r["SGST (Net)"], BF.format(n=j),
               '=TEXT(B{0},"mmm-yy")'.format(j), "", S(r.get("Is Amendment", "")), "Summary level", ""])
NG = j
for c_, w in zip("ABCDEFGHIJKLMNO", [18, 12, 24, 16, 18, 15, 12, 12, 12, 18, 12, 20, 11, 22, 18]): g1.column_dimensions[c_].width = w
g1.freeze_panes = "A2"
G1 = lambda c: "'GSTR-1 Data'!$%s$2:$%s$%d" % (c, c, NG)
G1m = {0: "F", 1: "G", 2: "H", 3: "I"}
# ---- 3B Data
b3 = wb.create_sheet("3B Data")
b3.append(["State", "GSTIN", "Month", "Taxable 3.1(a)", "IGST", "CGST", "SGST"])
for c in b3[1]: c.font = HF; c.fill = HB
T3["g"] = T3["gstin"].map(S)
YM = {m: pd.to_datetime(m, format="%b-%y").strftime("%b %Y") for m in MONTHS}
for g_, st in GSTINS:
    row = T3[T3["g"] == g_]
    for m in MONTHS:
        k = YM[m]
        vals = [float(row.iloc[0].get("%s|%s" % (k, x), 0) or 0) if len(row) else 0.0 for x in ("taxable", "igst", "cgst", "sgst")]
        b3.append([st, g_, m] + [round(v, 2) for v in vals])
N3 = b3.max_row
for c_, w in zip("ABCDEFG", [20, 18, 9, 16, 14, 14, 14]): b3.column_dimensions[c_].width = w
B3 = lambda c: "'3B Data'!$%s$2:$%s$%d" % (c, c, N3)
B3m = {0: "D", 1: "E", 2: "F", 3: "G"}
# ---- GL Data + FS Revenue Data: copy values from the step files
for src_p, src_sh, new_nm in [(r"C:\Users\pawar\Downloads\VEL_Step4_Reco_GL_vs_SR_DRAFT.xlsx", "GL Data", "GL Data"),
                              (r"C:\Users\pawar\Downloads\VEL_Step9_FS_vs_GST_DRAFT.xlsx", "FS Revenue Data", "FS Revenue Data")]:
    sw = openpyxl.load_workbook(src_p, read_only=True)
    ss = sw[src_sh]
    ds = wb.create_sheet(new_nm)
    for row in ss.iter_rows(values_only=True):
        ds.append(list(row))
    sw.close()
    for c in ds[1]: c.font = HF; c.fill = HB
    ds.freeze_panes = "A2"
NGL = wb["GL Data"].max_row; NFS = wb["FS Revenue Data"].max_row
wb.save(OUT)
print("part 1a done: base + helper + data sheets | GSTR-1 rows %d | 3B rows %d | GL rows %d | FS rows %d" % (NG - 1, N3 - 1, NGL - 1, NFS - 1))
