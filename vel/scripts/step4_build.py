"""Step 4 — GL vs Sales Register (output tax), formula-driven with evidence.
Sheets: Summary | GL vs Sales Register (state x C/S/I, live SUMIFS, auto-explanation + remarks) |
Month-on-Month (long, live) | Exceptions (document-level evidence) | Remarks | GL Data | SR Docs
Source: Clients Data\\GLs\\Outward Tax\\GST Output.xlsx  (ledgers 2610080100=CGST, ...101=SGST, ...102=IGST)
Scope: Year/Month 2025/01-12 (FY 25-26); sales-origin document types RV, DA, DG, DR, XD.
Sign: ledger credits are negative; amounts are flipped so positive = output tax liability."""
import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
SP = os.path.dirname(os.path.abspath(__file__))
GLP = "//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/GLs/Outward Tax/GST Output.xlsx"
def S(v): return "" if pd.isna(v) else str(v).strip()
def num(v):
    v = S(v)
    try: return str(int(float(v)))
    except Exception: return v
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); GF = PatternFill("solid", fgColor="DDEBF7")
TOT = Font(bold=True); thin = Side(style="thin", color="B0B0B0"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
RED = PatternFill("solid", fgColor="FFC7CE")
BP2STATE = {"AP01": "Andhra Pradesh", "AR01": "Arunachal Pradesh", "AS01": "Assam", "BR01": "Bihar", "CG01": "Chhattisgarh",
            "GU01": "Gujarat", "HR01": "Haryana", "JH01": "Jharkhand", "JK01": "Jammu & Kashmir", "KA01": "Karnataka",
            "KL01": "Kerala", "MH01": "Maharashtra", "MP01": "Madhya Pradesh", "PB01": "Punjab", "RJ01": "Rajasthan",
            "TG01": "Telangana", "TN01": "Tamil Nadu", "UP01": "Uttar Pradesh", "WB01": "West Bengal"}
HEAD = {"2610080100": "CGST", "2610080101": "SGST", "2610080102": "IGST"}
SALES = {"RV", "DA", "DG", "DR", "XD"}
YM2M = {"2025/%02d" % m: pd.Timestamp(2025 if m <= 9 else 2026, (m + 3) if m <= 9 else (m - 9), 1).strftime("%b-%y") for m in range(1, 13)}
MONTHS = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26"]

# ---------------- load GL ----------------
d = pd.read_excel(GLP, sheet_name="Data", header=5)
d.columns = [S(c) for c in d.columns]
d = d[d["Document Number"].notna()].copy()
d["amt"] = -pd.to_numeric(d["Amount in Local Currency"], errors="coerce").fillna(0)   # + = liability
d["gl"] = d["G/L Account"].map(num); d["head"] = d["gl"].map(HEAD)
d["dt"] = d["Document Type"].map(S); d["ym"] = d["Year/Month"].map(S)
d["bp"] = d["Business place"].map(S); d["state"] = d["bp"].map(BP2STATE)
d["ref"] = d["Reference"].map(S).str.upper()
d["mon"] = d["ym"].map(YM2M)
d["in_fy"] = d["ym"].isin(set(YM2M))
d["sales_origin"] = d["dt"].isin(SALES)
f = d[d["in_fy"]].copy()

# ---------------- register doc-level ----------------
R = pd.read_pickle(os.path.join(SP, "register.pkl"))
R["adv"] = R["dtc"].astype(str).str.startswith("MOB")
SD = R.groupby(["gstin", "docno", "dtc"], dropna=False).agg(
    state=("state", "first"), tax=("tax", "sum"), igst=("igst", "sum"), cgst=("cgst", "sum"), sgst=("sgst", "sum"),
    adv=("adv", "first"), ddate=("ddate", "first"), src=("src", "first")).reset_index()
SD["mon"] = SD.apply(lambda r: pd.to_datetime(r["ddate"]).strftime("%b-%y") if pd.notna(r["ddate"]) else S(r["src"]), axis=1)
SD["ref"] = SD["docno"].map(S).str.upper()

# ---------------- comparison: invoices at document level, advances at state level ----------------
import re as _re
reg_nonadv_refs = {r for r in SD[~SD["adv"]]["ref"]}
def gl_class(ref):
    if ref in reg_nonadv_refs: return "DOC"
    if ref.startswith("MOB") or "ADV" in ref or ref in ("COLLECTION",): return "ADVANCE"
    return "GL_ONLY_DOC"
fs = f[f["sales_origin"]].copy()
fs["cls"] = fs["ref"].map(gl_class)
piv = fs.pivot_table(index=["state", "ref", "cls"], columns="head", values="amt", aggfunc="sum", fill_value=0).reset_index()
for h in ("CGST", "SGST", "IGST"):
    if h not in piv: piv[h] = 0.0
glmap = {(S(r["state"]), r["ref"]): (r["CGST"], r["SGST"], r["IGST"]) for _, r in piv[piv["cls"] == "DOC"].iterrows()}
status = {}; exc = []   # exc rows: (status, state, ref/label, type, month, reg CGST, GL CGST, contribution GL-SR)
reg_doc = SD[~SD["adv"]]
for _, r in reg_doc.iterrows():
    k = (S(r["state"]), r["ref"])
    g = glmap.get(k)
    key = (S(r["gstin"]).upper(), r["ref"], S(r["dtc"]))
    if g is None:
        status[key] = "Not in GL"
        exc.append(("IN REGISTER, NOT IN GL", S(r["state"]), r["ref"], S(r["dtc"]), r["mon"], round(r["cgst"], 2), 0.0, round(-r["cgst"], 2)))
    elif abs(round(r["cgst"] - g[0], 2)) < 1:
        status[key] = "Matched (GL)"
    else:
        status[key] = "Value differs (GL)"
        exc.append(("VALUE DIFFERS", S(r["state"]), r["ref"], S(r["dtc"]), r["mon"], round(r["cgst"], 2), round(g[0], 2), round(g[0] - r["cgst"], 2)))
for _, r in piv[piv["cls"] == "GL_ONLY_DOC"].iterrows():
    if abs(r["CGST"]) >= 1 or abs(r["IGST"]) >= 1:
        exc.append(("IN GL, NOT IN REGISTER", S(r["state"]), r["ref"], "", "", 0.0, round(r["CGST"], 2), round(r["CGST"], 2)))
# advance component per state
gl_adv = fs[fs["cls"] == "ADVANCE"].pivot_table(index="state", columns="head", values="amt", aggfunc="sum", fill_value=0)
reg_adv = SD[SD["adv"]].groupby("state")["cgst"].sum()
for st_ in sorted(set(gl_adv.index) | set(S(x) for x in reg_adv.index)):
    g = float(gl_adv["CGST"].get(st_, 0.0)) if "CGST" in gl_adv else 0.0
    rr = float(reg_adv.get(st_, 0.0))
    if abs(round(g - rr, 2)) >= 1:
        exc.append(("ADVANCE COMPONENT (state level)", st_, "advances: GL MOB/COLLECTION refs vs register MOB rows", "", "", round(rr, 2), round(g, 2), round(g - rr, 2)))
exc.sort(key=lambda x: (x[0], x[1], -abs(x[7])))
print("doc statuses:", collections.Counter(status.values()))
print("exceptions:", collections.Counter(e[0] for e in exc))
# ---------------- workbook ----------------
wb = Workbook()
# GL Data
gd = wb.active; gd.title = "GL Data"
gd.append(["State", "Business place", "Reference (invoice no)", "FI Document", "Doc Type", "Sales-origin?", "Year/Month", "Month",
           "Tax head", "Amount (+ = liability)", "Posting Date", "Clearing Document", "Text"])
for c in gd[1]: c.font = HF; c.fill = HB
i = 1
for _, r in f.iterrows():
    i += 1
    gd.append([S(r["state"]) or S(r["bp"]), r["bp"], r["ref"], num(r["Document Number"]), r["dt"], "Y" if r["sales_origin"] else "N",
               r["ym"], r["mon"], r["head"], round(r["amt"], 2), r["Posting Date"], num(r["Clearing Document"]), S(r["Text"])])
NG = i
for c_, w in zip("ABCDEFGHIJKLM", [18, 12, 18, 14, 9, 11, 10, 9, 9, 16, 12, 14, 24]): gd.column_dimensions[c_].width = w
gd.freeze_panes = "A2"
GD = lambda c: "'GL Data'!$%s$2:$%s$%d" % (c, c, NG)
# SR Docs
sd = wb.create_sheet("SR Docs")
sd.append(["State", "My GSTIN", "Document Number", "Doc Type", "Month", "Advance?", "Taxable", "IGST", "CGST", "SGST", "GL match status"])
for c in sd[1]: c.font = HF; c.fill = HB
j = 1
for _, r in SD.iterrows():
    j += 1
    st_ = status.get((S(r["gstin"]).upper(), r["ref"], S(r["dtc"])), "Advance - see GST Advance ledger" if r["adv"] else "")
    sd.append([S(r["state"]), S(r["gstin"]), r["ref"], S(r["dtc"]), r["mon"], "Y" if r["adv"] else "N",
               round(r["tax"], 2), round(r["igst"], 2), round(r["cgst"], 2), round(r["sgst"], 2), st_])
ND = j
for c_, w in zip("ABCDEFGHIJK", [18, 18, 18, 12, 9, 9, 16, 12, 14, 14, 26]): sd.column_dimensions[c_].width = w
sd.freeze_panes = "A2"
SDr = lambda c: "'SR Docs'!$%s$2:$%s$%d" % (c, c, ND)

# Remarks
rm = wb.create_sheet("Remarks")
rm.append(["State", "Month", "Remark", "Key (auto)"])
for c in rm[1]: c.font = HF; c.fill = HB
AUTO = [("Gujarat", "All", "AUTO: GL output higher than register - see Exceptions for the documents"),
        ("Rajasthan", "All", "AUTO: GL output higher than register - see Exceptions for the documents"),
        ("Madhya Pradesh", "All", "AUTO: GL output higher than register - see Exceptions for the documents")]
for k, (st_, mo, txt) in enumerate(AUTO, start=2): rm.append([st_, mo, txt, '=A%d&"|"&B%d' % (k, k)])
for k in range(len(AUTO) + 2, len(AUTO) + 42): rm.append(["", "", "", '=A%d&"|"&B%d' % (k, k)])
NR = rm.max_row
for c_, w in zip("ABCD", [20, 10, 100, 24]): rm.column_dimensions[c_].width = w
REM = lambda key: '=IFERROR(INDEX(Remarks!$C$2:$C$%d,MATCH(%s,Remarks!$D$2:$D$%d,0)),"")' % (NR, key, NR)

# GL vs Sales Register (CA-style state grid)
ws = wb.create_sheet("GL vs Sales Register", 0)
ws["A1"] = "Vikran Engineering Limited"; ws["A1"].font = Font(bold=True, size=13)
ws["A2"] = "Output GL vs Sales Register — FY 2025-26"; ws["A2"].font = Font(bold=True, size=12)
ws["A3"] = ("GL = 'GST Output.xlsx' ledgers 2610080100/101/102 (CGST/SGST/IGST), Year/Month 2025/01-12, sales-origin doc types RV/DA/DG/DR/XD, "
            "credits sign-flipped. Every figure is a live SUMIFS over 'GL Data' / 'SR Docs'.")
GRP = [("As per GL (output ledgers)", "GL"), ("As per Sales Register", "SR"), ("Difference (GL - SR)", "D")]
MEAS = ["CGST", "SGST", "IGST", "Total"]
ws.cell(5, 1, "Business place"); ws.cell(5, 2, "State")
for gi, (gn, _) in enumerate(GRP):
    c = 3 + gi * 4
    ws.cell(4, c, gn).font = Font(bold=True); ws.cell(4, c).alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 3); ws.cell(4, c).fill = GF
    for mi, mn in enumerate(MEAS): ws.cell(5, c + mi, mn)
xcols = ["Auto: EXPLAINED (sum of Exceptions contributions)", "UNEXPLAINED residual", "Remark"]
for j, t in enumerate(xcols, start=15):
    ws.cell(5, j, t)
for c in range(1, 20):
    x = ws.cell(5, c); x.font = HF; x.fill = HB; x.alignment = Alignment(horizontal="center", wrap_text=True)
first = 6
STATES = sorted({S(x) for x in SD["state"]} | {v for k, v in BP2STATE.items() if k in set(f["bp"])})
for i2, st_ in enumerate(STATES):
    r = first + i2
    bp = next((k for k, v in BP2STATE.items() if v == st_), "")
    ws.cell(r, 1, bp); ws.cell(r, 2, st_)
    for mi, hd in enumerate(["CGST", "SGST", "IGST"]):
        ws.cell(r, 3 + mi, '=SUMIFS(%s,%s,$B%d,%s,"%s",%s,"Y")' % (GD("J"), GD("A"), r, GD("I"), hd, GD("F")))
        ws.cell(r, 7 + mi, "=SUMIFS(%s,%s,$B%d)" % (SDr({"CGST": "I", "SGST": "J", "IGST": "H"}[hd]), SDr("A"), r))
        ws.cell(r, 11 + mi, "=%s%d-%s%d" % (L(3 + mi), r, L(7 + mi), r))
    for base in (3, 7, 11):
        ws.cell(r, base + 3, "=SUM(%s%d:%s%d)" % (L(base), r, L(base + 2), r))
    ws.cell(r, 17, REM('$B%d&"|All"' % r))
    for c in range(3, 17): ws.cell(r, c).number_format = "#,##0.00"; ws.cell(r, c).border = BD
tr = first + len(STATES)
ws.cell(tr, 2, "Total").font = TOT
for c in range(3, 17):
    ws.cell(tr, c, "=SUM(%s%d:%s%d)" % (L(c), first, L(c), tr - 1)); ws.cell(tr, c).font = TOT
    ws.cell(tr, c).number_format = "#,##0.00"; ws.cell(tr, c).border = BD
ws.column_dimensions["A"].width = 13; ws.column_dimensions["B"].width = 20
for c in range(3, 19): ws.column_dimensions[L(c)].width = 15
ws.column_dimensions["Q"].width = 46
ws.freeze_panes = "C6"

# fix the explanation columns properly:
#  O = register docs not in GL (CGST, sum from SR Docs status)     -> contributes NEGATIVE to (GL-SR) diff
#  P = GL docs not in register (CGST, from Exceptions listing)     -> contributes POSITIVE
#  Q = value differences (CGST, from Exceptions listing)           -> contributes POSITIVE
#  R = residual = K - (P + Q - O)
# Exceptions sheet first, then P/Q reference it.
ex = wb.create_sheet("Exceptions", 1)
ex.append(["Status", "State", "Reference / Doc No", "Reg type", "Month", "Register CGST", "GL CGST", "Contribution to (GL - SR)"])
for c in ex[1]: c.font = HF; c.fill = HB
for e in exc: ex.append(list(e))
NE = ex.max_row
for c_, w in zip("ABCDEFGH", [24, 18, 20, 10, 9, 16, 16, 20]): ex.column_dimensions[c_].width = w
ex.freeze_panes = "A2"
for i2, st_ in enumerate(STATES):
    r = first + i2
    ws.cell(r, 15, '=SUMIFS(Exceptions!$H$2:$H$%d,Exceptions!$B$2:$B$%d,$B%d)' % (NE, NE, r))
    ws.cell(r, 16, "=K%d-O%d" % (r, r))
    for c in (15, 16): ws.cell(r, c).number_format = "#,##0.00"

# Month-on-Month (long, CGST focus + IGST)
mm = wb.create_sheet("Month-on-Month", 2)
mm["A1"] = "Output GL vs Sales Register — State | Month (live). Difference = GL minus register."; mm["A1"].font = Font(bold=True)
mm.append([]); mm.append(["State", "Month", "GL CGST", "SR CGST", "Diff CGST", "GL IGST", "SR IGST", "Diff IGST", "Remark (from 'Remarks')"])
for c in mm[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
r = 3
for st_ in STATES:
    for mo in MONTHS:
        r += 1
        mm.cell(r, 1, st_); mm.cell(r, 2, mo)
        mm.cell(r, 3, '=SUMIFS(%s,%s,$A%d,%s,$B%d,%s,"CGST",%s,"Y")' % (GD("J"), GD("A"), r, GD("H"), r, GD("I"), GD("F")))
        mm.cell(r, 4, '=SUMIFS(%s,%s,$A%d,%s,$B%d)' % (SDr("I"), SDr("A"), r, SDr("E"), r))
        mm.cell(r, 5, "=C%d-D%d" % (r, r))
        mm.cell(r, 6, '=SUMIFS(%s,%s,$A%d,%s,$B%d,%s,"IGST",%s,"Y")' % (GD("J"), GD("A"), r, GD("H"), r, GD("I"), GD("F")))
        mm.cell(r, 7, '=SUMIFS(%s,%s,$A%d,%s,$B%d)' % (SDr("H"), SDr("A"), r, SDr("E"), r))
        mm.cell(r, 8, "=F%d-G%d" % (r, r))
        mm.cell(r, 9, REM('$A%d&"|"&$B%d' % (r, r)))
        for c in range(3, 9): mm.cell(r, c).number_format = "#,##0.00"
from openpyxl.formatting.rule import CellIsRule
mm.conditional_formatting.add("E4:E%d" % r, CellIsRule(operator="notBetween", formula=["-1", "1"], fill=RED))
mm.auto_filter.ref = "A3:I%d" % r
for c_, w in zip("ABCDEFGHI", [20, 9] + [15] * 6 + [50]): mm.column_dimensions[c_].width = w
mm.freeze_panes = "C4"

# Summary
su = wb.create_sheet("Summary", 0)
rows = [["STEP 4 — Output GL vs Sales Register (FY 2025-26) — all figures live", ""], [],
 ["GL output (CGST, sales-origin, FY)", "='GL vs Sales Register'!C%d" % tr],
 ["Register CGST", "='GL vs Sales Register'!G%d" % tr],
 ["Difference (CGST)", "='GL vs Sales Register'!K%d" % tr],
 ["   explained by the Exceptions listing", "='GL vs Sales Register'!O%d" % tr],
 ["   UNEXPLAINED residual", "='GL vs Sales Register'!P%d" % tr], [],
 ["Register documents matched to GL", '=COUNTIF(%s,"Matched (GL)")' % SDr("K")],
 ["   value differs", '=COUNTIF(%s,"Value differs (GL)")' % SDr("K")],
 ["   not in GL", '=COUNTIF(%s,"Not in GL")' % SDr("K")],
 ["   advances (see GST Advance ledger)", '=COUNTIF(%s,"Advance*")' % SDr("K")], [],
 ["Evidence", "'Exceptions' lists every document behind the difference; 'GL Data' / 'SR Docs' are the sources the formulas read"]]
for rr in rows: su.append(rr)
su["A1"].font = Font(bold=True, size=12)
for r_ in range(3, 10): su.cell(r_, 2).number_format = "#,##0.00"
su.column_dimensions["A"].width = 52; su.column_dimensions["B"].width = 90
OUT = r"C:\Users\pawar\Downloads\VEL_Step4_Reco_GL_vs_SR_DRAFT.xlsx"
wb.save(OUT)
print("WROTE", OUT)
print("GL Data rows:", NG - 1, "| SR Docs:", ND - 1, "| exceptions:", NE - 1, "| states:", len(STATES), "| grid total row:", tr)
