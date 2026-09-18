"""Step 5 — HSN summary (GSTR-9 Table 17 shape) + Rate-wise summary, formula-driven.
Sheets: Summary | HSN Summary (GSTIN x HSN x UQC x Rate, live SUMIFS: qty, taxable, I/C/S) |
Rate-wise Summary (state x rate) | Checks | SR Items (item-level source).
Advances excluded (no HSN, not supplies with an HSN line). Services (SAC 99xx): qty/UQC blank; blank UQC written as '-' (an empty SUMIFS criterion cell would mean 'equals 0')."""
import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
SP = os.path.dirname(os.path.abspath(__file__))
REG = r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
def S(v): return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); GF = PatternFill("solid", fgColor="DDEBF7")
TOT = Font(bold=True)

# ---- read item rows straight from the register (values incl. cached formulas)
wbr = openpyxl.load_workbook(REG, read_only=True, data_only=True)
ws = wbr["SR_2025-26"]
hdr = None; rows = []
for i, r in enumerate(ws.iter_rows(min_row=4, max_row=27006, values_only=True), start=4):
    if i == 4: hdr = [S(v) for v in r]; continue
    rows.append(r)
wbr.close()
ix = {n: k for k, n in enumerate(hdr)}
def g(r, n): return r[ix[n]]
items = []
for r in rows:
    dtc = S(g(r, "Document Type Code"))
    if dtc.startswith("MOB"): continue
    items.append({
        "state": S(g(r, "My State")), "gstin": S(g(r, "My GSTIN")), "dtc": dtc,
        "hsn": S(g(r, "HSN or SAC Code")), "gs": ("S" if S(g(r, "HSN or SAC Code")).startswith("99") else ("G" if S(g(r, "HSN or SAC Code")) else "")),
        "uqc": S(g(r, "Unit of Measurement for uploading")) or "-", "rate": g(r, "GST Rate") or 0,
        "qty": g(r, "Quantity for uploading") or 0, "tax": g(r, "Taxable Value") or 0,
        "igst": g(r, "IGST Amount") or 0, "cgst": g(r, "CGST Amount") or 0, "sgst": g(r, "SGST Amount") or 0})
D = pd.DataFrame(items)
print("item rows (ex-advances):", len(D), "| blank HSN:", int((D['hsn'] == '').sum()))

wb = Workbook()
# ---- SR Items (source of the SUMIFS)
si = wb.active; si.title = "SR Items"
si.append(["My GSTIN", "State", "Doc Type", "HSN/SAC", "G/S", "UQC", "Rate", "Quantity", "Taxable", "IGST", "CGST", "SGST"])
for c in si[1]: c.font = HF; c.fill = HB
for _, r in D.iterrows():
    si.append([r["gstin"], r["state"], r["dtc"], r["hsn"], r["gs"], r["uqc"], r["rate"],
               round(float(r["qty"] or 0), 3), round(float(r["tax"]), 2), round(float(r["igst"]), 2),
               round(float(r["cgst"]), 2), round(float(r["sgst"]), 2)])
NS = si.max_row
for c_, w in zip("ABCDEFGHIJKL", [18, 18, 9, 12, 6, 8, 7, 12, 15, 12, 13, 13]): si.column_dimensions[c_].width = w
si.freeze_panes = "A2"
SI = lambda c: "'SR Items'!$%s$2:$%s$%d" % (c, c, NS)

# ---- HSN Summary: one row per GSTIN x HSN x UQC x Rate (combos from data; cells live)
combos = D.groupby(["gstin", "state", "hsn", "gs", "uqc", "rate"], dropna=False).size().reset_index().rename(columns={0: "n"})
combos = combos.sort_values(["state", "hsn", "rate"])
hs = wb.create_sheet("HSN Summary", 0)
hs["A1"] = "HSN-wise summary of outward supplies — GSTR-9 Table 17 shape (per GSTIN). Advances excluded. Live SUMIFS over 'SR Items'."
hs["A1"].font = Font(bold=True)
hs.append([]); hs.append(["State", "My GSTIN", "HSN/SAC", "Goods/Service", "UQC", "Rate %", "Total Quantity", "Taxable Value", "IGST", "CGST", "SGST", "Total Tax"])
for c in hs[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
r = 3
for _, cb in combos.iterrows():
    r += 1
    hs.cell(r, 1, cb["state"]); hs.cell(r, 2, cb["gstin"]); hs.cell(r, 3, cb["hsn"] or "(blank)")
    hs.cell(r, 4, cb["gs"]); hs.cell(r, 5, cb["uqc"]); hs.cell(r, 6, cb["rate"])
    crit = '%s,$B%d,%s,$C%d,%s,$E%d,%s,$F%d' % (SI("A"), r, SI("D"), r, SI("F"), r, SI("G"), r)
    for j, col in enumerate(["H", "I", "J", "K", "L"], start=7):
        hs.cell(r, j, "=SUMIFS(%s,%s)" % (SI(col), crit))
        hs.cell(r, j).number_format = "#,##0.00"
    hs.cell(r, 12, "=I%d*0+SUM(I%d:K%d)-I%d" % (r, r, r, r))
    hs.cell(r, 12).value = "=SUM(I%d:K%d)" % (r, r)
    hs.cell(r, 12).number_format = "#,##0.00"
NH = r
r += 1
hs.cell(r, 1, "Total").font = TOT
for j in range(7, 13):
    hs.cell(r, j, "=SUM(%s4:%s%d)" % (L(j), L(j), NH)); hs.cell(r, j).font = TOT; hs.cell(r, j).number_format = "#,##0.00"
hs.auto_filter.ref = "A3:L%d" % NH
for c_, w in zip("ABCDEFGHIJKL", [18, 18, 12, 12, 8, 8, 14, 16, 13, 13, 13, 14]): hs.column_dimensions[c_].width = w
hs.freeze_panes = "A4"

# ---- Rate-wise Summary: state x rate
rw = wb.create_sheet("Rate-wise Summary", 1)
rw["A1"] = "Rate-wise summary of outward supplies (advances excluded). Live SUMIFS over 'SR Items'."
rw["A1"].font = Font(bold=True)
rw.append([]); rw.append(["State", "My GSTIN", "Rate %", "Taxable Value", "IGST", "CGST", "SGST", "Total Tax", "Actual rate check %"])
for c in rw[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
rc = D.groupby(["gstin", "state", "rate"], dropna=False).size().reset_index().sort_values(["state", "rate"])
r = 3
for _, cb in rc.iterrows():
    r += 1
    rw.cell(r, 1, cb["state"]); rw.cell(r, 2, cb["gstin"]); rw.cell(r, 3, cb["rate"])
    crit = '%s,$B%d,%s,$C%d' % (SI("A"), r, SI("G"), r)
    for j, col in enumerate(["I", "J", "K", "L"], start=4):
        rw.cell(r, j, "=SUMIFS(%s,%s)" % (SI(col), crit)); rw.cell(r, j).number_format = "#,##0.00"
    rw.cell(r, 8, "=SUM(E%d:G%d)" % (r, r)); rw.cell(r, 8).number_format = "#,##0.00"
    rw.cell(r, 9, '=IF(N(D%d)=0,"",ROUND(H%d/D%d*100,2))' % (r, r, r)); rw.cell(r, 9).number_format = "0.00"
NR_ = r
r += 1
rw.cell(r, 1, "Total").font = TOT
for j in range(4, 9):
    rw.cell(r, j, "=SUM(%s4:%s%d)" % (L(j), L(j), NR_)); rw.cell(r, j).font = TOT; rw.cell(r, j).number_format = "#,##0.00"
rw.auto_filter.ref = "A3:I%d" % NR_
for c_, w in zip("ABCDEFGHI", [18, 18, 8, 16, 13, 13, 13, 14, 14]): rw.column_dimensions[c_].width = w
rw.freeze_panes = "A4"

# ---- Checks
ck = wb.create_sheet("Checks", 2)
ck["A1"] = "HSN data-quality checks (live)"; ck["A1"].font = Font(bold=True)
ck.append([]); rows_ = [
 ["Item rows covered (ex-advances)", "=COUNTA(%s)" % SI("A")],
 ["Taxable total here", "=SUM(%s)" % SI("I"), "must equal register taxable ex-advances 84,86,10,5710.82".replace(",", "")],
 ["Rows with blank HSN", '=COUNTIF(%s,"")' % SI("D"), "advances excluded, so should be 0"],
 ["Service rows (SAC 99xx)", '=COUNTIF(%s,"S")' % SI("E"), "quantity/UQC legitimately blank"],
 ["Goods rows with zero quantity", '=COUNTIFS(%s,"G",%s,0)' % (SI("E"), SI("H")), "flag - Table 17 needs quantity for goods"],
 ["HSN codes not 6 or 8 digits", '=SUMPRODUCT(--(LEN(%s)<>6),--(LEN(%s)<>8),--(%s<>""))' % (SI("D"), SI("D"), SI("D")), "6-digit minimum applies (turnover > 5 crore)"],
]
for x in rows_: ck.append(x)
ck.column_dimensions["A"].width = 44; ck.column_dimensions["B"].width = 20; ck.column_dimensions["C"].width = 60

# ---- Summary
su = wb.create_sheet("Summary", 0)
su.append(["STEP 5 — HSN & Rate-wise summaries (FY 2025-26) — all figures live"]); su["A1"].font = Font(bold=True, size=12)
su.append([])
su.append(["Distinct GSTIN x HSN x UQC x Rate combinations", NH - 3])
su.append(["Distinct GSTIN x Rate combinations", NR_ - 3])
su.append(["Distinct HSN/SAC codes", int(D[D['hsn'] != '']['hsn'].nunique())])
su.append(["Goods vs services (item rows)", "G: %d | S: %d" % (int((D['gs'] == 'G').sum()), int((D['gs'] == 'S').sum()))])
su.append([])
su.append(["Note", "Advances carry no HSN and are excluded (they are not HSN-line supplies). Table 17 is filled per GSTIN - filter column B."])
su.column_dimensions["A"].width = 48; su.column_dimensions["B"].width = 60
OUT = r"C:\Users\pawar\Downloads\VEL_Step5_HSN_Rate_Summary_DRAFT.xlsx"
wb.save(OUT)
print("WROTE", OUT)
print("HSN combos:", NH - 3, "| rate combos:", NR_ - 3, "| SR Items rows:", NS - 1)
