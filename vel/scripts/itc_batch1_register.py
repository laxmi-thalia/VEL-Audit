"""Batch-1 step B: rebuild 'ITC Register 2025-26' per the CA rulings of 17-Sep.
- drop G Category as per 3B, AD Invoice Level Match (text -> Reco Remarks), AU Review Remarks, BI-BL, BS POS Query Description, BY Query
- add PO Number (ZFI06) after Expense GL Element
- KEY = normalised vendor GSTIN & invoice (formula); Countif = Consider/NA (stamped: first line per matched 2B key)
- B_ block = SUMIFS over the register by KEY on Consider lines (NA otherwise); 2B_ block = SUMIFS over 'GSTR-2B Apr25-Aug26' by KEY2
- KEY2 = the 2B key the cascade matched (stamped); as-per-2B cols = INDEX/MATCH into the 2B sheet
- POS Check TRUE/blank (as is) + POS Query text; Expense GL Element/PO from hidden 'ZFI06 Data'; Expense Description from 'TB Groupings'
- cascade re-run: register vs new Octa 2B; unmatched vs FY 24-25 rows of old 'GSTR-2B ITC Data' (for 6A1 component 1)
- dependents ('ITCR vs 3B Net ITC', 'ITCR vs 2B GSTN Level', 'RCM Paid vs ITC Claimed') re-pointed by column map."""
import os, re, warnings, datetime as dt, collections, pickle
from copy import copy
warnings.filterwarnings("ignore")
import openpyxl, pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v)
    except Exception: return 0.0
def norm(s): return re.sub(r"[ \-/.'_]", "", S(s).upper())
SP = os.path.dirname(os.path.abspath(__file__))
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
Z = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/ZFI Reports/ZFI06_01.04.2025 to 31.3.2026.xlsx"
open(P, "r+b").close()
NF = "#,##0.00"
wb = openpyxl.load_workbook(P)
reg = wb["ITC Register 2025-26"]
H = [S(reg.cell(5, c).value) for c in range(1, reg.max_column + 1)]
while H and H[-1] == "": H.pop()
NC0 = len(H); R0 = 6
last = max(r for r in range(R0, reg.max_row + 1) if reg.cell(r, 4).value)
print("register: %d cols, rows %d..%d" % (NC0, R0, last))
rows = [[reg.cell(r, c).value for c in range(1, NC0 + 1)] for r in range(R0, last + 1)]
row4 = [reg.cell(4, c).value for c in range(1, NC0 + 1)]
styles = {c: (copy(reg.cell(5, c).font), copy(reg.cell(5, c).fill), copy(reg.cell(5, c).alignment), reg.cell(R0, c).number_format, reg.column_dimensions[L(c)].width) for c in range(1, NC0 + 1)}
old_i = {h: i for i, h in enumerate(H)}
# ---------------- 2B (new Octa sheet) ----------------
b2 = wb["GSTR-2B Apr25-Aug26"]
BH = [S(b2.cell(2, c).value) for c in range(1, b2.max_column + 1)]
bi = {h: i for i, h in enumerate(BH)}
b2rows = [[b2.cell(r, c).value for c in range(1, len(BH) + 1)] for r in range(3, b2.max_row + 1) if b2.cell(r, 1).value]
NB = 2 + len(b2rows)
print("new 2B rows:", len(b2rows))
# 2B KEY column -> live formula with the same normalisation as the register
kcol = bi["KEY"] + 1; gcol = L(bi["Supplier GSTIN"] + 1); dcol = L(bi["Doc No"] + 1)
def normf(expr): return 'UPPER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(%s," ",""),"-",""),"/",""),".",""),"\'",""),"_",""))' % expr
for r in range(3, NB + 1): b2.cell(r, kcol).value = "=" + normf("%s%d&%s%d" % (gcol, r, dcol, r))
by_inv = collections.defaultdict(list); by_dt = collections.defaultdict(list); by_amt = collections.defaultdict(list)
b2key = []
for i, r in enumerate(b2rows):
    vg = S(r[bi["Supplier GSTIN"]]).upper(); k = vg + norm(r[bi["Doc No"]]); b2key.append(k)
    tot = round(num(r[bi["IGST (Net)"]]) + num(r[bi["CGST (Net)"]]) + num(r[bi["SGST (Net)"]]), 2)
    d = r[bi["Doc Date"]]
    if vg:
        by_inv[k].append(i)
        if isinstance(d, dt.datetime): by_dt[(vg, d.date(), tot)].append(i)
        by_amt[(vg, tot)].append(i)
# old 2B FY 24-25 portion (for 6A1 component 1)
ob = wb["GSTR-2B ITC Data"]
OH = [S(ob.cell(5, c).value) for c in range(1, ob.max_column + 1)]; oi = {h: i for i, h in enumerate(OH)}
o_inv = collections.defaultdict(list); o_dt = collections.defaultdict(list); orows = []
for r in ob.iter_rows(min_row=6, values_only=True):
    if not r or not r[0] or S(r[oi["FY (derived)"]]) != "2024-25": continue
    vg = S(r[oi["GSTIN of supplier"]]).upper(); k = vg + norm(r[oi["Invoice number"]]); orows.append((k, r))
    tot = round(num(r[oi["Integrated Tax(₹)"]]) + num(r[oi["Central Tax(₹)"]]) + num(r[oi["State/UT Tax(₹)"]]), 2)
    d = r[oi["Invoice Date"]]
    if vg:
        o_inv[k].append(len(orows) - 1)
        if isinstance(d, dt.datetime): o_dt[(vg, d.date(), tot)].append(len(orows) - 1)
print("old 2B FY 24-25 rows:", len(orows))
# ---------------- cascade ----------------
iVG, iINV, iDATE = old_i["Vendor GSTIN"], old_i["Invoice No."], old_i["Invoice Date"]
iI, iC, iS_ = old_i["IGST"], old_i["CGST"], old_i["SGST"]
used = {}; verdict = []; key2 = []
def take(c, ri):
    for b in c:
        if b not in used: used[b] = ri; return b
    return c[0]
for ri, r in enumerate(rows):
    vg = S(r[iVG]).upper(); inv = norm(r[iINV]); k = vg + inv
    tot = round(num(r[iI]) + num(r[iC]) + num(r[iS_]), 2); d = r[iDATE]
    if not vg or not re.match(r"^\d{2}[A-Z0-9]{13}$", vg):
        verdict.append("No vendor GSTIN (URD/ISD/RCM-self)"); key2.append(""); continue
    c1 = by_inv.get(k) if inv else None
    if c1: b = take(c1, ri); verdict.append("Matched (invoice no)"); key2.append(b2key[b]); continue
    c2 = by_dt.get((vg, d.date(), tot)) if isinstance(d, dt.datetime) else None
    if c2: b = take(c2, ri); verdict.append("Matched (date+amount)"); key2.append(b2key[b]); continue
    c3 = by_amt.get((vg, tot))
    if c3 and len(c3) <= 3 and tot: b = take(c3, ri); verdict.append("Matched (amount, <=3 candidates)"); key2.append(b2key[b]); continue
    # FY 24-25 old 2B
    if o_inv.get(k) or (isinstance(d, dt.datetime) and o_dt.get((vg, d.date(), tot))):
        verdict.append("Matched in FY 24-25 2B (working files) - 6A1 component 1"); key2.append("PY:" + k); continue
    verdict.append("NOT FOUND IN 2B (Apr25-Aug26)"); key2.append("")
print("verdicts:", dict(collections.Counter(verdict)))
seen = set(); consider = []
for k in key2:
    if k and not k.startswith("PY:") and k not in seen: seen.add(k); consider.append("Consider")
    else: consider.append("NA")
print("Consider lines:", consider.count("Consider"))
# ---------------- ZFI06 hidden sheet ----------------
z = pd.read_excel(Z, dtype=str); z.columns = [S(c) for c in z.columns]
z = z[[c for c in z.columns if c and not c.startswith("Unnamed")]]
if "ZFI06 Data" in wb.sheetnames: del wb["ZFI06 Data"]
zs = wb.create_sheet("ZFI06 Data"); zs.row_dimensions[1].height = 21
for c, h in enumerate(z.columns, 1): x = zs.cell(2, c, h); x.font = Font(bold=True)
for i, row in enumerate(z.itertuples(index=False), 3):
    for c, v in enumerate(row, 1):
        if isinstance(v, str) and v not in ("nan", ""): zs.cell(i, c, v)
NZ = 2 + len(z); zs.sheet_state = "hidden"
zc = {h: L(i + 1) for i, h in enumerate(z.columns)}
print("ZFI06 Data embedded:", len(z), "rows")
# TB Groupings range
tb = wb["TB Groupings"]; tb_first = next(r for r in range(1, 10) if isinstance(tb.cell(r, 1).value, (int, float)) or S(tb.cell(r, 1).value).isdigit()); tb_last = tb.max_row
# ---------------- new column layout ----------------
DROP = {"Category as per 3B", "Invoice Level Match", "Review Remarks", "RCM Paid Month", "Remarks", "Remarks 2", "Comments", "POS Query Description", "Query"}
NEWH = []
for h in H:
    if h in DROP: continue
    NEWH.append(h)
    if h == "Expense GL Element": NEWH.append("PO Number")
NEWH[NEWH.index("KEY2")] = "KEY2 (matched 2B key)"
ni = {h: i + 1 for i, h in enumerate(NEWH)}  # 1-based new col
NC = len(NEWH)
print("new cols:", NC)
del wb["ITC Register 2025-26"]
ws = wb.create_sheet("ITC Register 2025-26", wb.sheetnames.index("GSTR-2B Apr25-Aug26"))
ws.row_dimensions[1].height = 21
ws.cell(2, 1, "VIKRAN ENGINEERING LIMITED").font = Font(bold=True)
ws.cell(3, 1, "Input Tax Credit (ITC) Register for FY 2025-26 - poured from 'ITC All state FY 2025-26.xlsx' (Working, Company=VEL). 2B side = 'GSTR-2B Apr25-Aug26' (Octa). B_/2B_ blocks = SUMIFS by KEY on the Consider line of each document (CA ruling 17-09); Countif/KEY2 stamped from the invoice-level cascade - regenerate on data change.").font = Font(bold=True)
oldcol_of = {}
for h in NEWH:
    if h in old_i: oldcol_of[h] = old_i[h] + 1
    elif h == "KEY2 (matched 2B key)": oldcol_of[h] = old_i["KEY2"] + 1
for h, c in ni.items():
    x = ws.cell(5, c, h)
    src = oldcol_of.get(h, old_i["Expense GL Element"] + 1)
    st = styles[src]; x.font = copy(st[0]); x.fill = copy(st[1]); x.alignment = copy(st[2]); ws.column_dimensions[L(c)].width = st[4] or 14
c = lambda h: L(ni[h])
RG = lambda h: "$%s$%d:$%s$%d" % (c(h), R0, c(h), R0 + len(rows) - 1)
B2 = lambda h: "'GSTR-2B Apr25-Aug26'!$%s$3:$%s$%d" % (L(bi[h] + 1), L(bi[h] + 1), NB)
ZF = lambda h: "'ZFI06 Data'!$%s$3:$%s$%d" % (zc[h], zc[h], NZ)
FORM = {
    "Total GST": "=SUM(%s{r}:%s{r})" % (c("IGST"), c("SGST")),
    "KEY": "=" + normf("$%s{r}&$%s{r}" % (c("Vendor GSTIN"), c("Invoice No."))),
    "B_IGST": '=IF($%s{r}="Consider",SUMIFS(%s,%s,$%s{r}),"NA")' % (c("Countif"), RG("IGST"), RG("KEY2 (matched 2B key)"), c("KEY2 (matched 2B key)")),
    "B_CGST": '=IF($%s{r}="Consider",SUMIFS(%s,%s,$%s{r}),"NA")' % (c("Countif"), RG("CGST"), RG("KEY2 (matched 2B key)"), c("KEY2 (matched 2B key)")),
    "B_SGST": '=IF($%s{r}="Consider",SUMIFS(%s,%s,$%s{r}),"NA")' % (c("Countif"), RG("SGST"), RG("KEY2 (matched 2B key)"), c("KEY2 (matched 2B key)")),
    "B_Total GST": '=IF($%s{r}="Consider",%s{r}+%s{r}+%s{r},"NA")' % (c("Countif"), c("B_IGST"), c("B_CGST"), c("B_SGST")),
    "2B_IGST": '=IF($%s{r}="Consider",SUMIFS(%s,%s,$%s{r}),"NA")' % (c("Countif"), B2("IGST (Net)"), B2("KEY"), c("KEY2 (matched 2B key)")),
    "2B_CGST": '=IF($%s{r}="Consider",SUMIFS(%s,%s,$%s{r}),"NA")' % (c("Countif"), B2("CGST (Net)"), B2("KEY"), c("KEY2 (matched 2B key)")),
    "2B_SGST": '=IF($%s{r}="Consider",SUMIFS(%s,%s,$%s{r}),"NA")' % (c("Countif"), B2("SGST (Net)"), B2("KEY"), c("KEY2 (matched 2B key)")),
    "2B_Total GST": '=IF($%s{r}="Consider",%s{r}+%s{r}+%s{r},"NA")' % (c("Countif"), c("2B_IGST"), c("2B_CGST"), c("2B_SGST")),
    "D_IGST": '=IF($%s{r}="Consider",%s{r}-%s{r},"NA")' % (c("Countif"), c("B_IGST"), c("2B_IGST")),
    "D_CGST": '=IF($%s{r}="Consider",%s{r}-%s{r},"NA")' % (c("Countif"), c("B_CGST"), c("2B_CGST")),
    "D_SGST": '=IF($%s{r}="Consider",%s{r}-%s{r},"NA")' % (c("Countif"), c("B_SGST"), c("2B_SGST")),
    "D_Total GST": '=IF($%s{r}="Consider",%s{r}-%s{r},"NA")' % (c("Countif"), c("B_Total GST"), c("2B_Total GST")),
    "Invoice as per 2B": '=IF($%s{r}="","",IFERROR(INDEX(%s,MATCH($%s{r},%s,0)),""))' % (c("KEY2 (matched 2B key)"), B2("Doc No"), c("KEY2 (matched 2B key)"), B2("KEY")),
    "2B Period": '=IF($%s{r}="","",IFERROR(INDEX(%s,MATCH($%s{r},%s,0)),""))' % (c("KEY2 (matched 2B key)"), B2("Tax Period"), c("KEY2 (matched 2B key)"), B2("KEY")),
    "My GSTN as per 2B": '=IF($%s{r}="","",IFERROR(INDEX(%s,MATCH($%s{r},%s,0)),""))' % (c("KEY2 (matched 2B key)"), B2("Company GSTIN"), c("KEY2 (matched 2B key)"), B2("KEY")),
    "Supplier GSTN as per 2B": '=IF($%s{r}="","",IFERROR(INDEX(%s,MATCH($%s{r},%s,0)),""))' % (c("KEY2 (matched 2B key)"), B2("Supplier GSTIN"), c("KEY2 (matched 2B key)"), B2("KEY")),
    "Vendor": "=LEFT($%s{r},2)" % c("Vendor GSTIN"),
    "My GSTN": "=LEFT($%s{r},2)" % c("VEL GSTIN"),
    "As per State": '=IF($%s{r}=$%s{r},"Intra State","Inter State")' % (c("Vendor"), c("My GSTN")),
    "As per Amounts": '=IF(N($%s{r})=0,"Intra State","Inter State")' % c("IGST"),
    "POS Check": '=IF($%s{r}=$%s{r},TRUE,"")' % (c("As per Amounts"), c("As per State")),
    "POS Query": '=IF($%s{r}=TRUE,"","POS mismatch - vendor state code "&$%s{r}&" vs tax head "&$%s{r})' % (c("POS Check"), c("Vendor"), c("As per Amounts")),
    "Expense GL Element": '=IFERROR(INDEX(%s,MATCH(TEXT($%s{r},"0"),%s,0)),"")' % (ZF("Expense-GL Element"), c("Document Number"), ZF("Document Number")),
    "PO Number": '=IFERROR(INDEX(%s,MATCH(TEXT($%s{r},"0"),%s,0)),"")' % (ZF("PO Number"), c("Document Number"), ZF("Document Number")),
    "Expense Description": '=IF($%s{r}="","",IFERROR(INDEX(\'TB Groupings\'!$B$%d:$B$%d,MATCH(VALUE($%s{r}),\'TB Groupings\'!$A$%d:$A$%d,0)),""))' % (c("Expense GL Element"), tb_first, tb_last, c("Expense GL Element"), tb_first, tb_last),
}
STAMP = {"Countif": consider, "KEY2 (matched 2B key)": key2, "Reco Remarks": verdict}
NUMC = {"Taxable Value", "IGST", "CGST", "SGST", "Total GST", "B_IGST", "B_CGST", "B_SGST", "B_Total GST", "2B_IGST", "2B_CGST", "2B_SGST", "2B_Total GST", "D_IGST", "D_CGST", "D_SGST", "D_Total GST"}
for i, r in enumerate(rows):
    rr = R0 + i
    for h, cc in ni.items():
        if h in FORM: ws.cell(rr, cc).value = FORM[h].replace("{r}", str(rr))
        elif h in STAMP: ws.cell(rr, cc).value = STAMP[h][i]
        elif h in oldcol_of: ws.cell(rr, cc).value = r[oldcol_of[h] - 1]
        if h in NUMC: ws.cell(rr, cc).number_format = NF
        elif h in oldcol_of and styles[oldcol_of[h]][3] != "General": ws.cell(rr, cc).number_format = styles[oldcol_of[h]][3]
RN = R0 + len(rows) - 1
for h, cc in ni.items():
    if h in ("Taxable Value", "IGST", "CGST", "SGST", "Total GST"):
        ws.cell(4, cc, "=SUBTOTAL(9,%s%d:%s%d)" % (L(cc), R0, L(cc), RN)).number_format = NF; ws.cell(4, cc).font = Font(bold=True)
ws.freeze_panes = "A6"; ws.auto_filter.ref = "A5:%s%d" % (L(NC), RN)
# amber for CA-typed
AMB = PatternFill("solid", fgColor="FFF2CC")
for h in ("Reco Remarks", "Eligibility", "Consider 8A reco", "Considered in Table 6A1", "Remarks for accounting entries- For 6A1"):
    if h in ni: ws.cell(5, ni[h]).fill = AMB; ws.cell(5, ni[h]).font = Font(bold=True)
# ---------------- re-point dependents ----------------
MAP = {}
for h, newc in ni.items():
    if h in oldcol_of: MAP[L(oldcol_of[h])] = L(newc)
pat = re.compile(r"('ITC Register 2025-26'!\$?)([A-Z]{1,3})(\$?\d+)(?::(\$?)([A-Z]{1,3})(\$?\d+))?")
def sub(m):
    a, c1, r1, d, c2, r2 = m.groups()
    out = a + MAP.get(c1, c1) + r1
    if c2: out += ":" + d + MAP.get(c2, c2) + r2
    return out
n = 0
for nm in wb.sheetnames:
    if nm == "ITC Register 2025-26": continue
    for rowc in wb[nm].iter_rows():
        for cell in rowc:
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "ITC Register 2025-26" in v:
                nv = pat.sub(sub, v)
                if nv != v: cell.value = nv; n += 1
print("dependent formulas re-pointed:", n)
wb.save(P)
pickle.dump({"NEWH": NEWH, "verdict": verdict, "key2": key2, "consider": consider, "RN": RN}, open(os.path.join(SP, "itc_b1_meta.pkl"), "wb"))
print("saved; register rows %d..%d, %d cols" % (R0, RN, NC))
