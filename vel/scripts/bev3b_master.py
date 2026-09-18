"""'3B Birds Eye View' sheet in the master, replicating Priyesh's Birds_Eye_View.xlsx layout:
3-row header (group band r5 / exact Section text r6 / exact Type text r7 = SUMIFS criteria), all
GSTR-3B tables in form order incl. the nil ones (amber), Cess per table, tie-out check cols,
then OUR cols last (SR Taxable / Diff / Vice-versa). Per state: 12 month rows + TOTAL; ALL STATES
block; notes. Raw Octa matrices land as hidden 'GSTR-3B XX' sheets (row 1 reserved, header r3,
data r4..) so every cell is a live SUMIFS. '3B Data' is NOT touched."""
import os, re, warnings
from copy import copy
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
def S(v): return "" if v is None else str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
T = r"C:\Users\pawar\Downloads\Birds_Eye_View.xlsx"
B = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-3B/"
f = open(P, "r+b"); f.close()
tpl = openpyxl.load_workbook(T)["Birds Eye View"]
# ---- template column spec (rows 5-7) up to the tie-out block
NCOLS_T = 70   # D..BR data cols in template = 4..70
spec = []      # (group, section, type, amber)
for c in range(4, NCOLS_T + 1):
    spec.append((tpl.cell(5, c).value, S(tpl.cell(6, c).value), S(tpl.cell(7, c).value), tpl.cell(6, c).fill.fgColor.rgb == "FFFFF2CC"))
groups = {}    # first col of each group -> (label, last col)
cur = None
for i, (g, sec, typ, amb) in enumerate(spec):
    if g: cur = i; groups[cur] = [g, i]
    else: groups[cur][1] = i
print("template data columns:", len(spec), "| groups:", [(v[0][:28], L(4 + k), L(4 + v[1])) for k, v in groups.items()])
# ---- styles from template
def sty(src, dst):
    dst.font = copy(src.font); dst.fill = copy(src.fill); dst.border = copy(src.border); dst.alignment = copy(src.alignment); dst.number_format = src.number_format
CODE = {"Jammu and Kashmir": "JK", "Punjab": "PB", "Rajasthan": "RJ", "Uttar Pradesh": "UP", "Bihar": "BH", "Arunachal Pradesh": "AR", "Assam": "AS",
        "West Bengal": "WB", "Jharkhand": "JH", "Chhattisgarh": "CG", "Madhya Pradesh": "MP", "Gujarat": "GJ", "Maharashtra": "MH", "Karnataka": "KA",
        "Kerala": "KL", "Tamil Nadu": "TN", "Telangana": "TG", "Andhra Pradesh": "AP"}
NAME = {"Jammu and Kashmir": "Jammu & Kashmir"}
wb = openpyxl.load_workbook(P)
# ---- raw state sheets
states = []   # (code, name, gstin, sheetname, last_row)
for x in sorted(os.listdir(B)):
    if not x.lower().endswith(".xlsx") or x.startswith("~$"): continue
    st = re.sub(r"-2025-26.*$", "", x.split("LIMITED-")[1]); code = CODE.get(st, st[:2].upper())
    w = openpyxl.load_workbook(B + x, read_only=True, data_only=True)
    ov = w["Overview"]; gstin = None
    for row in ov.iter_rows(values_only=True):
        if len(row) > 2 and S(row[1]).lower() == "gstin": gstin = S(row[2]).split("(")[0].strip()
    m = w["GSTR-3B"]; rows = [r for r in m.iter_rows(values_only=True)]
    nm = "GSTR-3B " + code
    if nm in wb.sheetnames: del wb[nm]
    ws = wb.create_sheet(nm)
    ws.row_dimensions[1].height = 21
    ws.cell(2, 2, "%s (%s) - raw Octa GSTR-3B matrix, source of '3B Birds Eye View'" % (gstin, st)).font = Font(bold=True)
    for i, r in enumerate(rows):
        for c, v in enumerate(r[:16], 1):
            cell = ws.cell(3 + i, c, v)
            if i == 0: cell.font = Font(bold=True)
            elif isinstance(v, (int, float)): cell.number_format = "#,##0.00"
    last = 3 + len(rows) - 1
    ws.column_dimensions["B"].width = 60; ws.column_dimensions["C"].width = 18
    ws.sheet_state = "hidden"
    states.append((code, NAME.get(st, st), gstin, nm, last))
    w.close()
states.sort(key=lambda s: s[2])   # GST state-code order
print("raw state sheets:", len(states))
# ---- Birds Eye View sheet
NM = "3B Birds Eye View"
if NM in wb.sheetnames: del wb[NM]
ws = wb.create_sheet(NM, wb.sheetnames.index("3B Data") + 1)
ws.row_dimensions[1].height = 21
D0 = 4; ND = len(spec); TIE0 = D0 + ND; OUR0 = TIE0 + 3; NC = OUR0 + 2
ws.cell(2, 1, "GSTR-3B - STATE-WISE / MONTH-WISE CONSOLIDATED SUMMARY (Birds Eye View)"); sty(tpl["A1"], ws.cell(2, 1))
ws.cell(3, 1, "FY 2025-26 (Apr 2025 to Mar 2026)  |  PAN AAECR0503Q  |  Every figure is a live SUMIFS over the hidden 'GSTR-3B XX' raw Octa sheets in this workbook - no value is typed in. Last three columns = books tie (SR_2025-26)."); sty(tpl["A2"], ws.cell(3, 1))
ws.cell(4, 1, "Rows 6 and 7 hold the exact Section and Type text used as the SUMIFS criteria. Change nothing in rows 6-7 or the lookups break."); sty(tpl["A3"], ws.cell(4, 1))
# header rows 5/6/7 (template rows 5/6/7)
for i, (g, sec, typ, amb) in enumerate(spec):
    c = D0 + i
    ws.cell(5, c, g); sty(tpl.cell(5, 4 + i), ws.cell(5, c))
    ws.cell(6, c, sec); sty(tpl.cell(6, 4 + i), ws.cell(6, c))
    ws.cell(7, c, typ); sty(tpl.cell(7, 4 + i), ws.cell(7, c))
    ws.column_dimensions[L(c)].width = tpl.column_dimensions[L(4 + i)].width or 15
for k, (lbl, last) in groups.items():
    ws.merge_cells(start_row=5, start_column=D0 + k, end_row=5, end_column=D0 + last)
for j, h in enumerate(("TIE-OUT CHECK", "TIE-OUT CHECK", "TIE-OUT CHECK")):
    c = TIE0 + j; ws.cell(5, c, h); sty(tpl.cell(5, 71 + j), ws.cell(5, c)); ws.merge_cells(start_row=5, start_column=c, end_row=6, end_column=c)
    ws.cell(7, c, tpl.cell(7, 71 + j).value); sty(tpl.cell(7, 71 + j), ws.cell(7, c)); ws.column_dimensions[L(c)].width = 17
for j, h in enumerate(("SR Taxable (books, live)", "Diff (3B 3.1.A taxable - books)", "Vice-versa status")):
    c = OUR0 + j; ws.cell(5, c, "BOOKS TIE (added by DPS)"); sty(tpl.cell(5, 71), ws.cell(5, c)); ws.merge_cells(start_row=5, start_column=c, end_row=6, end_column=c)
    ws.cell(7, c, h); sty(tpl.cell(7, 71), ws.cell(7, c)); ws.column_dimensions[L(c)].width = 20 if j < 2 else 40
for c in range(1, 4):
    for r in (5, 6, 7): sty(tpl.cell(r, c), ws.cell(r, c))
ws.cell(7, 1, "State"); ws.cell(7, 2, "GSTIN"); ws.cell(7, 3, "Month")
for c, w in zip("ABC", (18, 19, 16)): ws.column_dimensions[c].width = w
for r, h in ((5, 30), (6, 61.5), (7, 30)): ws.row_dimensions[r].height = h
# data rows
MONTHS = ["Apr 2025", "May 2025", "Jun 2025", "Jul 2025", "Aug 2025", "Sep 2025", "Oct 2025", "Nov 2025", "Dec 2025", "Jan 2026", "Feb 2026", "Mar 2026"]
MSHORT = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26"]
tpl_data = tpl.cell(8, 4); tpl_tot = tpl.cell(19, 4); tpl_lab = tpl.cell(8, 1); tpl_totlab = tpl.cell(19, 1)
i31a = next(i for i, s_ in enumerate(spec) if s_[1].startswith("3.1.A") and s_[2] == "Supply Value")
SR = lambda col: "'SR_2025-26'!$%s$6:$%s$27007" % (col, col)
r = 7
month_rows = {m: [] for m in MONTHS}
for code, name, gstin, snm, last in states:
    first = r + 1
    for mi, m in enumerate(MONTHS):
        r += 1
        ws.cell(r, 1, name); ws.cell(r, 2, gstin); ws.cell(r, 3, m)
        for c in range(1, 4): sty(tpl_lab, ws.cell(r, c))
        mc = L(4 + mi)
        rng = "'%s'!$%s$4:$%s$%d" % (snm, mc, mc, last); secr = "'%s'!$B$4:$B$%d" % (snm, last); typr = "'%s'!$C$4:$C$%d" % (snm, last)
        for i in range(ND):
            c = D0 + i
            ws.cell(r, c, "=SUMIFS(%s,%s,%s$6,%s,%s$7)" % (rng, secr, L(c), typr, L(c))); sty(tpl_data, ws.cell(r, c))
        ws.cell(r, TIE0, "=SUM(%s%d:%s%d)" % (L(D0), r, L(D0 + ND - 1), r)); sty(tpl_data, ws.cell(r, TIE0))
        ws.cell(r, TIE0 + 1, "=SUM(%s)" % rng); sty(tpl_data, ws.cell(r, TIE0 + 1))
        ws.cell(r, TIE0 + 2, "=ROUND(%s%d-%s%d,2)" % (L(TIE0), r, L(TIE0 + 1), r)); sty(tpl.cell(8, 73), ws.cell(r, TIE0 + 2))
        ws.cell(r, OUR0, "=SUMIFS(%s,%s,$B%d,%s,\"%s\")" % (SR("T"), SR("C"), r, SR("A"), MSHORT[mi])); sty(tpl_data, ws.cell(r, OUR0))
        ws.cell(r, OUR0 + 1, "=%s%d-%s%d" % (L(D0 + i31a), r, L(OUR0), r)); sty(tpl_data, ws.cell(r, OUR0 + 1))
        ws.cell(r, OUR0 + 2, '=IF(ABS(%s%d)<1,"Matched with books",IF(%s%d=0,"In books, NOT reported in 3B",IF(%s%d=0,"Reported in 3B, NOT in books","3B differs from books - see S3 SR vs 3B MoM")))' % (L(OUR0 + 1), r, L(OUR0), r, L(D0 + i31a), r))
        month_rows[m].append(r)
    r += 1
    ws.cell(r, 1, name); ws.cell(r, 2, gstin); ws.cell(r, 3, "TOTAL FY 2025-26")
    for c in range(1, 4): sty(tpl_totlab, ws.cell(r, c))
    for c in list(range(D0, TIE0 + 2)) + [OUR0, OUR0 + 1]:
        ws.cell(r, c, "=SUM(%s%d:%s%d)" % (L(c), first, L(c), r - 1)); sty(tpl_tot, ws.cell(r, c))
    ws.cell(r, TIE0 + 2, "=ROUND(%s%d-%s%d,2)" % (L(TIE0), r, L(TIE0 + 1), r)); sty(tpl.cell(19, 73), ws.cell(r, TIE0 + 2))
# ALL STATES block
r += 1
for mi, m in enumerate(MONTHS + ["TOTAL FY 2025-26"]):
    r += 1
    ws.cell(r, 1, "ALL STATES"); ws.cell(r, 2, "PAN AAECR0503Q"); ws.cell(r, 3, m)
    for c in range(1, 4): sty(tpl_totlab, ws.cell(r, c))
    for c in list(range(D0, TIE0 + 2)) + [OUR0, OUR0 + 1]:
        if mi < 12: ws.cell(r, c, "=" + "+".join("%s%d" % (L(c), rr) for rr in month_rows[m]))
        else: ws.cell(r, c, "=SUM(%s%d:%s%d)" % (L(c), r - 12, L(c), r - 1))
        sty(tpl_tot, ws.cell(r, c))
    ws.cell(r, TIE0 + 2, "=ROUND(%s%d-%s%d,2)" % (L(TIE0), r, L(TIE0 + 1), r)); sty(tpl.cell(19, 73), ws.cell(r, TIE0 + 2))
RN = r
# notes (template rows 166-177 text, adapted)
r += 2
notes = [tpl.cell(x, 1).value for x in range(166, 178) if tpl.cell(x, 1).value]
for t in notes:
    t = t.replace("rows 4:70", "rows 4 to the sheet end").replace("4 to 70", "4 to the sheet end").replace("11 state", "18 state").replace("(03 Punjab to 37 Andhra Pradesh)", "(01 J&K to 37 Andhra Pradesh)")
    t = t.replace("Merged_GSTR-3B.xlsx, as provided. The 11 state sheets are retained untouched in this workbook", "the 18 Octa AnnualReport files (Portal Reports\\GSTR-3B); their GSTR-3B matrices are retained as hidden 'GSTR-3B XX' sheets in this workbook")
    ws.cell(r, 1, t).font = copy(tpl.cell(167, 1).font); r += 1
ws.cell(r, 1, "e.  Haryana (06AAECR0503Q1ZC) has no GSTR-3B file in Portal Reports - not included. Last three columns (books tie) are DPS additions, kept after the tie-out check as agreed with Priyesh.").font = copy(tpl.cell(167, 1).font)
ws.freeze_panes = "D8"; ws.auto_filter.ref = "A7:%s%d" % (L(NC), RN)
# INDEX row
ix = wb["INDEX"]; thin = Side(style="thin"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
have = {S(ix.cell(r2, 2).value) for r2 in range(5, ix.max_row + 1)}
desc = "GSTR-3B Birds Eye View - every 3B table, state x month, live over raw Octa sheets (Priyesh format)"
if desc not in have:
    r2 = ix.max_row + 1
    ix.cell(r2, 1, "Sales"); ix.cell(r2, 2, desc)
    x = ix.cell(r2, 7); x.value = '=HYPERLINK("#\'%s\'!A1","3B Birds Eye View")' % NM; x.font = Font(color="0563C1", underline="single")
    for c in range(1, 11): ix.cell(r2, c).border = BD
wb.save(P)
print("saved: %s rows 8..%d, %d data cols + 3 tie-out + 3 ours (last col %s)" % (NM, RN, ND, L(NC)))
