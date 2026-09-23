"""A7 (M1, Pawan 23-09): new sheet 'ITCR vs 2B - 3B month' in the RECO Format Pawan gave.

Two blocks, one per 2B year. Each block: the ITC Register side on the left (State | 3B Month | IGST | CGST | SGST) and the
GSTR-2B side on the right in the same shape, the filter criteria written above each side, and - added - a Difference
block so the reader sees the gap. Everything is a LIVE SUMIFS:
  Block 1, 2B Year 2025-26: register lines with Category = ITC and '2B Year' (col AZ) = 2025-26, by State + 3B Claim Month;
     2B side = 'GSTR-2B Apr25-Aug26' with Reverse Charge = No, ITC Eligible = Yes, FY (2B period) = 2025-26, by the state's
     GSTIN + the sheet's 3B Claim Month.
  Block 2, 2B Year 2024-25: register lines with '2B Year' = 2024-25; 2B side = 'GSTR-2B ITC Data' (the FY 24-25 base) with
     Supply Attract Reverse = No, ITC Availability = Yes, by State folder + the GSTR 3B Month filled by A12.
Rows are the (state, month) pairs that carry a value on either side. Register lines with a blank 2B Year are listed
separately below block 1 so nothing is hidden."""
import collections, datetime as dt, shutil, time, warnings, openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_a7.xlsx")
S = lambda v: "" if v is None else str(v).strip(); n = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
E0 = dt.datetime(1899, 12, 30); ser = lambda d: (d - E0).days
NAME = "ITCR vs 2B - 3B month"
wb = openpyxl.load_workbook(P, read_only=True, data_only=True)
isum = wb["ITC Summary"]; states = [(S(r[0]), S(r[2])) for r in isum.iter_rows(min_row=6, max_row=40, values_only=True) if r[2] and S(r[2])[:2].isdigit()]
st_of = {g: s for s, g in states}
reg = wb["ITC Register 2025-26"]; H = {c.value: i for i, c in enumerate(next(reg.iter_rows(min_row=5, max_row=5))) if c.value}
b2 = wb["GSTR-2B Apr25-Aug26"]; BH = {c.value: i for i, c in enumerate(next(b2.iter_rows(min_row=2, max_row=2))) if c.value}; NB = 2
od = wb["GSTR-2B ITC Data"]; OH = {c.value: i for i, c in enumerate(next(od.iter_rows(min_row=5, max_row=5))) if c.value}; NO = 5
pairs = {"2025-26": set(), "2024-25": set()}; blank_2b_year = collections.Counter()
mon = lambda v: dt.datetime(v.year, v.month, 1) if isinstance(v, dt.datetime) else None
RN = 5
for x in reg.iter_rows(min_row=6, values_only=True):
    if not x[3]: continue
    RN += 1
    if S(x[H["Category"]]).upper() != "ITC": continue
    m = mon(x[H["3B Claim  Month"]]); yr = S(x[H["2B Year"]])
    if yr in pairs and m: pairs[yr].add((S(x[H["State Name"]]), m))
    elif m and not yr: blank_2b_year[S(x[H["State Name"]])] += n(x[H["IGST"]]) + n(x[H["CGST"]]) + n(x[H["SGST"]])
for x in b2.iter_rows(min_row=3, values_only=True):
    if not x[0]: continue
    NB += 1
    if S(x[BH["Reverse Charge"]]) == "No" and S(x[BH["ITC Eligible"]]) == "Yes" and S(x[BH["FY (2B period)"]]) == "2025-26":
        m = mon(x[BH["3B Claim Month"]]); st = st_of.get(S(x[BH["Company GSTIN"]]).upper())
        if m and st: pairs["2025-26"].add((st, m))
for x in od.iter_rows(min_row=6, values_only=True):
    if not x[0]: continue
    NO += 1
    if S(x[OH["Supply Attract Reverse"]]) == "No" and S(x[OH["ITC Availability"]]) == "Yes":
        m = mon(x[OH["GSTR 3B Month"]])
        if m: pairs["2024-25"].add((S(x[OH["State folder"]]), m))
wb.close()
print("register rows %d | 2B rows %d | 2B base rows %d | pairs: 25-26 %d, 24-25 %d | register lines with blank 2B Year: %d states, tax %.2f" % (RN - 5, NB - 2, NO - 5, len(pairs["2025-26"]), len(pairs["2024-25"]), len(blank_2b_year), sum(blank_2b_year.values())))
c = lambda h: L(H[h] + 1); bc = lambda h: L(BH[h] + 1); oc = lambda h: L(OH[h] + 1)
def reg_f(col, yr, r):
    return ("=SUMIFS('ITC Register 2025-26'!$%s$6:$%s$%d,'ITC Register 2025-26'!$%s$6:$%s$%d,$A%d,'ITC Register 2025-26'!$%s$6:$%s$%d,$B%d,"
            "'ITC Register 2025-26'!$%s$6:$%s$%d,\"ITC\",'ITC Register 2025-26'!$%s$6:$%s$%d,\"%s\")"
            % (c(col), c(col), RN, c("State Name"), c("State Name"), RN, r, c("3B Claim  Month"), c("3B Claim  Month"), RN, r,
               c("Category"), c("Category"), RN, c("2B Year"), c("2B Year"), RN, yr))
def b2_f(col, r):
    return ("=SUMIFS('GSTR-2B Apr25-Aug26'!$%s$3:$%s$%d,'GSTR-2B Apr25-Aug26'!$%s$3:$%s$%d,INDEX('ITC Summary'!$C$6:$C$40,MATCH($A%d,'ITC Summary'!$A$6:$A$40,0)),"
            "'GSTR-2B Apr25-Aug26'!$%s$3:$%s$%d,$B%d,'GSTR-2B Apr25-Aug26'!$%s$3:$%s$%d,\"No\",'GSTR-2B Apr25-Aug26'!$%s$3:$%s$%d,\"Yes\",'GSTR-2B Apr25-Aug26'!$%s$3:$%s$%d,\"2025-26\")"
            % (bc(col), bc(col), NB, bc("Company GSTIN"), bc("Company GSTIN"), NB, r, bc("3B Claim Month"), bc("3B Claim Month"), NB, r,
               bc("Reverse Charge"), bc("Reverse Charge"), NB, bc("ITC Eligible"), bc("ITC Eligible"), NB, bc("FY (2B period)"), bc("FY (2B period)"), NB))
def od_f(col, r):
    return ("=SUMIFS('GSTR-2B ITC Data'!$%s$6:$%s$%d,'GSTR-2B ITC Data'!$%s$6:$%s$%d,$A%d,'GSTR-2B ITC Data'!$%s$6:$%s$%d,$B%d,"
            "'GSTR-2B ITC Data'!$%s$6:$%s$%d,\"No\",'GSTR-2B ITC Data'!$%s$6:$%s$%d,\"Yes\")"
            % (oc(col), oc(col), NO, oc("State folder"), oc("State folder"), NO, r, oc("GSTR 3B Month"), oc("GSTR 3B Month"), NO, r,
               oc("Supply Attract Reverse"), oc("Supply Attract Reverse"), NO, oc("ITC Availability"), oc("ITC Availability"), NO))
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); xl.Calculation = -4135
    for sh in w.Worksheets:
        if sh.Name == NAME: sh.Name = NAME + " (old)"
    ws = w.Worksheets.Add(After=w.Worksheets("ITC Register 2026-27")); ws.Name = NAME
    ws.Cells(1, 1).Value = "RECO Format – ITC Register FY 25-26 vs GSTR-2B, matched on the GSTR-3B claim month (A7, Pawan 23-09). Live formulas."; ws.Cells(1, 1).Font.Bold = True
    dark, blue, white = 0x4F3F33, 0xB09784, 0xFFFFFF
    def head(r, col, txt, fill):
        x = ws.Cells(r, col); x.Value = txt; x.Font.Bold = True; x.Font.Color = white; x.Interior.Color = fill
    row = 3
    for yr, side_lbl, right_f, right_src in (("2025-26", "FY 2025-26", b2_f, "GSTR-2B Apr25-Aug26"), ("2024-25", "FY 2024-25", od_f, "GSTR-2B ITC Data")):
        ws.Cells(row, 1).Value = "2B Year (Column AZ)"; ws.Cells(row, 2).Value = side_lbl; ws.Cells(row, 7).Value = "2B Year"; ws.Cells(row, 8).Value = side_lbl
        ws.Cells(row + 1, 1).Value = "Category"; ws.Cells(row + 1, 2).Value = "ITC"; ws.Cells(row + 1, 7).Value = "Reverse Charge"; ws.Cells(row + 1, 8).Value = "No"
        ws.Cells(row + 2, 7).Value = "ITC available"; ws.Cells(row + 2, 8).Value = "Yes"; ws.Cells(row + 2, 1).Value = "Source"; ws.Cells(row + 2, 2).Value = "ITC Register 2025-26"; ws.Cells(row + 2, 9).Value = right_src
        ws.Cells(row + 3, 1).Value = "ITCR"; ws.Cells(row + 3, 7).Value = "GSTR-2B"; ws.Cells(row + 3, 13).Value = "Difference (ITCR – 2B)"
        for j, h in enumerate(("State", "3B Month", "IGST", "CGST", "SGST")): head(row + 4, 1 + j, h, dark); head(row + 4, 7 + j, h, dark)
        for j, h in enumerate(("IGST", "CGST", "SGST", "Total")): head(row + 4, 13 + j, h, blue)
        keys = sorted(pairs[yr], key=lambda t: (t[0], t[1])); r0 = row + 5
        vals = [[st, ser(m), None, None, None] for st, m in keys]
        if vals: ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + len(vals) - 1, 5)).Value = vals; ws.Range(ws.Cells(r0, 7), ws.Cells(r0 + len(vals) - 1, 8)).Value = [[st, ser(m)] for st, m in keys]
        for i, (st, m) in enumerate(keys):
            r = r0 + i
            for j, col in enumerate(("IGST", "CGST", "SGST")):
                ws.Cells(r, 3 + j).Formula = reg_f(col, yr, r)
                ws.Cells(r, 9 + j).Formula = right_f(("IGST (Net)", "CGST (Net)", "SGST (Net)")[j] if yr == "2025-26" else ("Integrated Tax(₹)", "Central Tax(₹)", "State/UT Tax(₹)")[j], r)
                ws.Cells(r, 13 + j).Formula = "=%s%d-%s%d" % (L(3 + j), r, L(9 + j), r)
            ws.Cells(r, 16).Formula = "=SUM(M%d:O%d)" % (r, r)
        rl = r0 + max(len(keys), 1) - 1
        for col in (3, 4, 5, 9, 10, 11, 13, 14, 15, 16):
            t = ws.Cells(row + 3, col); t.Formula = "=SUBTOTAL(9,%s%d:%s%d)" % (L(col), r0, L(col), rl); t.Font.Bold = True; t.NumberFormat = "#,##0.00"
        for col in (2, 8): ws.Range(ws.Cells(r0, col), ws.Cells(rl, col)).NumberFormat = "mmm-yy"
        for col in (3, 4, 5, 9, 10, 11, 13, 14, 15, 16): ws.Range(ws.Cells(r0, col), ws.Cells(rl, col)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(row + 4, 1), ws.Cells(rl, 16)).AutoFilter()
        row = rl + 4
    if blank_2b_year:
        ws.Cells(row, 1).Value = "Register ITC lines with a BLANK '2B Year' (not in either block) – total tax by state"; ws.Cells(row, 1).Font.Italic = True
        for i, (st, t) in enumerate(sorted(blank_2b_year.items()), 1): ws.Cells(row + i, 1).Value = st; ws.Cells(row + i, 3).Value = round(t, 2); ws.Cells(row + i, 3).NumberFormat = "#,##0.00"
    for col, wdt in ((1, 20), (2, 10), (7, 20), (8, 10), (13, 14), (14, 14), (15, 14), (16, 14)): ws.Columns(col).ColumnWidth = wdt
    for col in (3, 4, 5, 9, 10, 11): ws.Columns(col).ColumnWidth = 14
    ws.Activate(); xl.ActiveWindow.FreezePanes = False; ws.Range("A8").Select(); xl.ActiveWindow.FreezePanes = True
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for sh in w.Worksheets:
        try: e += sh.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    tot = lambda r, cols: [round(ws.Cells(r, cc).Value or 0, 2) for cc in cols]
    print("block 1 (25-26) totals ITCR %s | 2B %s | diff %s" % (tot(6, (3, 4, 5)), tot(6, (9, 10, 11)), tot(6, (13, 14, 15))))
    r2 = 3 + 5 + len(pairs["2025-26"]) + 4 + 3
    print("block 2 (24-25) totals ITCR %s | 2B %s | diff %s" % (tot(r2, (3, 4, 5)), tot(r2, (9, 10, 11)), tot(r2, (13, 14, 15))))
    print("error cells %d | golden %.2f" % (e, w.Worksheets("ITC Register 2025-26").Range("V4").Value))
    assert e == 0 and abs(w.Worksheets("ITC Register 2025-26").Range("V4").Value - 1069969542.15) < 0.01
    w.Save(); w.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); print("re-open in a fresh instance: OK (%.0fs)" % (time.time() - t0))
    if xlsb_free(): w.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    w.Close(False)
finally: xl.Quit()
