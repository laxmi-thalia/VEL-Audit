"""Redo of the 23-09 format revert WITHOUT cross-workbook PasteSpecial (that save produced an xlsx Excel refuses to open).

For every column the dd-mm-yy pass wrongly touched, the per-cell number formats are read from the pre-pass snapshot with
openpyxl, compressed into runs of identical format, and written back as plain NumberFormat strings through COM. Then the
saved file is re-opened in a FRESH Excel instance as proof it opens, and the xlsb is exported from that instance."""
import os, time, warnings, openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import column_index_from_string as CI
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
SNAP = "master2_snapshot_before_dates.xlsx"          # formats as they were BEFORE the dd-mm-yy pass
REVERT = {"ITC Register 2025-26": "K AY", "GSTR-2B ITC Data": "I", "T6A1 Extract - 24-25": "M",
          "RCM Register": "Q AK B BK BN BO BP BQ BR", "ITC Summary": "BS DT DU"}
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
# ---- read the original per-cell formats from the snapshot, as runs
t0 = time.time(); wb = openpyxl.load_workbook(SNAP, read_only=True)
runs = {}   # (sheet, col) -> [(row_from, row_to, fmt)]
for sh, cols in REVERT.items():
    ws = wb[sh]
    for c in cols.split():
        ci = CI(c); seq = []
        for r, row in enumerate(ws.iter_rows(min_col=ci, max_col=ci), 1):
            seq.append(getattr(row[0], "number_format", "General") or "General")
        out = []; start = 1
        for r in range(2, len(seq) + 2):
            if r > len(seq) or seq[r - 1] != seq[start - 1]:
                out.append((start, r - 1, seq[start - 1])); start = r
        runs[(sh, c)] = out
wb.close()
print("formats read from snapshot in %.0fs | runs per column: %s" % (time.time() - t0, {k[0][:12] + "!" + k[1]: len(v) for k, v in runs.items()}))
# ---- apply through COM as plain NumberFormat writes
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); xl.Calculation = -4135; n = 0
    for (sh, c), rr in runs.items():
        ws = w.Worksheets(sh)
        for a, b_, fmt in rr:
            ws.Range("%s%d:%s%d" % (c, a, c, b_)).NumberFormat = fmt; n += 1
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for s in w.Worksheets:
        try: e += s.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    rg = w.Worksheets("ITC Register 2025-26")
    print("ranges written %d | error cells %d | golden %.2f | K139 shows %r | 2B base I6 %r | RCM Q6 %r | ITC Summary BS25 %r" % (
        n, e, rg.Range("V4").Value, rg.Range("K139").Text, w.Worksheets("GSTR-2B ITC Data").Range("I6").Text,
        w.Worksheets("RCM Register").Range("Q6").Text, w.Worksheets("ITC Summary").Range("BS25").Text))
    assert e == 0 and abs(rg.Range("V4").Value - 1069969542.15) < 0.01
    w.Save(); w.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
# ---- proof: a fresh instance must open the saved file; export the xlsb from it
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); print("re-open in a fresh instance: OK, %d sheets (%.0fs)" % (w.Worksheets.Count, time.time() - t0))
    if xlsb_free(): w.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    w.Close(False)
finally: xl.Quit()
