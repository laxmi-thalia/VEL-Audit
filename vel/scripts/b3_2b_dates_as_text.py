"""Pawan 24-09: the 8 date-typed invoice numbers on 'GSTR-2B ITC Data' are rewritten as TEXT in the same dd-mm-yyyy form the
2B export shows, so they stop being date values but look identical. The Reason flag stays. The KEY formula passes text through,
so the row key becomes the digits of that text on both the sheet and in cascade_fix (consistent with each other)."""
import datetime as dt, shutil, time, warnings, openpyxl, win32com.client as win32, pythoncom
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for f in (P, B):
    try: open(f, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + f)
shutil.copy(P, "master2_snapshot_before_b3_text.xlsx")
S = lambda v: "" if v is None else str(v).strip(); SHEET = "GSTR-2B ITC Data"
wb = openpyxl.load_workbook(P, read_only=True, data_only=True); ws = wb[SHEET]
H = {S(c.value): i + 1 for i, c in enumerate(next(ws.iter_rows(min_row=5, max_row=5))) if c.value}; INV = H["Invoice number"]; KEY = H[next(h for h in H if h.startswith("KEY ("))]
rows = [(i, x[INV - 1]) for i, x in enumerate(ws.iter_rows(min_row=6, values_only=True), 6) if x[0] and isinstance(x[INV - 1], dt.datetime)]
wb.close(); print("date-typed invoice numbers:", len(rows)); assert len(rows) == 8, rows
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = w.Worksheets(SHEET)
    for i, d in rows:
        c = sh.Cells(i, INV); c.NumberFormat = "@"; c.Value = d.strftime("%d-%m-%Y")
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for s_ in w.Worksheets:
        if s_.Name == "T6A1 Extract - 24-25": continue
        try: e += s_.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    got = [(S(sh.Cells(i, INV).Value), S(sh.Cells(i, KEY).Value)[15:30]) for i, _ in rows]
    print("now:", got[:3], "| error cells %d | golden %.2f" % (e, w.Worksheets("ITC Register 2025-26").Range("V4").Value))
    assert e == 0 and all(isinstance(sh.Cells(i, INV).Value, str) for i, _ in rows)
    w.Save(); w.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); print("re-open in a fresh instance: OK (%.0fs)" % (time.time() - t0)); w.SaveAs(B, FileFormat=50); w.Close(False); print("xlsb exported")
finally: xl.Quit()
