"""A14 (Pawan 23-09): ITC Register 2026-27 holds only invoices DATED in FY 25-26. Rows whose Invoice Date falls in
another FY (labelled 2025-26 by the client but dated April 2026) are deleted, bottom-up, inside the data range so every
dependent range shrinks with them. fy2627_reco.py + remarks_mirror.py re-run afterwards to recompute the reco block."""
import datetime as dt, shutil, time, warnings, openpyxl, win32com.client as win32, pythoncom
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
shutil.copy(P, "master2_snapshot_before_a14.xlsx")
WANT = "2025-26"
def fy(d):
    if not isinstance(d, dt.datetime): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
wb = openpyxl.load_workbook(P, read_only=True, data_only=True); n = wb["ITC Register 2026-27"]
H = {c.value: i for i, c in enumerate(next(n.iter_rows(min_row=5, max_row=5))) if c.value}
rows = [(i, r) for i, r in enumerate(n.iter_rows(min_row=6, values_only=True), 6) if r[3]]
kill = [(i, r[H["Document Number"]], fy(r[H["Invoice Date"]])) for i, r in rows if fy(r[H["Invoice Date"]]) and fy(r[H["Invoice Date"]]) != WANT]
wb.close(); print("rows", len(rows), "| to delete", len(kill), [(k[0], k[1], k[2]) for k in kill])
assert kill, "nothing to delete"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wbx = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wbx.Worksheets("ITC Register 2026-27")
    isum = wbx.Worksheets("ITC Summary"); before = isum.UsedRange.Value
    for i, doc, _ in sorted(kill, reverse=True):
        assert str(sh.Cells(i, H["Document Number"] + 1).Value).split(".")[0] == str(doc), (i, sh.Cells(i, H["Document Number"] + 1).Value, doc)
        sh.Rows(i).Delete()
    last = sh.Cells(sh.Rows.Count, 4).End(-4162).Row; assert last - 5 == len(rows) - len(kill), (last, len(rows), len(kill))
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    after = isum.UsedRange.Value
    diffs = [(i + 1, j + 1, a, b) for i, (ra, rb) in enumerate(zip(before, after)) for j, (a, b) in enumerate(zip(ra, rb))
             if a != b and not (isinstance(a, float) and isinstance(b, float) and abs(a - b) < 0.01)]
    print("deleted %d rows | last data row now %d | error cells %d | ITC Summary cells changed %d" % (len(kill), last, e, len(diffs)))
    for d in diffs[:12]: print("   ITC Summary r%d c%d: %s -> %s" % d)
    assert e == 0
    wbx.Save(); wbx.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
