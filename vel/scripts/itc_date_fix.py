"""ITC Register 2025-26: the 41,111 rows loaded from the client's register carry Invoice Date / Posting Date as IST-midnight-in-UTC
(18:30 of the previous day) - same tz defect as the RCM register. Normalise to midnight (serial .7708 -> next integer).
Idempotent; values only; formats kept. Pawan 18-09."""
import shutil, time, collections, win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("MASTER IS OPEN IN EXCEL - close it (without saving) and rerun")
shutil.copy(P, "master2_snapshot_before_itcdatefix.xlsx")
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wb.Worksheets("ITC Register 2025-26")
    H = {sh.Cells(5, c).Value: c for c in range(1, 90) if sh.Cells(5, c).Value}; RN = sh.Cells(sh.Rows.Count, 4).End(-4162).Row
    out = {}
    for h in ("Invoice Date", "Posting Date"):
        rng = sh.Range(sh.Cells(6, H[h]), sh.Cells(RN, H[h])); vals = rng.Value2; new = []; n = 0
        for (v,) in vals:
            if isinstance(v, (int, float)) and abs(v - int(v) - 18.5 / 24) < 1e-6: new.append([float(int(v) + 1)]); n += 1
            else: new.append([v])
        if n: rng.Value2 = new
        out[h] = n
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("normalised:", out, "| rows 6..%d | error cells %d | Total GST %.2f (%.0fs)" % (RN, e, sh.Cells(4, H["Total GST"]).Value, time.time() - t0))
    wb.Save(); wb.Close(False)
finally: xl.Quit()
print("saved")
