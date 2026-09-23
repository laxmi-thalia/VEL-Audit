"""Pawan 23-09: every full-date column in the master displays as dd-mm-yy (month-only columns such as 'Apr-25' untouched).
Columns found by the 22/23-09 format survey (full-date formats: dd.mm.yyyy, mm-dd-yy, d-mmm-yy, dd-mmm-yy, d/mmm/yy, dd-mm-yyyy,
yyyy-mm-dd h:mm:ss, m/d/yy h:mm). Display only - no values or formulas change. Single COM session; verifies 0 error cells + golden."""
import time, shutil, win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_dates.xlsx")
COLS = {"SR_2025-26": "E Q", "Step 1 Flags": "I", "GSTR-1 Data": "F AF", "GL Data": "E S R", "S7 CN Time-bar": "C F G",
        "RCM Register": "A B BK BN BO BP BQ BR N R AL Q AK", "Month wise RCM vs 3B": "T", "RCM GL": "J L", "RCM vs 2B": "B E",
        "ITC Register 2025-26": "E AX I L K AY", "ITC Register 2026-27": "E AX I L", "GSTR-2B Apr25-Aug26": "B F AD AE AM Z AB",
        "2B ISD Apr25-Aug26": "B E R S", "GSTR-2B ITC Data": "D K I", "T6A1 Extract - 24-25": "L M", "ITC Summary": "BS DT DU",
        "ITCR vs 3B Net ITC": "AJ", "RCM Paid vs ITC Claimed": "M", "2B ISD Data": "C F", "LY 24-25 claims": "D"}
FMT = "dd-mm-yy"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    rg = wb.Worksheets("ITC Register 2025-26"); golden = rg.Range("V4").Value; done = 0
    for sh, cols in COLS.items():
        ws = wb.Worksheets(sh); last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        for c in cols.split():
            ws.Range("%s1:%s%d" % (c, c, last)).NumberFormat = FMT; done += 1
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("columns reformatted %d on %d sheets | error cells %d | golden %.2f -> %.2f | sample: register L6 shows %s, extract L5 shows %s" % (
        done, len(COLS), e, golden, rg.Range("V4").Value, rg.Range("L6").Text, wb.Worksheets("T6A1 Extract - 24-25").Range("L5").Text))
    assert e == 0 and abs(golden - rg.Range("V4").Value) < 0.01
    wb.Save()
    if xlsb_free(): wb.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    wb.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
