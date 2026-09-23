"""Undo the dd-mm-yy pass on columns that are NOT date columns (invoice/reference numbers, ITC Summary amounts, RCM mixed columns):
their original cell formats are pasted back from master2_snapshot_before_dates.xlsx (PasteSpecial formats). Values untouched."""
import os, time, shutil, win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
SNAP = os.path.abspath("master2_snapshot_before_dates.xlsx")
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_daterevert.xlsx")
REVERT = {"ITC Register 2025-26": "K AY", "GSTR-2B ITC Data": "I", "T6A1 Extract - 24-25": "M", "RCM Register": "Q AK B BK BN BO BP BQ BR",
          "ITC Summary": "BS DT DU"}
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; ws_snap = xl.Workbooks.Open(SNAP, ReadOnly=True)
    n = 0
    for sh, cols in REVERT.items():
        a, b_ = wb.Worksheets(sh), ws_snap.Worksheets(sh); last = max(a.UsedRange.Row + a.UsedRange.Rows.Count - 1, b_.UsedRange.Row + b_.UsedRange.Rows.Count - 1)
        for c in cols.split():
            b_.Range("%s1:%s%d" % (c, c, last)).Copy(); a.Range("%s1:%s%d" % (c, c, last)).PasteSpecial(-4122); n += 1
    xl.CutCopyMode = False; ws_snap.Close(False)
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    rg = wb.Worksheets("ITC Register 2025-26")
    print("columns restored %d | error cells %d | golden %.2f | register K139 shows %r | 2B base I6 shows %r | extract M5 shows %r | RCM Q6 shows %r | ITC Summary BS25 shows %r" % (
        n, e, rg.Range("V4").Value, rg.Range("K139").Text, wb.Worksheets("GSTR-2B ITC Data").Range("I6").Text, wb.Worksheets("T6A1 Extract - 24-25").Range("M5").Text,
        wb.Worksheets("RCM Register").Range("Q6").Text, wb.Worksheets("ITC Summary").Range("BS25").Text))
    assert e == 0 and abs(rg.Range("V4").Value - 1069969542.15) < 0.01
    wb.Save()
    if xlsb_free(): wb.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    wb.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
