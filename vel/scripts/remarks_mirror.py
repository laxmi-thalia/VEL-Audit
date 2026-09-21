"""2B sheet 'Reco Remarks' = the register's numbered remark for the matched document (CA Priyesh 21-09: identical numbered remarks
on both sides; only '10 – Not in 2B' (books) / '11 – Not in books' (2B) differ). LIVE formula: look the 2B KEY up in
'ITC Register 2025-26' KEY2, then 'ITC Register 2026-27' KEY2, else '11 – Not in books – FY 25-26 claims'.
Single COM session; ITC Summary 6A1/8C blocks must be identical before/after (they read other columns)."""
import os, time, shutil, collections, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_mirror.xlsx")
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    isum = wb.Worksheets("ITC Summary"); snap = lambda: [round(isum.Cells(25, c).Value or 0, 2) for c in list(range(7, 22)) + [49, 50, 51, 83, 84, 85]]; before = snap()
    r25 = wb.Worksheets("ITC Register 2025-26"); H25 = {r25.Cells(5, c).Value: c for c in range(1, 90) if r25.Cells(5, c).Value}; N25 = r25.Cells(r25.Rows.Count, 4).End(-4162).Row
    r27 = wb.Worksheets("ITC Register 2026-27"); H27 = {r27.Cells(4, c).Value: c for c in range(1, 120) if r27.Cells(4, c).Value}; N27 = r27.Cells(r27.Rows.Count, 4).End(-4162).Row
    b2 = wb.Worksheets("GSTR-2B Apr25-Aug26"); BH = {b2.Cells(2, c).Value: c for c in range(1, 80) if b2.Cells(2, c).Value}; NB = b2.Cells(b2.Rows.Count, 1).End(-4162).Row
    k25, m25 = L(H25["KEY2 (matched 2B key)"]), L(H25["Reco Remarks"]); k27, m27 = L(H27["KEY2 (matched 2B key)"]), L(H27["Reco Remarks"])
    f = ('=IFERROR(INDEX(\'ITC Register 2025-26\'!$%s$6:$%s$%d,MATCH($AW3,\'ITC Register 2025-26\'!$%s$6:$%s$%d,0)),'
         'IFERROR(INDEX(\'ITC Register 2026-27\'!$%s$5:$%s$%d,MATCH($AW3,\'ITC Register 2026-27\'!$%s$5:$%s$%d,0)),"11 – Not in books – FY 25-26 claims"))'
         % (m25, m25, N25, k25, k25, N25, m27, m27, N27, k27, k27, N27))
    b2.Range("%s3:%s%d" % (L(BH["Reco Remarks"]), L(BH["Reco Remarks"]), NB)).Formula = f
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    vals = [v[0] for v in b2.Range("%s3:%s%d" % (L(BH["Reco Remarks"]), L(BH["Reco Remarks"]), NB)).Value]
    dist = collections.Counter(str(v).split(" – ")[0] + " – " + str(v).split(" – ")[1] if v and " – " in str(v) else str(v)[:30] for v in vals)
    after = snap(); e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("2B Reco Remarks (numbered, mirrored):", dict(dist.most_common(16)))
    print("ITC Summary blocks unchanged:", before == after, "| error cells", e); assert e == 0 and before == after
    wb.Save()
    if xlsb_free(): wb.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    wb.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
