"""ITC Register 2025-26: 'ZFI06 status' column next to Expense Description (Pawan 18-09) - explains why Expense GL Element /
PO Number / Expense Description are blank: the client's ZFI06 export covers 1,717 documents that are not the register's.
Single COM session, verifies the lookups exist on every row, saves + exports xlsb."""
import os, time, shutil, collections, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for p in (P, B):
    try: open(p, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED - close in Excel: " + p)
shutil.copy(P, "master2_snapshot_before_zfistatus.xlsx")
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wb.Worksheets("ITC Register 2025-26")
    H = {sh.Cells(5, c).Value: c for c in range(1, 90) if sh.Cells(5, c).Value}; RN = sh.Cells(sh.Rows.Count, 4).End(-4162).Row
    ins = H["Expense Description"] + 1
    if sh.Cells(5, ins).Value != "ZFI06 status":
        sh.Columns(ins).Insert(-4161); x = sh.Cells(5, ins); x.Value = "ZFI06 status"; x.Font.Bold = True; x.Font.Color = 0xFFFFFF; x.Interior.Color = 0xB09784; sh.Columns(ins).ColumnWidth = 46
        H = {sh.Cells(5, c).Value: c for c in range(1, 90) if sh.Cells(5, c).Value}
    gl, dn, st = L(H["Expense GL Element"]), L(H["Document Number"]), L(H["ZFI06 status"])
    sh.Range("%s6:%s%d" % (st, st, RN)).Formula = '=IF($%s6="","",IF($%s6<>"","In ZFI06 export","Not in ZFI06 export - client selection covers 1,717 documents only; fresh ZFI06 run requested"))' % (dn, gl)
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print(dict(collections.Counter(str(v[0])[:22] for v in sh.Range("%s6:%s%d" % (st, st, RN)).Value)), "| error cells", e); assert e == 0
    wb.Save(); wb.SaveAs(B, FileFormat=50); wb.Close(False); print("saved + xlsb %.1f MB (%.0fs)" % (os.path.getsize(B) / 1e6, time.time() - t0))
finally: xl.Quit()
