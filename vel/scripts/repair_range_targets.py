"""Targeted repair of the ranges shifted from $6 to $20 by the row-6 insert: rewrite only the cells enumerated in
range_repair_targets.pkl (outside the register body), bulk per (sheet, column, contiguous run). The register body's
B_ columns (AB/AC/AD) are rewritten by final_countif_rule.py afterwards. Verify: 0 error cells, no $20 starts left outside AB:AD."""
import pickle, time, collections, openpyxl, warnings, re, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; open(P, "r+b").close()
T = pickle.load(open("range_repair_targets.pkl", "rb"))
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; n = 0
    for sn, cells in T.items():
        ws = wb.Worksheets(sn); bycol = collections.defaultdict(list)
        for r, c, f in cells: bycol[c].append((r, f))
        for c, lst in bycol.items():
            lst.sort(); i = 0
            while i < len(lst):
                j = i
                while j + 1 < len(lst) and lst[j + 1][0] == lst[j][0] + 1: j += 1
                ws.Range(ws.Cells(lst[i][0], c), ws.Cells(lst[j][0], c)).Formula = [[f] for _, f in lst[i:j + 1]]; n += j - i + 1; i = j + 1
        print("  %s: %d cells rewritten" % (sn, len(cells)), flush=True)
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    sh = wb.Worksheets("ITC Register 2025-26"); H = {sh.Cells(5, c).Value: c for c in range(1, 90) if sh.Cells(5, c).Value}
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("rewritten %d | error cells %d | row-4 Total GST: %s = %.2f | Net-ITC %s (%.0fs)" % (n, e, sh.Cells(4, H["Total GST"]).Formula, sh.Cells(4, H["Total GST"]).Value, [round(wb.Worksheets("ITC Summary").Cells(25, 49 + j).Value or 0, 2) for j in range(3)], time.time() - t0))
    wb.Save(); wb.Close(False)
finally: xl.Quit()
print("saved")
