"""RCM Register POS block -> 'NA' when there is no vendor GSTIN (CA Priyesh, recording 21-09 #1: "if my vendor GSTIN is 0 / blank /
NA the remark should not come there - it will be NA"). Wraps the existing row formulas of As per State / As per Amounts / POS Check /
Query / Query Description (formula-based stays formula-based). Single COM session; verifies counts and the Statewise diff."""
import os, time, shutil, collections, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_rcmpos.xlsx")
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wb.Worksheets("RCM Register")
    hdr = [sh.Cells(5, c).Value for c in range(1, 100)]; RN = sh.Cells(sh.Rows.Count, 3).End(-4162).Row
    # the POS block is the second 'Query' pair: Vendor | My GSTN | As per State | As per Amounts | POS Check | Query | Query Description
    v = hdr.index("Vendor") + 1; block = hdr[v - 1:v + 6]
    assert block == ["Vendor", "My GSTN", "As per State", "As per Amounts", "POS Check", "Query", "Query Description"], block
    G = L(hdr.index("GSTN") + 1)
    sw = wb.Worksheets("Statewise RCM vs 3B"); sdiff0 = sw.Range("N26").Value
    gst = [x[0] for x in sh.Range("%s6:%s%d" % (G, G, RN)).Value]
    nogst = [i for i, x in enumerate(gst) if x in (None, "") or str(x).strip() in ("0", "NA", "0.0") or len(str(x).strip()) < 15]
    before = collections.Counter(str(x[0]) for x in sh.Range(sh.Cells(6, v + 4), sh.Cells(RN, v + 4)).Value)
    guard = 'OR($%s6="",TEXT($%s6,"0")="0",$%s6="NA",LEN($%s6)<15)' % (G, G, G, G)
    for off in (2, 3, 4, 5, 6):
        c = v + off; f = sh.Cells(6, c).Formula
        if not f.startswith("="): print("  %-18s has no formula - left as is" % hdr[c - 1]); continue
        if f.startswith("=IF(OR($%s6=\"\"" % G): continue                       # already wrapped
        sh.Range(sh.Cells(6, c), sh.Cells(RN, c)).Formula = "=IF(%s,\"NA\",%s)" % (guard, f[1:])
        print("  wrapped %-18s <- %s" % (hdr[c - 1], f[:90]))
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    after = collections.Counter(str(x[0]) for x in sh.Range(sh.Cells(6, v + 4), sh.Cells(RN, v + 4)).Value)
    st = [x[0] for x in sh.Range(sh.Cells(6, v + 2), sh.Cells(RN, v + 2)).Value]
    na_ok = sum(1 for i in nogst if st[i] == "NA")
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("rows without vendor GSTIN: %d -> 'As per State' = NA on %d | POS Check before %s after %s | Statewise diff %.2f -> %.2f | error cells %d" % (len(nogst), na_ok, dict(before), dict(after), sdiff0, sw.Range("N26").Value, e))
    assert e == 0 and na_ok == len(nogst) and abs(sw.Range("N26").Value - sdiff0) < 0.01
    wb.Save()
    if xlsb_free(): wb.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    wb.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
