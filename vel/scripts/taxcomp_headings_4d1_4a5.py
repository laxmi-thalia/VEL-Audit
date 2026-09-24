"""Pawan 24-09: the Tax comp report's two difference blocks S:U and V:X take last year's headings
('Difference to be added to 4D1', 'Difference in GSTR-2B Amounts glitch-4A5'). Headings only; no figures touched."""
import shutil, time, warnings, win32com.client as win32, pythoncom
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for f in (P, B):
    try: open(f, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + f)
shutil.copy(P, "master2_snapshot_before_taxcomp_headings.xlsx")
NEW = {"S6": "Difference to be added to 4D1", "V6": "Difference in GSTR-2B Amounts glitch-4A5"}
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); ws = w.Worksheets("Tax comp report")
    assert ws.Range("S7").Value == "IGST" and ws.Range("V7").Value == "IGST" and ws.Range("R7").Value == "Impact", "layout moved"
    for a, t in NEW.items(): print("%s: %r -> %r" % (a, ws.Range(a).Value, t)); ws.Range(a).Value = t
    print("other block headings kept: Y6=%r | AB6=%r" % (ws.Range("Y6").Value, ws.Range("AB6").Value))
    e = 0
    for s_ in w.Worksheets:
        if s_.Name == "T6A1 Extract - 24-25": continue
        try: e += s_.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("error cells %d | golden %.2f" % (e, w.Worksheets("ITC Register 2025-26").Range("V4").Value)); assert e == 0
    w.Save(); w.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); print("re-open in a fresh instance: OK (%.0fs)" % (time.time() - t0)); w.SaveAs(B, FileFormat=50); w.Close(False); print("xlsb exported")
finally: xl.Quit()
