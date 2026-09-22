"""ITC Register 2025-26: data rows that carry the dark header fill (rows inserted at row 6 inherited the header format) -> normal row
format copied from the first clean data row (fills, fonts; per-column number formats preserved). Single COM session; values untouched."""
import time, shutil, warnings, openpyxl, win32com.client as win32, pythoncom
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_unfill.xlsx")
wb = openpyxl.load_workbook(P, read_only=True); ws = wb["ITC Register 2025-26"]
dark = []; clean = None; last = 5
for i, row in enumerate(ws.iter_rows(min_row=6, max_col=1), 6):
    c = row[0]
    if c.value is None: continue
    last = i; rgb = getattr(getattr(c.fill, "fgColor", None), "rgb", None)
    if rgb and rgb.upper().endswith("333F4F"): dark.append(i)
    elif clean is None: clean = i
wb.close(); print("dark data rows:", len(dark), (dark[:5], dark[-5:]) if dark else "", "| clean template row:", clean, "| last row:", last)
if not dark: raise SystemExit("nothing to do")
# contiguous blocks
blocks = []; s = dark[0]; p = dark[0]
for r in dark[1:]:
    if r != p + 1: blocks.append((s, p)); s = r
    p = r
blocks.append((s, p)); print("blocks:", blocks[:10], "…" if len(blocks) > 10 else "")
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wbx = xl.Workbooks.Open(P); xl.Calculation = -4135; rg = wbx.Worksheets("ITC Register 2025-26")
    lastcol = rg.Cells(5, rg.Columns.Count).End(-4159).Column; golden = rg.Cells(4, 22).Value
    src = rg.Range(rg.Cells(clean, 1), rg.Cells(clean, lastcol)); src.Copy()
    for a, b in blocks: rg.Range(rg.Cells(a, 1), rg.Cells(b, lastcol)).PasteSpecial(-4122)
    xl.CutCopyMode = False
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("rows reformatted %d | first fixed row fill %s (clean row %s) | error cells %d | golden %.2f -> %.2f" % (len(dark), rg.Cells(dark[0], 1).Interior.ColorIndex, rg.Cells(clean, 1).Interior.ColorIndex, e, golden, rg.Cells(4, 22).Value))
    assert e == 0 and abs(golden - rg.Cells(4, 22).Value) < 0.01
    wbx.Save()
    if xlsb_free(): wbx.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    wbx.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
