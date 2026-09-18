"""Apply last year's (FY 24-25 9C) colour scheme / UI to the master - styling only, no values or formulas touched.
Measured from VEL_GSTR 9_9C FY 24-25.xlsb: Calibri 11; headers bold white on #333F4F with thin bottom border; secondary /
DPS-added bands white on #8497B0; summary grids black bold on #D6DCE4; titles bold 11 no fill; Index title bold 26; no tab colours.
Phase A (openpyxl, read-only): locate the coloured cells to remap (top 300 rows of every sheet).  Phase B (COM): apply.
Excluded: '3B Birds Eye View' (Priyesh's own Arial format)."""
import shutil, time, collections, warnings, openpyxl, win32com.client as win32, pythoncom
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("MASTER IS OPEN IN EXCEL - close it (without saving) and rerun")
shutil.copy(P, "master2_snapshot_before_restyle.xlsx")
SKIP = {"3B Birds Eye View"}
HDR = {"1F4E79", "0070C0", "DDDDDD", "1F3864"}                 # -> #333F4F white bold
BAND = {"7F6000", "DDEBF7", "C0E6F5", "DCE6F1", "BDD7EE", "9BC2E6"}   # -> #8497B0 white bold
def rgb(fill):
    try:
        if fill is None or fill.fill_type != "solid": return None
        c = fill.fgColor
        if c is None or c.type != "rgb" or not c.rgb: return None
        return str(c.rgb)[-6:].upper()
    except Exception: return None
t0 = time.time(); wb = openpyxl.load_workbook(P, read_only=True); todo = {}
for ws in wb.worksheets:
    if ws.title in SKIP: continue
    h, b = [], []
    for row in ws.iter_rows(min_row=1, max_row=300):
        for c in row:
            k = rgb(c.fill)
            if k in HDR: h.append(c.coordinate)
            elif k in BAND: b.append(c.coordinate)
    if h or b: todo[ws.title] = (h, b)
wb.close(); print("scan %.0fs | sheets with cells to remap: %d | header cells %d | band cells %d" % (time.time() - t0, len(todo), sum(len(v[0]) for v in todo.values()), sum(len(v[1]) for v in todo.values())))
def bgr(hexrgb): r, g, b = int(hexrgb[0:2], 16), int(hexrgb[2:4], 16), int(hexrgb[4:6], 16); return r + g * 256 + b * 65536
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wbx = xl.Workbooks.Open(P); xl.Calculation = -4135
    st = wbx.Styles("Normal"); st.Font.Name = "Calibri"; st.Font.Size = 11
    def chunks(cells, n=30):
        for i in range(0, len(cells), n): yield ",".join(cells[i:i + n])
    for ws in wbx.Worksheets:
        if ws.Name in SKIP: continue
        ws.Tab.ColorIndex = -4142
        ws.UsedRange.Font.Name = "Calibri"           # explicit Aptos Narrow -> Calibri (sizes untouched)
        h, b = todo.get(ws.Name, ([], []))
        for spec, fill in ((h, "333F4F"), (b, "8497B0")):
            for ref in chunks(spec):
                rg = ws.Range(ref); rg.Interior.Color = bgr(fill); rg.Font.Color = bgr("FFFFFF"); rg.Font.Bold = True
                rg.Borders(9).LineStyle = 1; rg.Borders(9).Weight = 2
    ix = wbx.Worksheets("INDEX"); ix.Cells(2, 1).Font.Size = 26; ix.Cells(2, 1).Font.Bold = True
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("error cells:", e, "| ITC B_Total %s | Statewise RCM diff %s" % (wbx.Worksheets("ITC Register 2025-26").Range("AJ4").Value, wbx.Worksheets("Statewise RCM vs 3B").Range("N26").Value))
    wbx.Save(); wbx.Close(False)
finally: xl.Quit()
print("done %.0fs" % (time.time() - t0))
