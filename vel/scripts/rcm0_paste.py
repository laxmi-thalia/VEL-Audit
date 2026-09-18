"""RCM step 0b: paste rcm_register.pkl into the master as 'RCM Register' (last year's 72-col
layout, reserved row-1 button row, titles r2/r3, headers r5, data r6+). Live formula columns
set range-wide. Adds INDEX row under RCM arena + back button. Verifies totals + error scan."""
import pandas as pd, numpy as np, datetime, time
import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
res = pd.read_pickle("rcm_register.pkl")
N = len(res)
COLS = list(res.columns)          # 65 logical; display headers fix Query2 names
HDR = [c.replace("Query2", "Query").replace("Query Description2", "Query Description") for c in COLS]
R0 = 6                            # first data row
RN = R0 + N - 1
FORMULA = {
    "Vikran State code": "=LEFT($D{r},2)",
    "Vendor state code": "=LEFT($V{r},2)",
    "Key":               "=$D{r}&$V{r}",
    "Total GST":         "=SUM(AS{r}:AU{r})",
    "Rate Check":        "=IFERROR($AV{r}/$AR{r}*100,\"\")",
    "Vendor":            "=LEFT($V{r},2)",
    "My GSTN":           "=LEFT($D{r},2)",
    "As per State":      "=IF($BC{r}=$BD{r},\"Intra State\",\"Inter State\")",
    "As per Amounts":    "=IF($AS{r}=0,\"Intra State\",\"Inter State\")",
    "POS Check":         "=IF($BF{r}=$BE{r},TRUE,\"\")",
}
def colL(n):
    s = ""
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P)
    xl.Calculation = -4135
    for w in list(wb.Worksheets):
        if w.Name == "RCM Register": w.Delete()
    anchor = wb.Worksheets("S9 Sales Reco")
    ws = wb.Worksheets.Add(None, anchor); ws.Name = "RCM Register"
    ws.Activate()
    ws.Rows(1).RowHeight = 21.0
    ws.Cells(2, 1).Value = "VIKRAN ENGINEERING LIMITED"; ws.Cells(2, 1).Font.Bold = True
    ws.Cells(3, 1).Value = "RCM Payable Register FY 2025-26"; ws.Cells(3, 1).Font.Bold = True
    hr = ws.Range(ws.Cells(5, 1), ws.Cells(5, len(HDR)))
    hr.Value = [HDR]
    hr.Font.Bold = True; hr.Font.Color = 0xFFFFFF; hr.Interior.Color = 0x794E1F
    hr.WrapText = False
    # values grid: formula cols -> None; dates -> datetime; NaN -> None
    grid = []
    fcols = set(FORMULA)
    for row in res.itertuples(index=False):
        rr = []
        for c, v in zip(COLS, row):
            if c in fcols: rr.append(None)
            elif v is None or (np.isscalar(v) or isinstance(v, (pd.Timestamp, type(pd.NaT)))) and pd.isna(v): rr.append(None)
            elif isinstance(v, pd.Timestamp): rr.append(v.to_pydatetime())
            else: rr.append(v)
        grid.append(rr)
    CH = 1000
    t0 = time.time()
    for i in range(0, N, CH):
        chunk = grid[i:i + CH]
        rng = ws.Range(ws.Cells(R0 + i, 1), ws.Cells(R0 + i + len(chunk) - 1, len(COLS)))
        rng.Value = chunk
    print("values pasted %.0fs" % (time.time() - t0), flush=True)
    for c, fpat in FORMULA.items():
        ci = COLS.index(c) + 1
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).Formula = fpat.replace("{r}", str(R0))
    # formats
    for name in ("Amount in Local Currency","Taxable Value As per filed return","IGST AS PER GST PORTAL",
                 "CGST AS PER GST PORTAL","SGST AS PER GST PORTAL","Taxable Value as per SAP","IGST AS PER SAP",
                 "CGST AS PER SAP","SGST AS PER SAP","Total GST","Inv. Value"):
        ci = COLS.index(name) + 1
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).NumberFormat = "#,##0.00"
    for name in ("Final 3B Month","Posting Date","Document Date","Inv. Date"):
        ci = COLS.index(name) + 1
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).NumberFormat = "DD.MM.YYYY"
    ws.Range(ws.Cells(R0, COLS.index("Rate Check") + 1), ws.Cells(RN, COLS.index("Rate Check") + 1)).NumberFormat = "0.00"
    # subtotal row above headers (master convention: row 4)
    for name in ("Taxable Value as per SAP","IGST AS PER SAP","CGST AS PER SAP","SGST AS PER SAP","Total GST",
                 "Taxable Value As per filed return","IGST AS PER GST PORTAL","CGST AS PER GST PORTAL","SGST AS PER GST PORTAL"):
        ci = COLS.index(name) + 1
        cell = ws.Cells(4, ci)
        cell.Formula = "=SUBTOTAL(9,%s%d:%s%d)" % (colL(ci), R0, colL(ci), RN)
        cell.Font.Bold = True; cell.NumberFormat = "#,##0.00"
    ws.Range("A6").Select()
    xl.ActiveWindow.SplitRow = 5; xl.ActiveWindow.SplitColumn = 0
    xl.ActiveWindow.FreezePanes = True
    ws.Range(ws.Cells(5, 1), ws.Cells(RN, len(COLS))).AutoFilter()
    for ci, w in ((1,12),(2,11),(3,12),(4,17),(5,16),(23,16),(22,17),(11,22),(31,15),(32,20),(49,26),(50,13)):
        ws.Columns(ci).ColumnWidth = w
    # back button
    shp = ws.Shapes.AddShape(5, 2.0, 2.0, 86.0, 17.0); shp.Name = "btnIndex"
    shp.Fill.ForeColor.RGB = 0x794E1F; shp.Line.Visible = False; shp.Placement = 2
    t = shp.TextFrame2.TextRange; t.Text = "<< INDEX"; t.Font.Bold = True; t.Font.Size = 9; t.Font.Fill.ForeColor.RGB = 0xFFFFFF
    ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
    # INDEX row: RCM arena
    ix = wb.Worksheets("INDEX")
    last = ix.Cells(ix.Rows.Count, 2).End(-4162).Row
    r = last + 1
    ix.Cells(r, 1).Value = "RCM"
    ix.Cells(r, 2).Value = "RCM Payable Register FY 2025-26"
    ix.Cells(r, 7).Value = "RCM Register"
    ix.Hyperlinks.Add(Anchor=ix.Cells(r, 7), Address="", SubAddress="'RCM Register'!A1", TextToDisplay="RCM Register")
    for c in range(1, 11): ix.Cells(r, c).Borders.LineStyle = 1
    ix.Range(ix.Cells(r, 1), ix.Cells(r, 10)).Interior.Color = 0xFFFFFF
    xl.Calculation = -4105
    xl.CalculateFullRebuild()
    # ---- verification
    fails = []
    def ck(nm, got, exp, tol=0.05):
        ok = got is not None and abs(got - exp) <= tol
        if not ok: fails.append(nm)
        print("  %-38s %18s %s" % (nm, format(got, ",.2f") if isinstance(got, (int, float)) else got, "OK" if ok else "*** FAIL exp %s" % exp))
    ck("rows", float(ws.Cells(ws.Rows.Count, 13).End(-4162).Row - R0 + 1), N, 0)
    ck("SAP taxable (AR4)", ws.Cells(4, COLS.index("Taxable Value as per SAP") + 1).Value, float(res["Taxable Value as per SAP"].sum()))
    ck("SAP IGST", ws.Cells(4, COLS.index("IGST AS PER SAP") + 1).Value, float(res["IGST AS PER SAP"].sum()))
    ck("SAP CGST", ws.Cells(4, COLS.index("CGST AS PER SAP") + 1).Value, float(res["CGST AS PER SAP"].sum()))
    ck("Total GST col", ws.Cells(4, COLS.index("Total GST") + 1).Value,
       float(res[["IGST AS PER SAP","CGST AS PER SAP","SGST AS PER SAP"]].sum().sum()))
    ck("portal taxable", ws.Cells(4, COLS.index("Taxable Value As per filed return") + 1).Value, float(res["Taxable Value As per filed return"].sum()))
    tot = 0
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells:", tot)
    if tot: fails.append("errors")
    ck("sales golden: register T3", wb.Worksheets("SR_2025-26").Range("T3").Value, 8721471166.37)
    ck("sales golden: S2 diff AA75", wb.Worksheets("S2 SR vs GSTR-1").Range("AA75").Value, 16876675.46)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:", "PASS" if not fails else "FAIL: %s" % fails)
finally:
    xl.Quit()
