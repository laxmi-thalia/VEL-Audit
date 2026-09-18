"""ITC step 1 apply: fill the register's 2B block (Correct-* cols, Invoice Level Match,
KEY live, B_ live, 2B_ stamped, D_ live) + write 'GSTR-2B ITC Data' sheet (merged 2B with
claim status). Countif stamped (live COUNTIF over 41k rows would cripple recalc - noted)."""
import pandas as pd, re, datetime, time
import win32com.client as win32, pythoncom
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
reg = pd.read_pickle("itc_working.pkl").reset_index(drop=True)
mt = pd.read_pickle("itc_reg_match.pkl").reset_index(drop=True)
b2 = pd.read_pickle("b2itc.pkl").reset_index(drop=True)
N = len(reg); NB = len(b2)
b2i = mt["b2i"]
def pull(colname):
    out = []
    col = b2[colname]
    for bi in b2i:
        out.append(None if bi is None or (isinstance(bi, float) and pd.isna(bi)) else col.iloc[int(bi)])
    return out
corr_inv = pull("Invoice number")
corr_date = pull("Invoice Date")
corr_gstin = pull("GSTIN of supplier")
p2b = pull("2B Return Period")
sup2b = corr_gstin
tax2b = {k: pull(k) for k in ("Integrated Tax(₹)", "Central Tax(₹)", "State/UT Tax(₹)")}
cnt = reg.groupby([reg["Vendor GSTIN"].map(lambda v: S(v).upper()),
                   reg["Reference"].map(lambda v: re.sub(r"[^A-Z0-9]", "", S(v).upper()))]).cumcount()
key_count = reg.groupby([reg["Vendor GSTIN"].map(lambda v: S(v).upper()),
                         reg["Reference"].map(lambda v: re.sub(r"[^A-Z0-9]", "", S(v).upper()))])["Company"].transform("size")
def col_arr(vals, dates=False):
    out = []
    for v in vals:
        if v is None or v is pd.NaT: out.append([None])
        elif isinstance(v, pd.Timestamp): out.append([v.to_pydatetime()])
        elif not isinstance(v, (str, datetime.datetime, datetime.date)) and pd.isna(v): out.append([None])
        else: out.append([v])
    return out
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    ws = wb.Worksheets("ITC Register 2025-26")
    H = {}
    for c in range(1, 90):
        h = str(ws.Cells(5, c).Value or "").strip()
        if h and h not in H: H[h] = c
    R0 = 6; RN = R0 + N - 1
    def L(n):
        s = ""
        while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
        return s
    def put(nm, vals, datefmt=False):
        ci = H[nm]
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).Value = col_arr(vals)
        if datefmt: ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).NumberFormat = "DD.MM.YYYY"
    t0 = time.time()
    put("Correct Invoice no.", corr_inv)
    put("Correct Invoice date", corr_date, datefmt=True)
    put("Correct GSTIN", corr_gstin)
    put("Invoice as per 2B", corr_inv)
    put("2B Period", p2b)
    put("Supplier GSTN as per 2B", sup2b)
    put("Invoice Level Match", list(mt["match"]))
    put("Countif", [int(x) for x in key_count])
    put("2B_IGST", tax2b["Integrated Tax(₹)"])
    put("2B_CGST", tax2b["Central Tax(₹)"])
    put("2B_SGST", tax2b["State/UT Tax(₹)"])
    print("stamps %.0fs" % (time.time() - t0), flush=True)
    # live formulas
    FR = {"KEY": "=$R{r}&$L{r}", "KEY2": "=$R{r}&$L{r}",
          "B_IGST": "=N($W{r})", "B_CGST": "=N($X{r})", "B_SGST": "=N($Y{r})",
          "B_Total GST": "=AG{r}+AH{r}+AI{r}",
          "2B_Total GST": "=N(AL{r})+N(AM{r})+N(AN{r})",
          "D_IGST": "=AG{r}-N(AL{r})", "D_CGST": "=AH{r}-N(AM{r})", "D_SGST": "=AI{r}-N(AN{r})",
          "D_Total GST": "=AJ{r}-AO{r}",
          "My GSTN as per 2B": "=$D{r}"}
    for nm, f_ in FR.items():
        ci = H[nm]
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).Formula = f_.replace("{r}", str(R0))
    for nm in ("B_IGST","B_CGST","B_SGST","B_Total GST","2B_IGST","2B_CGST","2B_SGST","2B_Total GST","D_IGST","D_CGST","D_SGST","D_Total GST"):
        ci = H[nm]
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).NumberFormat = "#,##0.00"
        ws.Cells(4, ci).Formula = "=SUBTOTAL(9,%s%d:%s%d)" % (L(ci), R0, L(ci), RN)
        ws.Cells(4, ci).NumberFormat = "#,##0.00"; ws.Cells(4, ci).Font.Bold = True
    print("register block done %.0fs" % (time.time() - t0), flush=True)
    # ---- 2B data sheet
    for w in list(wb.Worksheets):
        if w.Name == "GSTR-2B ITC Data": w.Delete()
    anchor = wb.Worksheets("ITC Register 2025-26")
    ws2 = wb.Worksheets.Add(None, anchor); ws2.Name = "GSTR-2B ITC Data"; ws2.Activate()
    ws2.Rows(1).RowHeight = 21.0
    ws2.Cells(2, 1).Value = "VIKRAN ENGINEERING LIMITED"; ws2.Cells(2, 1).Font.Bold = True
    ws2.Cells(3, 1).Value = ("GSTR-2B (cumulative, merged from the 19-state July 2026 monthly workings' 'GSTR 2B' sheets; Arunachal/J&K/Kerala had NO 2B sheet - missing). "
        "'Claimed in FY 25-26 register' + claim month are DERIVED by invoice-level matching - the client's own marking exists only for Bihar/Tamil Nadu.")
    ws2.Cells(3, 1).Font.Bold = True
    COLS = ["State folder","FY (derived)","F.Y","2B Return Period","GSTR 3B Month","Curent previous",
            "GSTIN of supplier","Trade/Legal name","Invoice number","Invoice type","Invoice Date",
            "Invoice Value(₹)","Place of supply","Supply Attract Reverse Charge","Rate","Taxable Value (₹)",
            "Integrated Tax(₹)","Central Tax(₹)","State/UT Tax(₹)","Cess(₹)","GSTR-1/5 Period",
            "ITC Availability","Reason","Source","IRN","claimed_by_reg","reg_claim_month"]
    HDR2 = [c.replace("claimed_by_reg","Claimed in FY 25-26 register?").replace("reg_claim_month","Claim month (derived)") for c in COLS]
    hr = ws2.Range(ws2.Cells(5, 1), ws2.Cells(5, len(COLS))); hr.Value = [HDR2]
    hr.Font.Bold = True; hr.Font.Color = 0xFFFFFF; hr.Interior.Color = 0x794E1F
    grid = []
    for row in b2[COLS].itertuples(index=False):
        rr = []
        for v in row:
            if v is None or v is pd.NaT: rr.append(None)
            elif isinstance(v, pd.Timestamp): rr.append(v.to_pydatetime())
            elif not isinstance(v, (str, datetime.datetime, datetime.date)) and pd.isna(v): rr.append(None)
            else: rr.append(v)
        grid.append(rr)
    CH = 1500
    for i in range(0, NB, CH):
        ch = grid[i:i + CH]
        ws2.Range(ws2.Cells(6 + i, 1), ws2.Cells(6 + i + len(ch) - 1, len(COLS))).Value = ch
    for nm in ("Invoice Value(₹)","Taxable Value (₹)","Integrated Tax(₹)","Central Tax(₹)","State/UT Tax(₹)","Cess(₹)"):
        ci = COLS.index(nm) + 1
        ws2.Range(ws2.Cells(6, ci), ws2.Cells(5 + NB, ci)).NumberFormat = "#,##0.00"
        ws2.Cells(4, ci).Formula = "=SUBTOTAL(9,%s6:%s%d)" % (L(ci), L(ci), 5 + NB)
        ws2.Cells(4, ci).NumberFormat = "#,##0.00"; ws2.Cells(4, ci).Font.Bold = True
    ws2.Range(ws2.Cells(6, COLS.index("Invoice Date") + 1), ws2.Cells(5 + NB, COLS.index("Invoice Date") + 1)).NumberFormat = "DD.MM.YYYY"
    ws2.Range(ws2.Cells(5, 1), ws2.Cells(5 + NB, len(COLS))).AutoFilter()
    ws2.Range("A6").Select(); xl.ActiveWindow.SplitRow = 5; xl.ActiveWindow.FreezePanes = True
    for ci, w_ in ((1,14),(2,11),(4,13),(5,15),(7,17),(8,24),(9,18),(11,12),(16,14),(17,13),(18,12),(19,12),(26,24),(27,18)):
        ws2.Columns(ci).ColumnWidth = w_
    shp = ws2.Shapes.AddShape(5, 2.0, 2.0, 86.0, 17.0); shp.Name = "btnIndex"
    shp.Fill.ForeColor.RGB = 0x794E1F; shp.Line.Visible = False; shp.Placement = 2
    t = shp.TextFrame2.TextRange; t.Text = "<< INDEX"; t.Font.Bold = True; t.Font.Size = 9; t.Font.Fill.ForeColor.RGB = 0xFFFFFF
    ws2.Hyperlinks.Add(Anchor=shp, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
    ix = wb.Worksheets("INDEX")
    last = ix.Cells(ix.Rows.Count, 2).End(-4162).Row
    have = {str(ix.Cells(r, 2).Value or "").strip() for r in range(5, last + 1)}
    if "GSTR-2B Data (cumulative, all states)" not in have:
        r = last + 1
        ix.Cells(r, 1).Value = "ITC"; ix.Cells(r, 2).Value = "GSTR-2B Data (cumulative, all states)"
        ix.Hyperlinks.Add(Anchor=ix.Cells(r, 7), Address="", SubAddress="'GSTR-2B ITC Data'!A1", TextToDisplay="GSTR-2B Data")
        for c in range(1, 11): ix.Cells(r, c).Borders.LineStyle = 1
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    fails = []
    def ck(nm, got, exp, tol=0.05):
        ok = got is not None and abs(got - exp) <= tol
        if not ok: fails.append(nm)
        print("  %-34s %18s %s" % (nm, format(got, ",.2f") if isinstance(got, (int, float)) else got, "OK" if ok else "*** FAIL exp %s" % exp))
    ck("B_Total GST subtotal", ws.Cells(4, H["B_Total GST"]).Value, 1069871702.15, 1)
    exp2b = float(sum(v for v in tax2b["Integrated Tax(₹)"] if v) + sum(v for v in tax2b["Central Tax(₹)"] if v) + sum(v for v in tax2b["State/UT Tax(₹)"] if v))
    ck("2B_Total GST subtotal", ws.Cells(4, H["2B_Total GST"]).Value, exp2b, 2)
    ck("D_Total = B - 2B", ws.Cells(4, H["D_Total GST"]).Value, 1069871702.15 - exp2b, 2)
    tot = 0
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells:", tot)
    if tot: fails.append("errors")
    ck("sales golden T3", wb.Worksheets("SR_2025-26").Range("T3").Value, 8721471166.37)
    ck("RCM SAP subtotal", wb.Worksheets("RCM Register").Cells(4, 44).Value, 115152683.11)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:", "PASS" if not fails else "FAIL: %s" % fails)
finally:
    xl.Quit()
