"""ITC step 0b: pour the Working extract into last year's 83-col 'ITC Register 2024-25'
layout as master sheet 'ITC Register 2025-26'. Live: Total GST, POS block. Pending blocks
left blank for their steps (2B reco cols, eligibility, 9C reporting, 8A/6A1).
Appended at end (new-cols-last rule): GST CREDIT YES/NO, F.Y/Booking Year, Company Code."""
import pandas as pd, numpy as np, datetime, time
import win32com.client as win32, pythoncom
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
df = pd.read_pickle("itc_working.pkl")
N = len(df)
ST2G = {"Jammu & Kashmir":"01AAECR0503Q1ZM","Punjab":"03AAECR0503Q1ZI","Haryana":"06AAECR0503Q1ZC",
"Rajasthan":"08AAECR0503Q1Z8","Uttar Pradesh":"09AAECR0503Q1Z6","Bihar":"10AAECR0503Q1ZN",
"Arunachal Pradesh":"12AAECR0503Q1ZJ","Assam":"18AAECR0503Q1Z7","West Bengal":"19AAECR0503Q1Z5",
"Jharkhand":"20AAECR0503Q1ZM","Chhattisgarh":"22AAECR0503Q1ZI","Madhya Pradesh":"23AAECR0503Q1ZG",
"Gujarat":"24AAECR0503Q1ZE","Maharashtra":"27AAECR0503Q1Z8","Karnataka":"29AAECR0503Q1Z4",
"Kerala":"32AAECR0503Q1ZH","Tamil Nadu":"33AAECR0503Q1ZF","TamilNadu":"33AAECR0503Q1ZF",
"Telangana":"36AAECR0503Q1Z9","Andhra Pradesh":"37AAECR0503Q1Z7","Madya Pradesh":"23AAECR0503Q1ZG"}
G2BP = {"01AAECR0503Q1ZM":"JK01","03AAECR0503Q1ZI":"PB01","06AAECR0503Q1ZC":"HR01","08AAECR0503Q1Z8":"RJ01",
"09AAECR0503Q1Z6":"UP01","10AAECR0503Q1ZN":"BR01","12AAECR0503Q1ZJ":"AR01","18AAECR0503Q1Z7":"AS01",
"19AAECR0503Q1Z5":"WB01","20AAECR0503Q1ZM":"JH01","22AAECR0503Q1ZI":"CG01","23AAECR0503Q1ZG":"MP01",
"24AAECR0503Q1ZE":"GU01","27AAECR0503Q1Z8":"MH01","29AAECR0503Q1Z4":"KA01","32AAECR0503Q1ZH":"KL01",
"33AAECR0503Q1ZF":"TN01","36AAECR0503Q1Z9":"TG01","37AAECR0503Q1Z7":"AP01"}
num = lambda c: pd.to_numeric(df[c], errors="coerce")
dt = lambda c: pd.to_datetime(df[c], errors="coerce")
def fy(d):
    if pd.isna(d): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
gstin = df["STATE NAME"].map(lambda v: ST2G.get(S(v), ""))
posting = dt("Posting Date"); docdate = dt("Document Date")
HEADERS = ["Company","Business place","State Name","VEL GSTIN","3B Claim  Month","Category",
"Category as per 3B","Document Type","Document Number","Posting Date","Posting Year","Invoice No.",
"Correct Invoice no.","Invoice Date","Correct Invoice date","Invoice Year","Vendor Code","Vendor GSTIN",
"Correct GSTIN","Vendor Name/RCM Category","Tax Rate","Taxable Value","IGST","CGST","SGST","Total GST",
"GSTR 9C_Reporting","Reasons","Matching of 12B of FY 25-26 and 12C of 24-25","Invoice Level Match",
"KEY","Countif","B_IGST","B_CGST","B_SGST","B_Total GST","KEY2","2B_IGST","2B_CGST","2B_SGST",
"2B_Total GST","D_IGST","D_CGST","D_SGST","D_Total GST","Reco Remarks","Review Remarks",
"Invoice as per 2B","2B Period","My GSTN as per 2B","Supplier GSTN as per 2B","Nature of Services",
"Type for GSTR9","G/L Account","G/L Account Text","SAP Period","GSTR 2B/6A Period","2B Year",
"Profit Center","Project Code","RCM Paid Month","Remarks","Remarks 2","Comments","Vendor","My GSTN",
"As per State","As per Amounts","POS Check","POS Query","POS Query Description","TYPE",
"Material Description","Expense GL Element","Expense Description","Eligibility","Query",
"Consider 8A reco","Considered in Table 6A1","Remarks for accounting entries- For 6A1",
"GST CREDIT YES/NO","F.Y/Booking Year","Company Code"]
col = {}
col["Company"] = df["Company"].map(S)
col["State Name"] = df["STATE NAME"].map(S)
col["VEL GSTIN"] = gstin
col["Business place"] = gstin.map(lambda g: G2BP.get(g, ""))
col["3B Claim  Month"] = df["GSTR3B Month"].map(S)
col["Category"] = df["TYPE"].map(S)
col["Document Type"] = df["Document Type"].map(S)
col["Document Number"] = df["Document Number"]
col["Posting Date"] = posting
col["Posting Year"] = posting.map(fy)
col["Invoice No."] = df["Reference"].map(S)
col["Invoice Date"] = docdate
col["Invoice Year"] = df["Invoice Year"].map(S)
col["Vendor Code"] = df["Vendor Code"]
col["Vendor GSTIN"] = df["Vendor GSTIN"].map(S)
col["Vendor Name/RCM Category"] = df["Vendor Name/RCM Category"].map(S)
col["Tax Rate"] = num("Tax Rate").round(2)
col["Taxable Value"] = num("Taxable Value").round(2)
col["IGST"] = num("IGST").round(2)
col["CGST"] = num("CGST").round(2)
col["SGST"] = num("SGST").round(2)
col["Nature of Services"] = df.apply(lambda r: S(r["Reference"]) if S(r["TYPE"]) in ("RCM","ISD") else "", axis=1)
col["G/L Account"] = df["G/L Account"]
col["G/L Account Text"] = df["G/L Account Text"].map(S)
col["SAP Period"] = df["SAP Period"].map(S)
col["GSTR 2B/6A Period"] = df["GSTR 2B/6A PERIOD"].map(S)
col["2B Year"] = df["2B YEAR"].map(S)
col["Profit Center"] = df["Profit Center"]
col["Project Code"] = df["Project Code"].map(S)
col["TYPE"] = df["Type"].map(S)
col["GST CREDIT YES/NO"] = df["GST CREDIT YES/NO"].map(S)
col["F.Y/Booking Year"] = df["F.Y/Booking Year"].map(S)
col["Company Code"] = df["Company Code"]
LIVE = {"Total GST": "=SUM(W{r}:Y{r})",
        "Vendor": "=LEFT($R{r},2)", "My GSTN": "=LEFT($D{r},2)",
        "As per State": '=IF($BM{r}=$BN{r},"Intra State","Inter State")',
        "As per Amounts": '=IF(N($W{r})=0,"Intra State","Inter State")',
        "POS Check": '=IF($BP{r}=$BO{r},TRUE,"")'}
letters = {}
def colL(n):
    s = ""
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s
for i, h in enumerate(HEADERS, 1): letters[h] = colL(i)
sums = {k: float(col[k].sum()) for k in ("Taxable Value","IGST","CGST","SGST")}
print("coerced sums:", {k: round(v, 2) for k, v in sums.items()})
grid = []
for i in range(N):
    row = []
    for h in HEADERS:
        if h in LIVE: row.append(None); continue
        s = col.get(h)
        if s is None: row.append(None); continue
        v = s.iloc[i]
        if v is None or v is pd.NaT: row.append(None)
        elif isinstance(v, pd.Timestamp): row.append(v.to_pydatetime())
        elif not isinstance(v, (str, datetime.datetime, datetime.date)) and pd.isna(v): row.append(None)
        else: row.append(v)
    grid.append(row)
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    for w in list(wb.Worksheets):
        if w.Name == "ITC Register 2025-26": w.Delete()
    anchor = wb.Worksheets("TB Scrutiny-RCM")
    ws = wb.Worksheets.Add(None, anchor); ws.Name = "ITC Register 2025-26"; ws.Activate()
    ws.Rows(1).RowHeight = 21.0
    ws.Cells(2, 1).Value = "VIKRAN ENGINEERING LIMITED"; ws.Cells(2, 1).Font.Bold = True
    ws.Cells(3, 1).Value = "Input Tax Credit (ITC) Register for FY 2025-26 - poured from 'ITC All state FY 2025-26.xlsx' (Working, Company=VEL). Pending blocks fill in their steps: 2B reco / eligibility / 9C reporting / 8A-6A1."
    ws.Cells(3, 1).Font.Bold = True
    hr = ws.Range(ws.Cells(5, 1), ws.Cells(5, len(HEADERS))); hr.Value = [HEADERS]
    hr.Font.Bold = True; hr.Font.Color = 0xFFFFFF; hr.Interior.Color = 0x794E1F
    R0 = 6; RN = R0 + N - 1
    t0 = time.time(); CH = 1500
    for i in range(0, N, CH):
        ch = grid[i:i + CH]
        ws.Range(ws.Cells(R0 + i, 1), ws.Cells(R0 + i + len(ch) - 1, len(HEADERS))).Value = ch
    print("pasted %.0fs" % (time.time() - t0), flush=True)
    for h, fpat in LIVE.items():
        ci = HEADERS.index(h) + 1
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).Formula = fpat.replace("{r}", str(R0))
    for h in ("Taxable Value","IGST","CGST","SGST","Total GST"):
        ci = HEADERS.index(h) + 1
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).NumberFormat = "#,##0.00"
        ws.Cells(4, ci).Formula = "=SUBTOTAL(9,%s%d:%s%d)" % (letters[h], R0, letters[h], RN)
        ws.Cells(4, ci).NumberFormat = "#,##0.00"; ws.Cells(4, ci).Font.Bold = True
    for h in ("Posting Date","Invoice Date"):
        ci = HEADERS.index(h) + 1
        ws.Range(ws.Cells(R0, ci), ws.Cells(RN, ci)).NumberFormat = "DD.MM.YYYY"
    ws.Range(ws.Cells(5, 1), ws.Cells(RN, len(HEADERS))).AutoFilter()
    ws.Range("A6").Select(); xl.ActiveWindow.SplitRow = 5; xl.ActiveWindow.FreezePanes = True
    for ci, w_ in ((1,9),(2,10),(3,17),(4,17),(5,13),(6,9),(9,13),(12,20),(20,26),(22,14),(23,12),(24,12),(25,12),(26,12),(81,14)):
        ws.Columns(ci).ColumnWidth = w_
    shp = ws.Shapes.AddShape(5, 2.0, 2.0, 86.0, 17.0); shp.Name = "btnIndex"
    shp.Fill.ForeColor.RGB = 0x794E1F; shp.Line.Visible = False; shp.Placement = 2
    t = shp.TextFrame2.TextRange; t.Text = "<< INDEX"; t.Font.Bold = True; t.Font.Size = 9; t.Font.Fill.ForeColor.RGB = 0xFFFFFF
    ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
    ix = wb.Worksheets("INDEX")
    last = ix.Cells(ix.Rows.Count, 2).End(-4162).Row
    have = {str(ix.Cells(r, 2).Value or "").strip() for r in range(5, last + 1)}
    if "ITC Register FY 2025-26" not in have:
        r = last + 1
        ix.Cells(r, 1).Value = "ITC"; ix.Cells(r, 2).Value = "ITC Register FY 2025-26"
        ix.Hyperlinks.Add(Anchor=ix.Cells(r, 7), Address="", SubAddress="'ITC Register 2025-26'!A1", TextToDisplay="ITC Register")
        for c in range(1, 11): ix.Cells(r, c).Borders.LineStyle = 1
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    fails = []
    def ck(nm, got, exp, tol=0.05):
        ok = got is not None and abs(got - exp) <= tol
        if not ok: fails.append(nm)
        print("  %-30s %18s %s" % (nm, format(got, ",.2f") if isinstance(got, (int, float)) else got, "OK" if ok else "*** FAIL exp %s" % exp))
    for h in ("Taxable Value","IGST","CGST","SGST"):
        ci = HEADERS.index(h) + 1
        ck(h, ws.Cells(4, ci).Value, sums[h])
    ck("Total GST", ws.Cells(4, HEADERS.index("Total GST") + 1).Value, sums["IGST"] + sums["CGST"] + sums["SGST"])
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
