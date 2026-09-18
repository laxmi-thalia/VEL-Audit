"""RCM step 4: 'RCM GL' sheet (output + input dumps, side-tagged) + output-leg matching
vs RCM Register. Separate Output / Input match columns on BOTH sides (input = post-ITC)."""
import pandas as pd, openpyxl, re, datetime
import win32com.client as win32, pythoncom
def S(v): return "" if v is None else str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()

GLNAME = {"2610080300":"CGST Output RCM","2610080301":"SGST Output RCM","2610080302":"IGST Output RCM",
          "1910060500":"CGST Input RCM","1910060501":"SGST Input RCM","1910060502":"IGST Input RCM"}
cc = openpyxl.load_workbook("cost_centres_new.xlsx", read_only=True, data_only=True)["Sheet4"]
PC2BP = {}
for i, r in enumerate(cc.iter_rows(values_only=True)):
    if i == 0 or r[0] is None: continue
    try: PC2BP[str(int(float(r[0])))] = S(r[6]) if len(r) > 6 else ""
    except Exception: pass
BP2ST = {"JK01":"Jammu & Kashmir","PB01":"Punjab","HR01":"Haryana","RJ01":"Rajasthan","BR01":"Bihar",
"AR01":"Arunachal Pradesh","AS01":"Assam","WB01":"West Bengal","JH01":"Jharkhand","MP01":"Madhya Pradesh",
"GU01":"Gujarat","MH01":"Maharashtra","KL01":"Kerala","TN01":"Tamil Nadu","TG01":"Telangana",
"AP01":"Andhra Pradesh","UP01":"Uttar Pradesh","CG01":"Chhattisgarh","KA01":"Karnataka","HOIS":"HO-ISD"}

def parse(fn, side):
    wb = openpyxl.load_workbook(fn, read_only=True, data_only=True)
    ws = wb["Data"]
    raw = list(ws.iter_rows(values_only=True)); wb.close()
    hdr_i = next(i for i, r in enumerate(raw) if any(S(v) == "Document Number" for v in r))
    H = {S(v): j for j, v in enumerate(raw[hdr_i]) if S(v)}
    # account markers: block END rows "Account NNNNNNNNNN"
    blocks = []
    for i, r in enumerate(raw):
        m = re.search(r"Account\s+(\d{8,})", " ".join(S(x) for x in r[:4]))
        if m: blocks.append((i, m.group(1)))
    rows = []
    bi = 0
    for i in range(hdr_i + 1, len(raw)):
        while bi < len(blocks) and i > blocks[bi][0]: bi += 1
        acct = blocks[bi][1] if bi < len(blocks) else ""
        r = raw[i]
        doc = S(r[H["Document Number"]])
        if not doc.isdigit(): continue
        dd = r[H["Document Date"]]
        pc = S(r[H["Profit Center"]])
        pc = str(int(float(pc))) if pc and pc.replace(".","").isdigit() else pc
        bp = PC2BP.get(pc, "")
        fy = ""
        if hasattr(dd, "year"):
            y = dd.year if dd.month >= 4 else dd.year - 1
            fy = "%d-%02d" % (y, (y + 1) % 100)
        rows.append({"Side": side, "G/L Account": acct, "GL Name": GLNAME.get(acct, ""),
            "Business place": bp, "State": BP2ST.get(bp, ""), "Document Number": doc,
            "Document Type": S(r[H["Document Type"]]), "Document Date": dd if hasattr(dd, "year") else None,
            "Doc FY": fy, "Posting Key": S(r[H["Posting Key"]]),
            "Amount in Local Currency": float(r[H["Amount in Local Currency"]] or 0),
            "Tax Code": S(r[H["Tax Code"]]), "Clearing Document": S(r[H["Clearing Document"]]),
            "Profit Center": pc, "Text": S(r[H["Text"]]), "Offsetting Account": S(r[H["Offsetting Account"]]),
            "Assignment": S(r[H["Assignment"]])})
    return pd.DataFrame(rows)

out = parse("rcm_gl_output.xlsx", "Output")
inp = parse("rcm_gl_input.xlsx", "Input")
gl = pd.concat([out, inp], ignore_index=True)
print("GL rows: output %d | input %d" % (len(out), len(inp)))

# ---- register keys (from pickle, same data as sheet)
res = pd.read_pickle("rcm_register.pkl")
def norm_acct(v):
    m = re.match(r"\s*(\d{8,})", str(v))
    if m: return m.group(1)
    try: return str(int(float(v)))
    except Exception: return S(str(v))
res["acct"] = res["G/L Account"].map(norm_acct)
res["doc"] = res["Document Number"].map(lambda v: str(int(float(v))) if isinstance(v,(int,float)) else S(str(v)))
def fy_of(d):
    import pandas as _pd
    if _pd.isna(d): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
res["dfy"] = res["Document Date"].map(fy_of)
reg_keys = set(zip(res["acct"], res["doc"], res["dfy"]))
reg_keys2 = set(zip(res["acct"], res["doc"]))
reg_amt = res.groupby(["acct","doc","dfy"])["Amount in Local Currency"].sum().to_dict()

def out_status(row):
    key = (row["G/L Account"], row["Document Number"], row["Doc FY"])
    if row["Side"] != "Output": return ""
    if key in reg_keys: return "Matched with RCM Register"
    if row["Doc FY"] == "" and (row["G/L Account"], row["Document Number"]) in reg_keys2:
        return "Matched with RCM Register (no date)"
    if row["Clearing Document"]: return "Cleared/transfer entry - not in register"
    if row["Doc FY"] in ("2025-26",): return "NOT IN RCM REGISTER (FY 25-26)"
    return "Other period (%s)" % (row["Doc FY"] or "no date")
gl["Matched with RCM Register (Output)"] = gl.apply(out_status, axis=1)
gl["Match Remarks (Output)"] = ""
amt_gl = gl[gl["Side"]=="Output"].groupby(["G/L Account","Document Number","Doc FY"])["Amount in Local Currency"].sum().to_dict()
mism = 0
rem = {}
for key, a in amt_gl.items():
    if key in reg_amt and abs(a - reg_amt[key]) > 0.5:
        rem[key] = "Amount differs: GL %.2f vs register %.2f" % (a, reg_amt[key]); mism += 1
gl.loc[gl["Side"]=="Output", "Match Remarks (Output)"] = gl[gl["Side"]=="Output"].apply(
    lambda r: rem.get((r["G/L Account"], r["Document Number"], r["Doc FY"]), ""), axis=1)
gl["Matched with RCM Register (Input)"] = ""   # post-ITC
gl["Match Remarks (Input)"] = ""
cnt = gl[gl["Side"]=="Output"]["Matched with RCM Register (Output)"].value_counts()
print("output-side statuses:", dict(cnt))
print("amount-mismatch docs:", mism)

# register-side stamps
gl_out_keys = set(amt_gl)
res["found_out"] = [ "Yes" if k in gl_out_keys else "No" for k in zip(res["acct"], res["doc"], res["dfy"]) ]
res["out_rem"] = [ rem.get(k, "") for k in zip(res["acct"], res["doc"], res["dfy"]) ]
print("register Found in Output GL:", dict(res["found_out"].value_counts()))
gl.to_pickle("rcm_gl.pkl"); res[["found_out","out_rem"]].to_pickle("rcm_reg_stamps.pkl")

# ---- write into master (COM)
COLS = ["Side","G/L Account","GL Name","State","Business place","Document Number","Document Type",
"Document Date","Doc FY","Posting Key","Amount in Local Currency","Tax Code","Clearing Document",
"Profit Center","Text","Offsetting Account","Assignment",
"Matched with RCM Register (Output)","Match Remarks (Output)",
"Matched with RCM Register (Input)","Match Remarks (Input)"]
grid = []
for row in gl[COLS].itertuples(index=False):
    rr = []
    for v in row:
        if v is None or (not isinstance(v,(str,datetime.datetime,datetime.date)) and pd.isna(v)): rr.append(None)
        elif isinstance(v, pd.Timestamp): rr.append(v.to_pydatetime())
        else: rr.append(v)
    grid.append(rr)
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    for w in list(wb.Worksheets):
        if w.Name == "RCM GL": w.Delete()
    anchor = wb.Worksheets("RCM ToS & Interest")
    ws = wb.Worksheets.Add(None, anchor); ws.Name = "RCM GL"; ws.Activate()
    ws.Rows(1).RowHeight = 21.0
    ws.Cells(2,1).Value = "VIKRAN ENGINEERING LIMITED"; ws.Cells(2,1).Font.Bold = True
    ws.Cells(3,1).Value = "RCM GL - Output & Input ledgers (SAP all-items dumps) vs RCM Register. Output matched now; Input column fills after the ITC register (claim months)."
    ws.Cells(3,1).Font.Bold = True
    hr = ws.Range(ws.Cells(5,1), ws.Cells(5,len(COLS))); hr.Value = [COLS]
    hr.Font.Bold = True; hr.Font.Color = 0xFFFFFF; hr.Interior.Color = 0x794E1F
    R0 = 6
    CH = 2000
    for i in range(0, len(grid), CH):
        ch = grid[i:i+CH]
        ws.Range(ws.Cells(R0+i,1), ws.Cells(R0+i+len(ch)-1,len(COLS))).Value = ch
    ws.Range(ws.Cells(R0,11), ws.Cells(R0+len(grid)-1,11)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(R0,8), ws.Cells(R0+len(grid)-1,8)).NumberFormat = "DD.MM.YYYY"
    ws.Cells(4,11).Formula = "=SUBTOTAL(9,K%d:K%d)" % (R0, R0+len(grid)-1)
    ws.Cells(4,11).NumberFormat = "#,##0.00"; ws.Cells(4,11).Font.Bold = True
    ws.Range(ws.Cells(5,1), ws.Cells(R0+len(grid)-1,len(COLS))).AutoFilter()
    ws.Range("A6").Select(); xl.ActiveWindow.SplitRow = 5; xl.ActiveWindow.FreezePanes = True
    for ci,w_ in ((1,8),(2,12),(3,17),(4,18),(5,12),(6,13),(8,12),(11,16),(15,26),(18,30),(19,34),(20,30),(21,24)):
        ws.Columns(ci).ColumnWidth = w_
    # register: rename + stamp
    rr2 = wb.Worksheets("RCM Register")
    RH = {}
    for c in range(1, 80):
        h = str(rr2.Cells(5,c).Value or "").strip()
        if h and h not in RH: RH[h] = c
    cFG = RH.get("Found in GL") or RH.get("Found in Output GL")
    cGR = RH.get("GL Remarks") or RH.get("Output GL Remarks")
    rr2.Cells(5, cFG).Value = "Found in Output GL"
    rr2.Cells(5, cGR).Value = "Output GL Remarks"
    lastc = max(RH.values())
    c_in1 = RH.get("Found in Input GL (post-ITC)", lastc+1)
    c_in2 = RH.get("Input GL Remarks", c_in1+1)
    for c, nm in ((c_in1,"Found in Input GL (post-ITC)"),(c_in2,"Input GL Remarks")):
        x = rr2.Cells(5,c); x.Value = nm; x.Font.Bold=True; x.Font.Color=0xFFFFFF; x.Interior.Color=0x794E1F
        rr2.Columns(c).ColumnWidth = 22
    stamps = pd.read_pickle("rcm_reg_stamps.pkl")
    n = len(stamps)
    rr2.Range(rr2.Cells(6, cFG), rr2.Cells(6+n-1, cFG)).Value = [[v] for v in stamps["found_out"]]
    rr2.Range(rr2.Cells(6, cGR), rr2.Cells(6+n-1, cGR)).Value = [[v] for v in stamps["out_rem"]]
    # INDEX row
    ix = wb.Worksheets("INDEX")
    have = set()
    r3 = ix.Cells(ix.Rows.Count,2).End(-4162).Row
    for r4 in range(5, r3+1): have.add(str(ix.Cells(r4,2).Value or "").strip())
    if "RCM GL (Output & Input) vs Register" not in have:
        r4 = r3+1
        ix.Cells(r4,1).Value="RCM"; ix.Cells(r4,2).Value="RCM GL (Output & Input) vs Register"
        ix.Cells(r4,7).Value="RCM GL"
        ix.Hyperlinks.Add(Anchor=ix.Cells(r4,7), Address="", SubAddress="'RCM GL'!A1", TextToDisplay="RCM GL")
        for c in range(1,11): ix.Cells(r4,c).Borders.LineStyle = 1
    # button
    shp = ws.Shapes.AddShape(5, 2.0, 2.0, 86.0, 17.0); shp.Name="btnIndex"
    shp.Fill.ForeColor.RGB=0x794E1F; shp.Line.Visible=False; shp.Placement=2
    t=shp.TextFrame2.TextRange; t.Text="<< INDEX"; t.Font.Bold=True; t.Font.Size=9; t.Font.Fill.ForeColor.RGB=0xFFFFFF
    ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells:", tot)
    print("sales golden T3:", format(wb.Worksheets("SR_2025-26").Range("T3").Value,",.2f"))
    print("RCM SAP subtotal:", format(rr2.Cells(4,44).Value,",.2f"))
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
print("RCM GL sheet written: %d rows" % len(grid))
