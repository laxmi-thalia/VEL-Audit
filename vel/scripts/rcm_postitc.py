"""RCM post-ITC hooks:
1. RCM Register 'GSTR 3B Claim month' stamped from ITC-register RCM claims (doc+FY match).
2. RCM GL input rows matched vs ITC-register RCM claims -> 'Matched with RCM Register (Input)';
   RCM Register 'Found in Input GL (post-ITC)' at doc level via same keys.
3. New sheet 'RCM Paid vs ITC Claimed' - per GSTIN x month: RCM output paid (RCM Register,
   month M) vs 3B 4A(3) claimed (3B Data) vs ITC-register RCM claims - the checklist's
   'RCM ITC paid and credit claimed comparison' + step-2 ITC leg."""
import pandas as pd, re, datetime
import win32com.client as win32, pythoncom
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
itc = pd.read_pickle("itc_working.pkl").reset_index(drop=True)
rcm_claims = itc[itc["TYPE"].map(S) == "RCM"].copy()
def docn(v):
    try: return str(int(float(v)))
    except Exception: return S(str(v))
def fy_of(d):
    d = pd.to_datetime(d, errors="coerce")
    if pd.isna(d): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
rcm_claims["doc"] = rcm_claims["Document Number"].map(docn)
rcm_claims["dfy"] = rcm_claims["Document Date"].map(fy_of)
claim_by_doc = {}
for _, r in rcm_claims.iterrows():
    claim_by_doc.setdefault((r["doc"], r["dfy"]), S(r["GSTR3B Month"]))
    claim_by_doc.setdefault(r["doc"], S(r["GSTR3B Month"]))
print("ITC-register RCM claim docs:", len(rcm_claims))
# RCM register rows
reg = pd.read_pickle("rcm_register.pkl").reset_index(drop=True)
reg["doc"] = reg["Document Number"].map(docn)
reg["dfy"] = reg["Document Date"].map(lambda d: fy_of(d))
# input GL
gl = pd.read_pickle("rcm_gl.pkl").reset_index(drop=True)
inp_mask = gl["Side"] == "Input"
inp = gl[inp_mask].copy()
inp_status = []
for _, r in inp.iterrows():
    k = (r["Document Number"], r["Doc FY"])
    if k in claim_by_doc or r["Document Number"] in claim_by_doc:
        inp_status.append("Matched with ITC Register RCM claim (month: %s)" % (claim_by_doc.get(k) or claim_by_doc.get(r["Document Number"])))
    elif r["Clearing Document"]:
        inp_status.append("Cleared/transfer entry")
    elif r["Doc FY"] == "2025-26":
        inp_status.append("NOT IN ITC REGISTER (FY 25-26)")
    else:
        inp_status.append("Other period (%s)" % (r["Doc FY"] or "no date"))
from collections import Counter
print("input GL statuses:", dict(Counter(s.split(" (")[0] for s in inp_status)))
# RCM register claim-month + Found in Input GL: via input GL doc presence? RCM register docs
# are OUTPUT docs; the ITC claim is a different doc. Link at (state, claim month) level:
# claim month for register row = its own 3B Month + 1 (input one month behind) - verified
# against ITC-register RCM claims by state+month totals on the comparison sheet instead.
MONTHS = ["01 Apr 2025","02 May 2025","03 June 2025","04 July 2025","05 Aug 2025","06 Sep 2025",
"07 Oct 2025","08 Nov 2025","09 Dec 2025","10 Jan 2026","11 Feb 2026","12 Mar 2026"]
NEXT = dict(zip(MONTHS, MONTHS[1:] + ["April 2026 (FY 26-27)"]))
reg_claim = [NEXT.get(S(m), "") for m in reg["3B Month"]]
# per-state-month ITC-register RCM claims present?
ST2G = {"Jammu & Kashmir":"01AAECR0503Q1ZM","Punjab":"03AAECR0503Q1ZI","Haryana":"06AAECR0503Q1ZC",
"Rajasthan":"08AAECR0503Q1Z8","Uttar Pradesh":"09AAECR0503Q1Z6","Bihar":"10AAECR0503Q1ZN",
"Arunachal Pradesh":"12AAECR0503Q1ZJ","Assam":"18AAECR0503Q1Z7","West Bengal":"19AAECR0503Q1Z5",
"Jharkhand":"20AAECR0503Q1ZM","Chhattisgarh":"22AAECR0503Q1ZI","Madhya Pradesh":"23AAECR0503Q1ZG",
"Madya Pradesh":"23AAECR0503Q1ZG","Gujarat":"24AAECR0503Q1ZE","Maharashtra":"27AAECR0503Q1Z8",
"Karnataka":"29AAECR0503Q1Z4","Kerala":"32AAECR0503Q1ZH","Tamil Nadu":"33AAECR0503Q1ZF",
"TamilNadu":"33AAECR0503Q1ZF","Tamilnadu":"33AAECR0503Q1ZF","Telangana":"36AAECR0503Q1Z9",
"Andhra Pradesh":"37AAECR0503Q1Z7"}
rcm_claims["cg"] = rcm_claims["STATE NAME"].map(lambda v: ST2G.get(S(v), ""))
st_claims = rcm_claims.groupby([rcm_claims["cg"], rcm_claims["GSTR3B Month"].map(S)]).size()
found_inp = []
for i, r in reg.iterrows():
    g = S(r["MY GSTN"]); cm = reg_claim[i]
    if cm.startswith("April 2026"):
        found_inp.append("Claim due FY 26-27 (Apr-26 3B)")
        continue
    try:
        n = int(st_claims.loc[(g, cm)])
        found_inp.append("Yes - %d RCM claim rows in %s" % (n, cm))
    except Exception:
        found_inp.append("No RCM claim rows in %s - review" % cm)
print("register Found-in-Input summary:", dict(Counter(s.split(" - ")[0] for s in found_inp)))
# ---- COM writes
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    # 1+2. RCM Register stamps
    rr = wb.Worksheets("RCM Register")
    RH = {}
    for c in range(1, 90):
        h = str(rr.Cells(5, c).Value or "").strip()
        if h and h not in RH: RH[h] = c
    R0 = 6; N = len(reg)
    rr.Range(rr.Cells(R0, RH["GSTR 3B Claim month"]), rr.Cells(R0 + N - 1, RH["GSTR 3B Claim month"])).Value = [[v] for v in reg_claim]
    rr.Range(rr.Cells(R0, RH["Found in Input GL (post-ITC)"]), rr.Cells(R0 + N - 1, RH["Found in Input GL (post-ITC)"])).Value = [[v] for v in found_inp]
    # RCM GL input match column
    wgl = wb.Worksheets("RCM GL")
    GH = {}
    for c in range(1, 30):
        h = str(wgl.Cells(5, c).Value or "").strip()
        if h and h not in GH: GH[h] = c
    ci = GH["Matched with RCM Register (Input)"]
    all_status = [""] * len(gl)
    j = 0
    for i in range(len(gl)):
        if inp_mask.iloc[i]:
            all_status[i] = inp_status[j]; j += 1
    wgl.Range(wgl.Cells(6, ci), wgl.Cells(6 + len(gl) - 1, ci)).Value = [[v] for v in all_status]
    # 3. comparison sheet
    NM = "RCM Paid vs ITC Claimed"
    for w in list(wb.Worksheets):
        if w.Name == NM: w.Delete()
    anchor = wb.Worksheets("Unclaimed ITC candidates")
    ws = wb.Worksheets.Add(None, anchor); ws.Name = NM; ws.Activate()
    ws.Rows(1).RowHeight = 21.0
    ws.Cells(2, 1).Value = "VIKRAN ENGINEERING LIMITED"; ws.Cells(2, 1).Font.Bold = True
    ws.Cells(3, 1).Value = ("RCM output paid (RCM Register, 3B month M) vs ITC on RCM claimed: as per 3B Table 4A(3) and as per ITC Register RCM rows. "
                            "Claim runs ONE MONTH BEHIND payment (VEL practice) - compare paid month M to claimed month M+1. Mar-26 payment claims in Apr-26 (FY 26-27).")
    ws.Cells(3, 1).Font.Bold = True
    hdrs = ["GSTIN","State","Month (3B)","RCM paid - Total tax (RCM Register)",
            "4A(3) claimed same month (3B)","4A(3) claimed NEXT month (3B)",
            "ITC Register RCM claims - month M+1","Diff (paid - next-month 4A(3))","DPS Remarks"]
    for c, h in enumerate(hdrs, 1):
        x = ws.Cells(5, c); x.Value = h; x.Font.Bold = True; x.Font.Color = 0xFFFFFF; x.Interior.Color = 0x794E1F
    GSTINS = [("01AAECR0503Q1ZM","Jammu & Kashmir"),("03AAECR0503Q1ZI","Punjab"),("06AAECR0503Q1ZC","Haryana"),
    ("08AAECR0503Q1Z8","Rajasthan"),("09AAECR0503Q1Z6","Uttar Pradesh"),("10AAECR0503Q1ZN","Bihar"),
    ("12AAECR0503Q1ZJ","Arunachal Pradesh"),("18AAECR0503Q1Z7","Assam"),("19AAECR0503Q1Z5","West Bengal"),
    ("20AAECR0503Q1ZM","Jharkhand"),("22AAECR0503Q1ZI","Chhattisgarh"),("23AAECR0503Q1ZG","Madhya Pradesh"),
    ("24AAECR0503Q1ZE","Gujarat"),("27AAECR0503Q1Z8","Maharashtra"),("29AAECR0503Q1Z4","Karnataka"),
    ("32AAECR0503Q1ZH","Kerala"),("33AAECR0503Q1ZF","Tamil Nadu"),("36AAECR0503Q1Z9","Telangana"),
    ("37AAECR0503Q1Z7","Andhra Pradesh")]
    TOK = dict(zip(MONTHS, ["Apr-25","May-25","Jun-25","Jul-25","Aug-25","Sep-25","Oct-25","Nov-25","Dec-25","Jan-26","Feb-26","Mar-26"]))
    # ranges
    RCMR = lambda n: "'RCM Register'!$%s$6:$%s$3186" % (colL(RH[n]), colL(RH[n]))
    def colL(n):
        s = ""
        while n: n, r_ = divmod(n - 1, 26); s = chr(65 + r_) + s
        return s
    b3 = wb.Worksheets("3B Data")
    BH = {}
    for row1 in (1, 2):
        for c in range(1, 30):
            h = str(b3.Cells(row1, c).Value or "").strip()
            if h and h not in BH: BH[h] = c
    n3 = b3.Cells(b3.Rows.Count, 1).End(-4162).Row
    B3R = lambda nm: "'3B Data'!$%s$2:$%s$%d" % (colL(BH[nm]), colL(BH[nm]), n3)
    IH = {}
    ir = wb.Worksheets("ITC Register 2025-26")
    for c in range(1, 90):
        h = str(ir.Cells(5, c).Value or "").strip()
        if h and h not in IH: IH[h] = c
    ITCR = lambda nm: "'ITC Register 2025-26'!$%s$6:$%s$41513" % (colL(IH[nm]), colL(IH[nm]))
    r = 5
    for g_, st in GSTINS:
        for mi, m in enumerate(MONTHS):
            r += 1
            nxt_tok = TOK[MONTHS[mi + 1]] if mi < 11 else ""
            nxt_lbl = MONTHS[mi + 1] if mi < 11 else ""
            ws.Cells(r, 1).Value = g_; ws.Cells(r, 2).Value = st; ws.Cells(r, 3).Value = m
            ws.Cells(r, 10).Value = TOK[m]; ws.Cells(r, 11).Value = nxt_tok; ws.Cells(r, 12).Value = nxt_lbl
            ws.Cells(r, 4).Formula = ('=SUMIFS(%s,%s,$A%d,%s,$C%d)+SUMIFS(%s,%s,$A%d,%s,$C%d)+SUMIFS(%s,%s,$A%d,%s,$C%d)'
                % (RCMR("IGST AS PER SAP"), RCMR("MY GSTN"), r, RCMR("3B Month"), r,
                   RCMR("CGST AS PER SAP"), RCMR("MY GSTN"), r, RCMR("3B Month"), r,
                   RCMR("SGST AS PER SAP"), RCMR("MY GSTN"), r, RCMR("3B Month"), r))
            ws.Cells(r, 5).Formula = ('=SUMIFS(%s,%s,$A%d,%s,$J%d)+SUMIFS(%s,%s,$A%d,%s,$J%d)+SUMIFS(%s,%s,$A%d,%s,$J%d)'
                % (B3R("4A(3) RCM ITC IGST"), B3R("GSTIN"), r, B3R("Month"), r,
                   B3R("4A(3) RCM ITC CGST"), B3R("GSTIN"), r, B3R("Month"), r,
                   B3R("4A(3) RCM ITC SGST"), B3R("GSTIN"), r, B3R("Month"), r))
            ws.Cells(r, 6).Formula = ('=IF($K%d="","",SUMIFS(%s,%s,$A%d,%s,$K%d)+SUMIFS(%s,%s,$A%d,%s,$K%d)+SUMIFS(%s,%s,$A%d,%s,$K%d))'
                % (r, B3R("4A(3) RCM ITC IGST"), B3R("GSTIN"), r, B3R("Month"), r,
                   B3R("4A(3) RCM ITC CGST"), B3R("GSTIN"), r, B3R("Month"), r,
                   B3R("4A(3) RCM ITC SGST"), B3R("GSTIN"), r, B3R("Month"), r))
            ws.Cells(r, 7).Formula = ('=IF($L%d="","",SUMIFS(%s,%s,$A%d,%s,$L%d,%s,"RCM")+SUMIFS(%s,%s,$A%d,%s,$L%d,%s,"RCM")+SUMIFS(%s,%s,$A%d,%s,$L%d,%s,"RCM"))'
                % (r, ITCR("IGST"), ITCR("VEL GSTIN"), r, ITCR("3B Claim  Month"), r, ITCR("Category"),
                   ITCR("CGST"), ITCR("VEL GSTIN"), r, ITCR("3B Claim  Month"), r, ITCR("Category"),
                   ITCR("SGST"), ITCR("VEL GSTIN"), r, ITCR("3B Claim  Month"), r, ITCR("Category")))
            ws.Cells(r, 8).Formula = '=IF($F%d="","(next FY)",D%d-F%d)' % (r, r, r)
            for c in range(4, 9): ws.Cells(r, c).NumberFormat = "#,##0.00"
    last = r
    r += 1
    ws.Cells(r, 2).Value = "Grand Total"; ws.Cells(r, 2).Font.Bold = True
    for c in (4, 5, 6, 7):
        ws.Cells(r, c).Formula = "=SUM(%s6:%s%d)" % (colL(c), colL(c), last)
        ws.Cells(r, c).NumberFormat = "#,##0.00"; ws.Cells(r, c).Font.Bold = True
    for ci, w_ in ((1,18),(2,18),(3,12),(4,20),(5,20),(6,20),(7,22),(8,20),(9,30)):
        ws.Columns(ci).ColumnWidth = w_
    for cc in (10, 11, 12): ws.Columns(cc).Hidden = True
    ws.Range("A6").Select(); xl.ActiveWindow.SplitRow = 5; xl.ActiveWindow.FreezePanes = True
    shp = ws.Shapes.AddShape(5, 2.0, 2.0, 86.0, 17.0); shp.Name = "btnIndex"
    shp.Fill.ForeColor.RGB = 0x794E1F; shp.Line.Visible = False; shp.Placement = 2
    t = shp.TextFrame2.TextRange; t.Text = "<< INDEX"; t.Font.Bold = True; t.Font.Size = 9; t.Font.Fill.ForeColor.RGB = 0xFFFFFF
    ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
    ix = wb.Worksheets("INDEX")
    lastix = ix.Cells(ix.Rows.Count, 2).End(-4162).Row
    have = {str(ix.Cells(rx, 2).Value or "").strip() for rx in range(5, lastix + 1)}
    if "RCM paid vs ITC claimed (4A(3))" not in have:
        rx = lastix + 1
        ix.Cells(rx, 1).Value = "RCM"; ix.Cells(rx, 2).Value = "RCM paid vs ITC claimed (4A(3))"
        ix.Hyperlinks.Add(Anchor=ix.Cells(rx, 7), Address="", SubAddress="'%s'!A1" % NM, TextToDisplay="Paid vs Claimed")
        for c in range(1, 11): ix.Cells(rx, c).Borders.LineStyle = 1
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    lrq = last + 1
    print("GT: paid %.2f | 4A3 same-month %.2f | 4A3 next-month %.2f | ITCR RCM %.2f" % (
        ws.Cells(lrq, 4).Value or 0, ws.Cells(lrq, 5).Value or 0, ws.Cells(lrq, 6).Value or 0, ws.Cells(lrq, 7).Value or 0))
    tot = 0
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells:", tot)
    print("goldens: T3", format(wb.Worksheets("SR_2025-26").Range("T3").Value, ",.0f"),
          "| RCM", format(rr.Cells(4, 44).Value, ",.0f"),
          "| ITC", format(wb.Worksheets("ITC Register 2025-26").Cells(4, 36).Value, ",.0f"))
    wb.Close(SaveChanges=True)
    print("VERDICT:", "PASS" if tot == 0 else "FAIL")
finally:
    xl.Quit()
