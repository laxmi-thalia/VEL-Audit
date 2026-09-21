"""Rebuild 'RCM GL' from the client's new GL folder (Clients Data / GLs / RCM) and key everything on
POSTING DATE + DOCUMENT NUMBER (Pawan 18-09: SAP document numbers repeat across fiscal years - 1,239 of 7,231 docs recur).
Sources (sheet Data, header row 6, G/L Account per row): RCM Output GL 2025-26 / RCM Output FY 2024-25 / RCM Output Open Line
Items / RCM Input GL 1.4.2025-31.07.2026. Output accounts 26100803xx, Input accounts 19100605xx (xx: 00 CGST, 01 SGST, 02 IGST);
the input debit sits in the SAME SAP document as the output credit (3,501 / 3,503 register rows).
RCM GL: 23 data cols + GL Key (formula) + 4 LIVE match cols (Output side vs RCM Register by GL Key; Input side vs RCM Register by
GL Key + mapped output account). RCM Register: 'GL Key' + 'Found in Input GL (GL Key)' + 'Input GL Remarks (GL Key)' appended;
'Found in Output GL' / 'Output GL Remarks' become LIVE by GL Key. Register dates carried an 18:30 tz shift (Conso rows) -> normalised.
Order matters: recreate the GL sheet BEFORE writing register formulas that reference it (a delete turns refs into #REF!)."""
import time, datetime as dt, collections, warnings
import openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
D = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/GLs/RCM/"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("MASTER IS OPEN IN EXCEL - close it (without saving) and rerun")
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v or 0)
    except Exception: return 0.0
E0 = dt.datetime(1899, 12, 30)
def serial(d): return (d - E0).days + (d - E0).seconds / 86400.0
def pfy(d):
    if not isinstance(d, dt.datetime): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
GLNAME = {"2610080300": "CGST Output RCM", "2610080301": "SGST Output RCM", "2610080302": "IGST Output RCM", "1910060500": "CGST Input RCM", "1910060501": "SGST Input RCM", "1910060502": "IGST Input RCM"}
BP2ST = {"JK01": "Jammu & Kashmir", "PB01": "Punjab", "HR01": "Haryana", "RJ01": "Rajasthan", "BR01": "Bihar", "AR01": "Arunachal Pradesh", "AS01": "Assam", "WB01": "West Bengal", "JH01": "Jharkhand", "MP01": "Madhya Pradesh", "GU01": "Gujarat", "MH01": "Maharashtra", "KL01": "Kerala", "TN01": "Tamil Nadu", "TG01": "Telangana", "AP01": "Andhra Pradesh", "UP01": "Uttar Pradesh", "CG01": "Chhattisgarh", "KA01": "Karnataka", "HOIS": "HO-ISD"}
FILES = ["RCM Output GL 2025-26.xlsx", "RCM Output FY 2024-25.xlsx", "RCM Output Open Line Items.xlsx", "RCM Input GL 1.4.2025 to 31.07.2026.xlsx"]
rows = []; t0 = time.time()
for fn in FILES:
    wb = openpyxl.load_workbook(D + fn, read_only=True, data_only=True); ws = wb["Data"]
    raw = list(ws.iter_rows(values_only=True)); wb.close()
    hi = next(i for i, r in enumerate(raw) if any(S(v) == "Document Number" for v in r)); H = {S(v): j for j, v in enumerate(raw[hi]) if S(v)}
    n = 0
    for r in raw[hi + 1:]:
        doc = S(r[H["Document Number"]])
        if not doc.isdigit(): continue
        if fn.startswith("RCM Output FY 2024-25"):
            _pd = r[H["Posting Date"]]
            if not (isinstance(_pd, dt.datetime) and _pd.year == 2025 and _pd.month == 3): continue      # Mar-25 only (claimed Apr-25)
        acct = S(r[H["G/L Account"]]); acct = str(int(float(acct))) if acct.replace(".", "").isdigit() else acct
        side = "Output" if acct.startswith("26100803") else ("Input" if acct.startswith("19100605") else "")
        bp = S(r[H["Business place"]]); pdt = r[H["Posting Date"]]; ddt = r[H["Document Date"]]
        pc = S(r[H["Profit Center"]]); pc = str(int(float(pc))) if pc.replace(".", "").isdigit() else pc
        rows.append([fn.replace(".xlsx", ""), side, acct, GLNAME.get(acct, ""), BP2ST.get(bp, ""), bp, S(r[H["Year/Month"]]), S(r[H["Fiscal Year"]]), doc,
                     serial(pdt) if isinstance(pdt, dt.datetime) else None, S(r[H["Document Type"]]), serial(ddt) if isinstance(ddt, dt.datetime) else None, pfy(pdt),
                     S(r[H["Posting Key"]]), num(r[H["Amount in Local Currency"]]), S(r[H["Tax Code"]]), S(r[H["Clearing Document"]]), pc, S(r[H["Text"]]),
                     S(r[H["Offsetting Account"]]), S(r[H["Assignment"]]), S(r[H["Cost Center"]]), S(r[H["Reference"]])]); n += 1
    print("  %-45s %6d rows (%.0fs)" % (fn, n, time.time() - t0))
COLS = ["Source file", "Side", "G/L Account", "GL Name", "State", "Business place", "Year/Month", "Fiscal Year", "Document Number", "Posting Date", "Document Type", "Document Date", "Posting FY",
        "Posting Key", "Amount in Local Currency", "Tax Code", "Clearing Document", "Profit Center", "Text", "Offsetting Account", "Assignment", "Cost Center", "Reference",
        "GL Key", "Matched with RCM Register (Output)", "Match Remarks (Output)", "Matched with RCM Register (Input)", "Match Remarks (Input)"]
C = {h: L(i + 1) for i, h in enumerate(COLS)}; NG = 5 + len(rows)
print("total rows:", len(rows), "| by side:", dict(collections.Counter(r[1] for r in rows)), "| by file:", dict(collections.Counter(r[0] for r in rows)))
GLR = lambda h: "'RCM GL'!$%s$6:$%s$%d" % (C[h], C[h], NG)
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    rg = wb.Worksheets("RCM Register"); RH = {rg.Cells(5, c).Value: c for c in range(1, 100) if rg.Cells(5, c).Value}
    RN = max(r for r in range(6, rg.UsedRange.Rows.Count + 6) if rg.Cells(r, RH["MY GSTN"]).Value or rg.Cells(r, RH["Business place"]).Value)
    for h in ("GL Key", "Found in Input GL (GL Key)", "Input GL Remarks (GL Key)"):
        if h not in RH:
            k = max(RH.values()) + 1; x = rg.Cells(5, k); x.Value = h; x.Font.Bold = True; x.Font.Color = 0xFFFFFF; x.Interior.Color = 0x794E1F; rg.Columns(k).ColumnWidth = 24; RH[h] = k
    rk = lambda h: L(RH[h]); RR = lambda h: "'RCM Register'!$%s$6:$%s$%d" % (rk(h), rk(h), RN)
    # -- register dates: Conso rows were stored as IST-midnight-in-UTC (18:30 previous day). Idempotent: only .7708 fractions move.
    shifted = {}
    for h in ("Posting Date", "Document Date", "Inv. Date"):
        rng = rg.Range("%s6:%s%d" % (rk(h), rk(h), RN)); vals = rng.Value2; new = []; n = 0
        for (v,) in vals:
            if isinstance(v, (int, float)) and abs(v - int(v) - 18.5 / 24) < 1e-6: new.append([float(int(v) + 1)]); n += 1
            else: new.append([v])
        if n: rng.Value2 = new; rng.NumberFormat = "dd-mm-yyyy"
        shifted[h] = n
    print("register dates normalised (18:30 -> next midnight):", shifted)
    # -- RCM GL sheet: recreate in place (no dependents besides the register formulas written below)
    pos = wb.Worksheets("RCM GL").Index; wb.Worksheets("RCM GL").Delete()
    ws = wb.Worksheets.Add(None, wb.Worksheets(pos - 1)); ws.Name = "RCM GL"; ws.Activate(); ws.Rows(1).RowHeight = 21
    ws.Cells(2, 1).Value = "VIKRAN ENGINEERING LIMITED"; ws.Cells(2, 1).Font.Bold = True
    ws.Cells(3, 1).Value = ("RCM GL - SAP FBL3N dumps from Clients Data\\GLs\\RCM (Output FY 25-26 & FY 24-25, Output open items Apr-Jun 26, Input 1.4.25-31.7.26). "
                            "GL Key = Posting Date | Document Number (SAP document numbers repeat across fiscal years). Posting FY from Posting Date. "
                            "Output rows matched LIVE to the RCM Register on GL Key; Input rows matched on GL Key + the paired output account (input debit sits in the same SAP document).")
    ws.Cells(3, 1).Font.Bold = True
    hr = ws.Range(ws.Cells(5, 1), ws.Cells(5, len(COLS))); hr.Value = [COLS]; hr.Font.Bold = True; hr.Font.Color = 0xFFFFFF; hr.Interior.Color = 0x794E1F
    for i in range(0, len(rows), 2000):
        ch = rows[i:i + 2000]; ws.Range(ws.Cells(6 + i, 1), ws.Cells(5 + i + len(ch), 23)).Value = ch
    for h in ("Posting Date", "Document Date"): ws.Range("%s6:%s%d" % (C[h], C[h], NG)).NumberFormat = "dd-mm-yyyy"
    ws.Range("%s6:%s%d" % (C["Amount in Local Currency"], C["Amount in Local Currency"], NG)).NumberFormat = "#,##0.00"
    ws.Range("%s6:%s%d" % (C["GL Key"], C["GL Key"], NG)).Formula = '=IF($%s6="","",TEXT($%s6,"yyyymmdd")&"|"&TEXT($%s6,"0"))' % (C["Posting Date"], C["Posting Date"], C["Document Number"])
    # register account may be numeric (2610080300) or the Conso combined label '2610080300-01' (CGST+SGST): expected register amount by key
    reg_amt_for = lambda acct_expr: "SUMIFS(%s,%s,$%s6,%s,%s)" % (RR("Amount in Local Currency"), RR("GL Key"), C["GL Key"], RR("G/L Account"), acct_expr)
    ws.Range("%s6:%s%d" % (C["Matched with RCM Register (Output)"], C["Matched with RCM Register (Output)"], NG)).Formula = (
        '=IF($%s6<>"Output","",IF(COUNTIFS(%s,$%s6)>0,"Matched with RCM Register",IF($%s6<>"2025-26",IF($%s6="2024-25","FY 24-25 – Mar-25 RCM claimed Apr-25","Other period ("&$%s6&")"),IF($%s6>0,"Debit line (payment / utilisation) - not a liability booking","NOT IN RCM REGISTER (FY 25-26)"))))'
        % (C["Side"], RR("GL Key"), C["GL Key"], C["Posting FY"], C["Posting FY"], C["Posting FY"], C["Amount in Local Currency"]))
    # amount check: GL credit lines of this key+account vs register lines of this key (same account, or the combined 300-01 label for CGST/SGST)
    gl_amt = "SUMIFS(%s,%s,$%s6,%s,$%s6)" % (GLR("Amount in Local Currency"), GLR("GL Key"), C["GL Key"], GLR("G/L Account"), C["G/L Account"])
    # Register conventions (verified 18-09): Conso rows carry ONE line per document for CGST/SGST labelled 2610080300 or
    # 2610080300-01 whose amount is the CGST leg (SGST equal, implied); monthly-working rows carry separate 300 and 301 lines.
    # Expected register amount for a GL leg: 300 -> lines 300 + 300-01; 301 -> line 301 if present else mirror of the CGST leg; 302 -> line 302.
    cgst_leg = "(%s+%s)" % (reg_amt_for('"2610080300"'), reg_amt_for('"2610080300-01"'))
    has301 = "COUNTIFS(%s,$%s6,%s,\"2610080301\")>0" % (RR("GL Key"), C["GL Key"], RR("G/L Account"))
    reg_amt = 'IF(RIGHT($%s6,2)="02",%s,IF(RIGHT($%s6,2)="00",%s,IF(%s,%s,%s)))' % (C["G/L Account"], reg_amt_for('"2610080302"'), C["G/L Account"], cgst_leg, has301, reg_amt_for('"2610080301"'), cgst_leg)
    ws.Range("%s6:%s%d" % (C["Match Remarks (Output)"], C["Match Remarks (Output)"], NG)).Formula = (
        '=IF($%s6<>"Matched with RCM Register","",IF($%s6>0,"",IF(ABS(%s-%s)<0.5,"","Amount differs: GL "&TEXT(%s,"#,##0.00")&" vs register "&TEXT(%s,"#,##0.00"))))'
        % (C["Matched with RCM Register (Output)"], C["Amount in Local Currency"], gl_amt, reg_amt, gl_amt, reg_amt))
    out_acct = '"26100803"&RIGHT($%s6,2)' % C["G/L Account"]
    ws.Range("%s6:%s%d" % (C["Matched with RCM Register (Input)"], C["Matched with RCM Register (Input)"], NG)).Formula = (
        '=IF($%s6<>"Input","",IF(COUNTIFS(%s,$%s6)>0,"Matched with RCM Register (same document)",IF($%s6="2026-27","FY 26-27 posting – Apr-26 claim of Mar-26 RCM (FY 26-27 scope)",IF($%s6<>"2025-26","Other period ("&$%s6&")",IF($%s6<0,"Credit line (utilisation / reversal) - not an ITC booking","NOT IN RCM REGISTER (FY 25-26)")))))'
        % (C["Side"], RR("GL Key"), C["GL Key"], C["Posting FY"], C["Posting FY"], C["Posting FY"], C["Amount in Local Currency"]))
    reg_amt_in = reg_amt   # same leg rule: input xx00/01/02 pairs with output 300/301/302
    ws.Range("%s6:%s%d" % (C["Match Remarks (Input)"], C["Match Remarks (Input)"], NG)).Formula = (
        '=IF($%s6<>"Matched with RCM Register (same document)","",IF($%s6<0,"",IF(ABS(%s+%s)<0.5,"","Amount differs: Input GL "&TEXT(%s,"#,##0.00")&" vs register "&TEXT(-%s,"#,##0.00"))))'
        % (C["Matched with RCM Register (Input)"], C["Amount in Local Currency"], gl_amt, reg_amt_in, gl_amt, reg_amt_in))
    ws.Cells(4, C["Amount in Local Currency"]).Formula = "=SUBTOTAL(9,%s6:%s%d)" % (C["Amount in Local Currency"], C["Amount in Local Currency"], NG); ws.Cells(4, C["Amount in Local Currency"]).NumberFormat = "#,##0.00"; ws.Cells(4, C["Amount in Local Currency"]).Font.Bold = True
    ws.Range(ws.Cells(5, 1), ws.Cells(NG, len(COLS))).AutoFilter()
    for h, w in (("Source file", 30), ("Side", 8), ("G/L Account", 12), ("GL Name", 17), ("State", 18), ("Business place", 12), ("Document Number", 13), ("Posting Date", 12), ("Document Date", 12), ("Amount in Local Currency", 16), ("Text", 26), ("GL Key", 20), ("Matched with RCM Register (Output)", 40), ("Match Remarks (Output)", 40), ("Matched with RCM Register (Input)", 44), ("Match Remarks (Input)", 40)):
        ws.Columns(C[h]).ColumnWidth = w
    ws.Range("A6").Select(); xl.ActiveWindow.SplitRow = 5; xl.ActiveWindow.SplitColumn = 0; xl.ActiveWindow.FreezePanes = True
    print("RCM GL rebuilt: rows 6..%d, %d cols" % (NG, len(COLS)))
    # -- register: GL Key + LIVE output/input GL checks (written AFTER the sheet exists)
    rg.Range("%s6:%s%d" % (rk("GL Key"), rk("GL Key"), RN)).Formula = '=IF(OR($%s6="",$%s6=""),"",TEXT($%s6,"yyyymmdd")&"|"&TEXT($%s6,"0"))' % (rk("Posting Date"), rk("Document Number"), rk("Posting Date"), rk("Document Number"))
    # GL amount for this register line: same account, or (combined 300-01 label) CGST + SGST accounts
    gl_for = lambda side_acct: "SUMIFS(%s,%s,$%s6,%s,%s)" % (GLR("Amount in Local Currency"), GLR("GL Key"), rk("GL Key"), GLR("G/L Account"), side_acct)
    acct_txt = 'TEXT($%s6,"0")' % rk("G/L Account")
    # register label -> the GL leg it represents (300-01 combined label = CGST leg); register side sums its own lines of the same key + label
    gl_out = 'IF(%s="2610080300-01",%s,%s)' % (acct_txt, gl_for('"2610080300"'), gl_for(acct_txt))
    gl_in = 'IF(%s="2610080300-01",%s,%s)' % (acct_txt, gl_for('"1910060500"'), gl_for('"19100605"&RIGHT(%s,2)' % acct_txt))
    reg_self = "SUMIFS(%s,%s,$%s6,%s,%s)" % (RR("Amount in Local Currency"), RR("GL Key"), rk("GL Key"), RR("G/L Account"), acct_txt)
    rg.Range("%s6:%s%d" % (rk("Found in Output GL"), rk("Found in Output GL"), RN)).Formula = '=IF($%s6="","",IF(COUNTIFS(%s,$%s6,%s,"Output")>0,"Yes","No"))' % (rk("GL Key"), GLR("GL Key"), rk("GL Key"), GLR("Side"))
    rg.Range("%s6:%s%d" % (rk("Output GL Remarks"), rk("Output GL Remarks"), RN)).Formula = (
        '=IF($%s6="","",IF($%s6="No","Not found in Output GL (FY 25-26 / FY 24-25 / open items) by posting date + document",IF(ABS(%s-%s)<0.5,"","Amount differs: GL "&TEXT(%s,"#,##0.00")&" vs register "&TEXT(%s,"#,##0.00"))))'
        % (rk("GL Key"), rk("Found in Output GL"), gl_out, reg_self, gl_out, reg_self))
    rg.Range("%s6:%s%d" % (rk("Found in Input GL (GL Key)"), rk("Found in Input GL (GL Key)"), RN)).Formula = '=IF($%s6="","",IF(COUNTIFS(%s,$%s6,%s,"Input")>0,"Yes","No"))' % (rk("GL Key"), GLR("GL Key"), rk("GL Key"), GLR("Side"))
    rg.Range("%s6:%s%d" % (rk("Input GL Remarks (GL Key)"), rk("Input GL Remarks (GL Key)"), RN)).Formula = (
        '=IF($%s6="","",IF($%s6="No","No RCM input line in the same SAP document",IF(ABS(%s+%s)<0.5,"","Input GL "&TEXT(%s,"#,##0.00")&" vs register "&TEXT(-%s,"#,##0.00"))))'
        % (rk("GL Key"), rk("Found in Input GL (GL Key)"), gl_in, reg_self, gl_in, reg_self))
    print("register: GL Key + live Found in Output GL / Output GL Remarks / Found in Input GL (GL Key) / Input GL Remarks (GL Key) written (rows 6..%d)" % RN)
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    cnt = lambda sh, col, n: collections.Counter(str(v[0])[:60] for v in sh.Range("%s6:%s%d" % (col, col, n)).Value if v[0] not in (None, ""))
    print("error cells:", e)
    print("register Found in Output GL:", dict(cnt(rg, rk("Found in Output GL"), RN)), "| Output GL Remarks:", dict(cnt(rg, rk("Output GL Remarks"), RN).most_common(4)))
    print("register Found in Input GL (GL Key):", dict(cnt(rg, rk("Found in Input GL (GL Key)"), RN)), "| Input GL Remarks:", dict(cnt(rg, rk("Input GL Remarks (GL Key)"), RN).most_common(4)))
    print("GL output side:", dict(cnt(ws, C["Matched with RCM Register (Output)"], NG)), "| remarks:", dict(cnt(ws, C["Match Remarks (Output)"], NG).most_common(3)))
    print("GL input side:", dict(cnt(ws, C["Matched with RCM Register (Input)"], NG)), "| remarks:", dict(cnt(ws, C["Match Remarks (Input)"], NG).most_common(3)))
    print("register SAP taxable subtotal %s | Statewise RCM vs 3B diff %s" % (rg.Range("AR4").Value, wb.Worksheets("Statewise RCM vs 3B").Range("N26").Value))
    wb.Save(); wb.Close(False)
finally: xl.Quit()
print("done %.0fs" % (time.time() - t0))
