"""ITC Register 2026-27 rebuilt from the client's Apr-Jun 26 state working 'Inputs' sheets (fy2627_inputs.pkl, already parsed):
rows dated FY 25-26 whose 'GSTR 2B PERIOD' falls in FY 26-27 = FY 25-26 invoices whose ITC is taken in FY 26-27 (Table 13 / 12C
detail) - with the real SAP Document Type / Number / Posting Date / vendor GSTIN / Profit Center / Project Code.
Single COM session: read layout, replace rows (insert inside the range so dependents follow, then delete the old rows), recalc,
verify, save, export xlsb. Aborts before writing if the layout differs from the expected header row 4."""
import pickle, re, os, time, shutil, collections, datetime as dt, win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for p in (P, B):
    try: open(p, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED (open in Excel) - close without saving: " + p)
shutil.copy(P, "master2_snapshot_before_fy2627.xlsx")
E0 = dt.datetime(1899, 12, 30)
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v or 0)
    except Exception: return 0.0
def ser(d): return (d - E0).days if isinstance(d, dt.datetime) else None
def fy(d): return ("%d-%02d" % (d.year if d.month >= 4 else d.year - 1, (d.year + (1 if d.month >= 4 else 0)) % 100)) if isinstance(d, dt.datetime) else None
MON = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}
def period(v):
    """'01 Apr 2026' / '03 June 2026' / datetime -> (y, m) or None"""
    if isinstance(v, dt.datetime): return (v.year, v.month)
    m = re.search(r"([A-Za-z]{3})[a-z]*\s*'?\s*(\d{4}|\d{2})\b", S(v))
    if not m or m.group(1).lower() not in MON: return None
    y = int(m.group(2)); y = y + 2000 if y < 100 else y
    return (y, MON[m.group(1).lower()])
LBL = {1: "10 Jan", 2: "11 Feb", 3: "12 Mar", 4: "01 Apr", 5: "02 May", 6: "03 June", 7: "04 July", 8: "05 Aug", 9: "06 Sep", 10: "07 Oct", 11: "08 Nov", 12: "09 Dec"}
STATE2G = {"andhra pradesh": "37AAECR0503Q1Z7", "arunachal pradesh": "12AAECR0503Q1ZJ", "assam": "18AAECR0503Q1Z7", "bihar": "10AAECR0503Q1ZN", "chattisgarh": "22AAECR0503Q1ZI", "chhattisgarh": "22AAECR0503Q1ZI",
           "gujrat": "24AAECR0503Q1ZE", "gujarat": "24AAECR0503Q1ZE", "haryana": "06AAECR0503Q1ZC", "jammu & kashmir": "01AAECR0503Q1ZM", "jharkhand": "20AAECR0503Q1ZM", "karnataka": "29AAECR0503Q1Z4",
           "kerala": "32AAECR0503Q1ZH", "madhya pradesh": "23AAECR0503Q1ZG", "maharashtra": "27AAECR0503Q1Z8", "punjab": "03AAECR0503Q1ZI", "rajasthan": "08AAECR0503Q1Z8", "tamil nadu": "33AAECR0503Q1ZF",
           "telangana": "36AAECR0503Q1Z9", "uttar pradesh": "09AAECR0503Q1Z6", "west bengal": "19AAECR0503Q1Z5"}
BP2ST = {"JK01": "Jammu & Kashmir", "PB01": "Punjab", "HR01": "Haryana", "RJ01": "Rajasthan", "BR01": "Bihar", "AR01": "Arunachal Pradesh", "AS01": "Assam", "WB01": "West Bengal", "JH01": "Jharkhand", "MP01": "Madhya Pradesh", "GU01": "Gujarat", "MH01": "Maharashtra", "KL01": "Kerala", "TN01": "Tamil Nadu", "TG01": "Telangana", "AP01": "Andhra Pradesh", "UP01": "Uttar Pradesh", "CG01": "Chhattisgarh", "KA01": "Karnataka"}
ST2BP = {v: k for k, v in BP2ST.items()}
def get(r, *names):
    for n in names:
        if n in r and r[n] not in (None, ""): return r[n]
    return None
rows = pickle.load(open("fy2627_inputs.pkl", "rb"))["rows"]
sel = []; seen = set(); src_month = collections.Counter()
for r in rows:
    iy = S(get(r, "Invoice Year", "Invoice year", "FY", "Document Year")).replace("2025-26", "25-26")
    if iy != "25-26": continue
    p = period(get(r, "GSTR 2B PERIOD", "GSTR2B Month"))
    if not p or p < (2026, 4): continue
    dn = S(get(r, "Document Number")).split(".")[0]; key = (dn, S(get(r, "G/L Account", "G-L Account")), S(get(r, "Vendor GSTIN", "GSTN")).upper(), S(get(r, "Reference", "Invoice no")), round(num(get(r, "IGST")), 2), round(num(get(r, "CGST")), 2), round(num(get(r, "SGST")), 2))
    if key in seen: continue
    seen.add(key); r["_p"] = p; sel.append(r); src_month[r["_month"]] += 1
tax = lambda r: num(get(r, "IGST")) + num(get(r, "CGST")) + num(get(r, "SGST"))
print("Inputs rows %d -> FY 25-26-dated with 2B period in FY 26-27 (deduped): %d | tax %.2f | by 2B period: %s | by file: %s" % (
    len(rows), len(sel), sum(tax(r) for r in sel), {("%d-%02d" % p): round(sum(tax(r) for r in sel if r["_p"] == p)) for p in sorted({r["_p"] for r in sel})}, dict(src_month)))
def state_of(r):
    st = S(r["_state"]); bp = S(get(r, "Business place", "Business Place")) or ST2BP.get(st.title(), "")
    stn = BP2ST.get(bp, st); g = STATE2G.get(stn.lower(), STATE2G.get(st.lower(), "")); return bp, stn, g
def build(r):
    bp, stn, g = state_of(r); pdt = get(r, "Posting Date"); ddt = get(r, "Document Date"); y, m = r["_p"]
    v = {"Company": "VEL", "Business place": bp, "STATE NAME": stn, "VEL GSTN": g, "3B Claim  Month": "%s %d" % (LBL[m], y), "Category": "ITC",
         "Document Type": S(get(r, "Document Type")) or "NA", "Document Number": get(r, "Document Number"), "Posting Date": ser(pdt), "Posting Year": fy(pdt),
         "Invoice No.": S(get(r, "Reference", "Invoice no")) or "NA", "Invoice Date": ser(ddt), "Invoice Year": "2025-26", "Vendor Code": get(r, "Vendor Code"),
         "Vendor GSTIN": S(get(r, "Vendor GSTIN", "GSTN")) or "Missing", "Vendor Name/RCM Category": S(get(r, "Vendor Name")), "Tax Rate": get(r, "Tax Rate"),
         "Taxable Value": round(num(get(r, "Taxable Value", "Taxable Amt", "Taxable Amount")), 2), "IGST": round(num(get(r, "IGST")), 2), "CGST": round(num(get(r, "CGST")), 2), "SGST": round(num(get(r, "SGST")), 2),
         "G/L Account": get(r, "G/L Account", "G-L Account"), "G/L Account Text": get(r, "G/L Account Text", "G-L Account Text"), "Profit Center": get(r, "Profit Center", "Profit center"), "Project Code": get(r, "Project Code"),
         "SAP Period": S(get(r, "SAP Period")), "GSTR 2B/6A PERIOD": "%s %d" % (LBL[m], y), "GST CREDIT YES/NO": S(get(r, "GST CREDIT", "GST CREDIT YES/NO")), "Type": get(r, "Type"), "Company Code": get(r, "Company Code", "Company  Code") or 1000,
         "Source": "Client FY 26-27 working - Inputs sheet (%s, %s)" % (r["_month"], r["_state"])}
    return v
NEW = [build(r) for r in sorted(sel, key=lambda r: (r["_p"], state_of(r)[1], str(get(r, "Document Number"))))]
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wb.Worksheets("ITC Register 2026-27")
    H = {sh.Cells(4, c).Value: c for c in range(1, 80) if sh.Cells(4, c).Value}; NC = max(H.values())
    assert H.get("Company") == 1 and H.get("VEL GSTN") == 4 and H.get("3B Claim  Month") == 5 and "Document Number" in H and "Vendor GSTIN" in H and "IGST" in H, ("layout differs", H)
    RN = sh.Cells(sh.Rows.Count, 4).End(-4162).Row; assert RN >= 5
    fcols = {c: sh.Cells(5, c).FormulaR1C1 for c in range(1, NC + 1) if sh.Cells(5, c).HasFormula}; vcols = {h: c for h, c in H.items() if c not in fcols}
    if "Source" not in H: NC += 1; x = sh.Cells(4, NC); x.Value = "Source"; x.Font.Bold = True; x.Font.Color = 0xFFFFFF; x.Interior.Color = 0xB09784; vcols["Source"] = NC; sh.Columns(NC).ColumnWidth = 44
    fmt = {h: sh.Cells(5, H[h]).NumberFormat for h in ("Posting Date", "Invoice Date") if h in H}
    K = len(NEW); old_n = RN - 4
    sh.Rows("6:%d" % (6 + K - 2)).Insert(-4121) if K > 1 else None            # inside the range -> dependents expand
    for c, f in fcols.items(): sh.Range(sh.Cells(5, c), sh.Cells(4 + K, c)).FormulaR1C1 = f
    for h, c in vcols.items(): sh.Range(sh.Cells(5, c), sh.Cells(4 + K, c)).Value = [[r.get(h)] for r in NEW]
    for h, f in fmt.items(): sh.Range(sh.Cells(5, H[h]), sh.Cells(4 + K, H[h])).NumberFormat = f or "dd-mm-yyyy"
    if old_n > 1: sh.Rows("%d:%d" % (5 + K, 5 + K + old_n - 2)).Delete(-4162)     # the old rows 6..RN, now shifted below
    unmatched = [h for h in vcols if h not in NEW[0]]
    sh.Cells(2, 1).Value = ("Input Tax Credit (ITC) Register for FY 2026-27 - FY 25-26 dated invoices whose ITC is taken in FY 26-27 (Table 13 / 12C detail). "
                            "Built from the client's Apr-Jun 26 state working files, 'Inputs' sheet: Invoice Year 25-26 with GSTR 2B PERIOD in FY 26-27, de-duplicated across the monthly files (%d rows, 18-09-2026). July files carry no FY 25-26 rows." % K)
    sh.Range(sh.Cells(4, 1), sh.Cells(4 + K, NC)).AutoFilter()
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    RN2 = sh.Cells(sh.Rows.Count, 4).End(-4162).Row
    tg = [v[0] for v in sh.Range(sh.Cells(5, H["Total GST"]), sh.Cells(RN2, H["Total GST"])).Value] if "Total GST" in H else []
    cm = [v[0] for v in sh.Range(sh.Cells(5, H["3B Claim  Month"]), sh.Cells(RN2, H["3B Claim  Month"])).Value]
    bym = collections.Counter()
    for a, b in zip(cm, tg): bym[S(a)] += round(float(b or 0), 2)
    print("rows 5..%d (was %d rows) | error cells %d | Total GST by claim month: %s | value cols left blank (no Inputs source): %s" % (RN2, old_n, e, {k: round(v) for k, v in bym.items()}, unmatched))
    isum = wb.Worksheets("ITC Summary"); print("ITC Summary total row cols 58-66 (8C block area):", [round(isum.Cells(25, c).Value or 0, 2) if isinstance(isum.Cells(25, c).Value, (int, float)) else isum.Cells(25, c).Value for c in range(58, 67)])
    assert e == 0 and RN2 == 4 + K, "VERIFICATION FAILED - not saved"
    wb.Save(); wb.SaveAs(B, FileFormat=50); wb.Close(False); print("saved + xlsb exported %.1f MB (%.0fs)" % (os.path.getsize(B) / 1e6, time.time() - t0))
finally: xl.Quit()
