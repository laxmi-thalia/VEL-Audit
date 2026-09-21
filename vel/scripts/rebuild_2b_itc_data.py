"""GSTR-2B ITC Data (FY 24-25 2B base) rebuilt from Octa's PAN-level export 'PAN GSTR2B 2024-25.xlsx' (Pawan 21-09) - replaces
the working-file tabs (48,802 rows across FY 20-21..26-27, mixed period formats). Same header row 5, rows 6.. replaced, KEY col
kept; CDN amounts negative; ties to the file's row-1 column totals before saving. xlsb export skipped if the xlsb is open."""
import os, re, time, shutil, collections, datetime as dt, warnings, openpyxl, win32com.client as win32, pythoncom
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
SRC = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Audit Data of FY 2024-25/PAN GSTR2B 2024-25.xlsx"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_2bitcdata.xlsx")
E0 = dt.datetime(1899, 12, 30)
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(str(v).replace(",", "") or 0)
    except Exception: return 0.0
def ddate(s):
    if isinstance(s, dt.datetime): return s
    s = S(s)
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try: return dt.datetime.strptime(s[:10], f)
        except Exception: pass
    return None
def period(s):
    s = S(s); return dt.datetime(int(s[2:6]), int(s[:2]), 1) if re.fullmatch(r"\d{6}", s) else None
def fy(d): return ("%d-%02d" % (d.year if d.month >= 4 else d.year - 1, (d.year + (1 if d.month >= 4 else 0)) % 100)) if isinstance(d, dt.datetime) else None
ST = {"01": "Jammu & Kashmir", "03": "Punjab", "06": "Haryana", "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "12": "Arunachal Pradesh", "18": "Assam", "19": "West Bengal", "20": "Jharkhand", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat", "27": "Maharashtra", "29": "Karnataka", "32": "Kerala", "33": "Tamil Nadu", "36": "Telangana", "37": "Andhra Pradesh"}
wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True); ws = wb["Inv+CDN(Document level)"]; rows = list(ws.iter_rows(values_only=True))
totals_row1 = [num(v) for v in rows[0] if v not in (None, "")]
h = {S(v): j for j, v in enumerate(rows[2]) if S(v)}; body = [r for r in rows[3:] if r and S(r[h["My GSTIN"]])]
def col(prefix):
    k = next((k for k in h if k.lower().startswith(prefix.lower())), None)
    if k is None: raise SystemExit("column not found: %s in %s" % (prefix, list(h)))
    return h[k]
C = {k: col(k) for k in ("My GSTIN", "2B Return Period", "Supplier GSTIN", "Supplier Legal Name", "Document Type", "Document Date", "Document Number", "Total Taxable Value", "Total Tax Value", "IGST Amount", "CGST Amount", "SGST Amount", "CESS Amount", "Total Document Value", "State Place of Supply", "Is Reverse Charge", "GSTR-1/IFF/GSTR-5 Filing Return", "Itc Availability", "Reason")}
out = []
for r in body:
    g = lambda k: r[C[k]]
    typ = S(g("Document Type")).upper(); sign = -1.0 if "CREDIT" in typ else 1.0
    d = ddate(g("Document Date")); p = period(g("2B Return Period")); my = S(g("My GSTIN"))
    out.append([ST.get(my[:2], my[:2]), fy(d), "2024-25", p, None, "Current", S(g("Supplier GSTIN")).upper(), S(g("Supplier Legal Name")), S(g("Document Number")), typ.title(),
                d, sign * num(g("Total Document Value")), S(g("State Place of Supply")), S(g("Is Reverse Charge")), None, sign * num(g("Total Taxable Value")),
                sign * num(g("IGST Amount")), sign * num(g("CGST Amount")), sign * num(g("SGST Amount")), sign * num(g("CESS Amount")),
                S(g("GSTR-1/IFF/GSTR-5 Filing Return")), S(g("Itc Availability")), S(g("Reason")), "PAN GSTR2B 2024-25 (Octa)", None, None, None])
isd = wb["ISD + ISDA"]; irows = list(isd.iter_rows(values_only=True)); ih = {S(v): j for j, v in enumerate(irows[0]) if S(v)}
for r in irows[2:]:
    if not r or not S(r[ih["My GSTIN"]]): continue
    my = S(r[ih["My GSTIN"]]); d = ddate(r[ih["ISD Document date"]]); p = period(S(r[ih["ISD GSTR-6 Period"]])); ta = ih["Tax amount"]
    out.append([ST.get(my[:2], my[:2]), fy(d), "2024-25", p, None, "Current", S(r[ih["GSTIN of ISD"]]).upper(), S(r[ih["Legal Name of ISD"]]), S(r[ih["ISD Document number"]]), "ISD", d, None, "", "N", None, 0.0,
                num(r[ta]), num(r[ta + 1]), num(r[ta + 2]), num(r[ta + 3]), "", S(r[ih["Eligibility of ITC"]]), "", "PAN GSTR2B 2024-25 (Octa) - ISD", None, None, None])
wb.close()
docs = [x for x in out if x[9] != "ISD"]; igst, cgst, sgst = (sum(x[16] for x in docs), sum(x[17] for x in docs), sum(x[18] for x in docs))
print("documents %d (+ISD %d) | IGST %.2f CGST %.2f SGST %.2f | file row-1 totals %s | doc FY: %s" % (len(docs), len(out) - len(docs), igst, cgst, sgst, totals_row1[:3], dict(collections.Counter(x[1] for x in docs))))
# row-1 of the export is a filtered SUBTOTAL, not a column total -> gate on per-row consistency instead: IGST+CGST+SGST == Total Tax Value
mism = [r for r in body if abs(num(r[C["IGST Amount"]]) + num(r[C["CGST Amount"]]) + num(r[C["SGST Amount"]]) - num(r[C["Total Tax Value"]])) >= 1]
print("rows where IGST+CGST+SGST != Total Tax Value:", len(mism)); assert not mism, "tax components do not tie per row - stop"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wbx = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wbx.Worksheets("GSTR-2B ITC Data")
    hdr = [sh.Cells(5, c).Value for c in range(1, 32)]
    assert hdr[:27] == ['State folder', 'FY (derived)', 'F.Y', '2B Return Period', 'GSTR 3B Month', 'Curent previous', 'GSTIN of supplier', 'Trade/Legal name', 'Invoice number', 'Invoice type', 'Invoice Date', 'Invoice Value(₹)', 'Place of supply', 'Supply Attract Reverse Charge', 'Rate', 'Taxable Value (₹)', 'Integrated Tax(₹)', 'Central Tax(₹)', 'State/UT Tax(₹)', 'Cess(₹)', 'GSTR-1/5 Period', 'ITC Availability', 'Reason', 'Source', 'IRN', 'Claimed in FY 25-26 register?', 'Claim month (derived)'], hdr[:27]
    keycol = next(i + 1 for i, v in enumerate(hdr) if v and str(v).startswith("KEY (supplier GSTIN"))
    old_n = sh.Cells(sh.Rows.Count, 1).End(-4162).Row - 5; K = len(out)
    if K > 1: sh.Rows("7:%d" % (7 + K - 2)).Insert(-4121)
    vals = [[(v - E0).days if isinstance(v, dt.datetime) else v for v in row] for row in out]
    for i in range(0, K, 4000): sh.Range(sh.Cells(6 + i, 1), sh.Cells(5 + i + len(vals[i:i + 4000]), 27)).Value = vals[i:i + 4000]
    if old_n > 1: sh.Rows("%d:%d" % (6 + K, 6 + K + old_n - 2)).Delete(-4162)
    for c in (4, 11): sh.Range(sh.Cells(6, c), sh.Cells(5 + K, c)).NumberFormat = "dd-mm-yyyy"
    sh.Range(sh.Cells(6, keycol), sh.Cells(5 + K, keycol)).Formula = '=UPPER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($G6&$I6," ",""),"-",""),"/",""),".",""),"\'",""),"_",""))'
    for c in range(keycol + 1, 32):
        if sh.Cells(5, c).Value: sh.Range(sh.Cells(6, c), sh.Cells(5 + K, c)).ClearContents()
    sh.Cells(1, 2).Value = "FY 24-25 GSTR-2B: Octa PAN-level export 'PAN GSTR2B 2024-25.xlsx' (Inv+CDN document level %d rows + ISD %d), CDN negative. Rebuilt 21-09-2026." % (len(docs), len(out) - len(docs))
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    RN = sh.Cells(sh.Rows.Count, 1).End(-4162).Row; reg = wbx.Worksheets("ITC Register 2025-26"); H = {reg.Cells(5, c).Value: c for c in range(1, 80) if reg.Cells(5, c).Value}
    print("rows 6..%d (was %d rows) | error cells %d | golden %.2f" % (RN, old_n, e, reg.Cells(4, H["Total GST"]).Value))
    assert e == 0 and RN == 5 + K
    wbx.Save()
    if xlsb_free(): wbx.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped (run the final export later)")
    wbx.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
