"""Cascade fix (Pawan 18-09): re-match ITC Register 2025-26 <-> GSTR-2B Apr25-Aug26 with
 1. leading-zero-insensitive invoice keys (SDIP/25-26/07 == SDIP/25-26/007) - KEY2 stamped as the 2B row's own KEY so SUMIFS still hit
 2. fallbacks (date+amount, amount<=3) may only take 2B documents NOT owned by an exact-number match, each once
 3. fallbacks must respect the invoice FY (register invoice FY == 2B document-date FY)
 4. fallback matches labelled "... - invoice no differs, review" in Reco Remarks (still Consider lines)
Then: re-stamp KEY2 / Countif labels / Reco Remarks / 6A1 marks, rebuild 'T6A1 Extract - 24-25' rows + PT_6A1,
re-point ITC Summary extract ranges, recalc, verify. All writes via COM (pivot survives)."""
import re, time, pickle, collections, datetime as dt, warnings
import openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
SP = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("MASTER IS OPEN IN EXCEL - close it (without saving) and rerun")
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v or 0)
    except Exception: return 0.0
def norm(s): return re.sub(r"[ \-/.'_]", "", S(s).upper())                 # = the sheets' KEY formula
def zkey(s): return re.sub(r"(?<!\d)0+(?=\d)", "", norm(s))                  # + leading zeros of digit runs dropped
def fy(d):
    if not isinstance(d, dt.datetime): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
def fy_of_label(v, d):
    s = S(v); m = re.search(r"(20)?(\d{2})-(\d{2})", s)
    return ("20%s-%s" % (m.group(2), m.group(3))) if m else fy(d)
GST = re.compile(r"^\d{2}[A-Z0-9]{13}$")
t0 = time.time()
wb = openpyxl.load_workbook(P, read_only=True)
reg = wb["ITC Register 2025-26"]; RH = {S(reg.cell(5, c).value): c for c in range(1, reg.max_column + 1)}
rows = [r for r in reg.iter_rows(min_row=6, values_only=True) if r[3]]; N = len(rows)
g = lambda r, h: r[RH[h] - 1]
b2 = wb["GSTR-2B Apr25-Aug26"]; BH = {S(b2.cell(2, c).value): c for c in range(1, b2.max_column + 1)}
B = [r for r in b2.iter_rows(min_row=3, values_only=True) if r[0]]; NB = 2 + len(B)
bg = lambda r, h: r[BH[h] - 1]
ob = wb["GSTR-2B ITC Data"]; OH = {S(ob.cell(5, c).value): c for c in range(1, ob.max_column + 1)}
o_z = collections.defaultdict(list); o_dt = collections.defaultdict(list)
for r in ob.iter_rows(min_row=6, values_only=True):
    if not r or not r[0] or S(r[OH["FY (derived)"] - 1]) != "2024-25": continue
    vg = S(r[OH["GSTIN of supplier"] - 1]).upper(); tot = round(num(r[OH["Integrated Tax(₹)"] - 1]) + num(r[OH["Central Tax(₹)"] - 1]) + num(r[OH["State/UT Tax(₹)"] - 1]), 2)
    o_z[vg + zkey(r[OH["Invoice number"] - 1])].append(1)
    d = r[OH["Invoice Date"] - 1]
    if isinstance(d, dt.datetime): o_dt[(vg, d.date(), tot)].append(1)
wb.close(); print("loaded %.0fs | register %d | 2B %d" % (time.time() - t0, N, len(B)))
# ---------------- matching (reco_lib: exact + malformed-GSTIN PAN rescue + recipient check + document-level layers + prior-year 2B)
import sys; sys.path.insert(0, r"C:\PROJECTS\gst-audit-engine")
from vel.scripts.reco_lib import match_register
reg_rows = [{"vendor_gstin": g(r, "Vendor GSTIN"), "invoice": g(r, "Invoice No."), "invoice_date": g(r, "Invoice Date"), "invoice_year": g(r, "Invoice Year"),
             "category": g(r, "Category"), "vel_gstin": g(r, "VEL GSTIN"), "igst": g(r, "IGST"), "cgst": g(r, "CGST"), "sgst": g(r, "SGST")} for r in rows]
b2_rows = [{"supplier_gstin": bg(r, "Supplier GSTIN"), "doc_no": bg(r, "Doc No"), "doc_date": bg(r, "Doc Date"), "company_gstin": bg(r, "Company GSTIN"),
            "key": S(bg(r, "Supplier GSTIN")).upper() + norm(bg(r, "Doc No")), "igst": bg(r, "IGST (Net)"), "cgst": bg(r, "CGST (Net)"), "sgst": bg(r, "SGST (Net)")} for r in B]
out = match_register(reg_rows, b2_rows, set(o_z), set(o_dt))
verdict = [o["verdict"] for o in out]; key2 = [o["key2"] for o in out]
b2key = [r["key"] for r in b2_rows]
print("verdicts:", dict(collections.Counter(v.split(" – ")[0] for v in verdict)))
# ---------------- Consider + labels
seen = set(); labels = []; six = []; sixrem = []
for i, r in enumerate(rows):
    cat = g(r, "Category"); k2 = key2[i]; v = verdict[i]
    if k2 and not k2.startswith("PY:") and k2 not in seen: seen.add(k2); labels.append("Consider")
    elif cat in ("RCM", "ISD"): labels.append("Not applicable - %s line (no 2B document)" % cat)
    elif k2.startswith("PY:"): labels.append("Not consider - matched in FY 24-25 2B (Table 6A1)")
    elif k2: labels.append("Not consider - already considered in the Consider line of this document")
    elif ("vendor GSTIN invalid" in v): labels.append("Not consider - vendor GSTIN invalid")
    elif ("Not applicable" in v): labels.append("Not consider - no vendor GSTIN")
    else: labels.append("Not in 2B (Apr-25 to Aug-26)")
    inv_fy = fy_of_label(g(r, "Invoice Year"), g(r, "Invoice Date"))
    if ("Matched with 2B of FY 24-25" in v): six.append("Yes"); sixrem.append("ITC dated 24-25 in 2B of 24-25 availed in 25-26")
    elif ("Not in 2B" in v) and inv_fy == "2024-25" and cat == "ITC": six.append("Yes"); sixrem.append("Correction Entries- ITC dated 24-25 reversed in 25-26")
    else: six.append(None); sixrem.append(None)
print("verdicts:", dict(collections.Counter(v.split(" - ")[0] for v in verdict)))
print("labels:", dict(collections.Counter(labels)))
# ---------------- T6A1 extract rows
G2ST = {"01AAECR0503Q1ZM": "Jammu & Kashmir", "03AAECR0503Q1ZI": "Punjab", "06AAECR0503Q1ZC": "Haryana", "08AAECR0503Q1Z8": "Rajasthan", "10AAECR0503Q1ZN": "Bihar", "12AAECR0503Q1ZJ": "Arunachal Pradesh", "18AAECR0503Q1Z7": "Assam", "19AAECR0503Q1Z5": "West Bengal", "20AAECR0503Q1ZM": "Jharkhand", "23AAECR0503Q1ZG": "Madhya Pradesh", "24AAECR0503Q1ZE": "Gujarat", "27AAECR0503Q1Z8": "Maharashtra", "32AAECR0503Q1ZH": "Kerala", "33AAECR0503Q1ZF": "Tamil Nadu", "36AAECR0503Q1Z9": "Telangana", "37AAECR0503Q1Z7": "Andhra Pradesh", "09AAECR0503Q1Z6": "Uttar Pradesh", "22AAECR0503Q1ZI": "Chhattisgarh", "29AAECR0503Q1Z4": "Karnataka", "27AAECR0503Q2Z7": "Maharashtra (ISD)"}
REM = L(RH["Remarks for accounting entries- For 6A1"]); B9C = L(BH["GSTR-9/9C"]); BCM = L(BH["3B Claim Month"])
ex = []
for i, r in enumerate(rows):
    if six[i] != "Yes": continue
    rr = 6 + i; cm = g(r, "3B Claim  Month")
    ex.append(["Final ITC Register 2025-26" if sixrem[i].startswith("Correction") else "2B vs 3B Reco - 2B merged - Client", g(r, "State Name"), g(r, "2B Period") or "", cm, "", cm.strftime("%b-%y") if isinstance(cm, dt.datetime) else "",
               '=IF(ISNUMBER(SEARCH("Correction",\'ITC Register 2025-26\'!$%s$%d)),"Correction Entries- ITC dated 24-25 reversed in 25-26","ITC dated 24-25 in 2B of 24-25 availed in 25-26")' % (REM, rr), "",
               g(r, "VEL GSTIN"), g(r, "Vendor GSTIN"), g(r, "Vendor Name/RCM Category"), g(r, "Invoice Date"), g(r, "Invoice No."), g(r, "Document Type"),
               num(g(r, "Taxable Value")), num(g(r, "IGST")), num(g(r, "CGST")), num(g(r, "SGST")), "", "reg!%d" % rr])
for j, r in enumerate(B):
    if fy(bg(r, "Doc Date")) != "2024-25": continue
    rr = 3 + j; gst = S(bg(r, "Company GSTIN"))
    ex.append(["2B vs 3B Reco - 2B merged - Client", G2ST.get(gst, gst), bg(r, "Tax Period"), "='GSTR-2B Apr25-Aug26'!$%s$%d" % (BCM, rr), "", "",
               '=IF(ISNUMBER(SEARCH("Unclaimed",\'GSTR-2B Apr25-Aug26\'!$%s$%d)),"ITC dated 24-25 in 2B of 25-26 - Unclaimed","ITC dated 24-25 in 2B of 25-26 availed in 25-26")' % (B9C, rr), "",
               gst, bg(r, "Supplier GSTIN"), bg(r, "Supplier Name"), bg(r, "Doc Date"), bg(r, "Doc No"), bg(r, "Doc Type"),
               num(bg(r, "Taxable Value (Net)")), num(bg(r, "IGST (Net)")), num(bg(r, "CGST (Net)")), num(bg(r, "SGST (Net)")), bg(r, "ITC Eligible"), "b2!%d" % rr])
for i, r in enumerate(rows):     # RCM Mar-25 -> Apr-25 = ITC-register RCM lines claimed 01-Apr-25 (last year's method)
    cm = g(r, "3B Claim  Month")
    if g(r, "Category") == "RCM" and isinstance(cm, dt.datetime) and (cm.month, cm.year) == (4, 2025):
        ex.append(["Final ITC Register 2025-26", g(r, "State Name"), "", cm, "", "Apr-25", "RCM paid in Mar-25 availed in Apr-25", "", g(r, "VEL GSTIN"), g(r, "Vendor GSTIN"), g(r, "Vendor Name/RCM Category"), g(r, "Invoice Date"), g(r, "Invoice No."), "RCM",
                   num(g(r, "Taxable Value")), num(g(r, "IGST")), num(g(r, "CGST")), num(g(r, "SGST")), "", "reg!%d" % (6 + i)])
print("extract rows:", len(ex), "| by source:", dict(collections.Counter(x[0] for x in ex)))
# ---------------- COM writes
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wbx = xl.Workbooks.Open(P); xl.Calculation = -4135
    rg = wbx.Worksheets("ITC Register 2025-26")
    col = lambda h: RH[h]
    rg.Range(rg.Cells(6, col("KEY2 (matched 2B key)")), rg.Cells(5 + N, col("KEY2 (matched 2B key)"))).Value = [[v] for v in key2]
    rg.Range(rg.Cells(6, col("Countif")), rg.Cells(5 + N, col("Countif"))).Value = [[v] for v in labels]
    rg.Range(rg.Cells(6, col("Reco Remarks")), rg.Cells(5 + N, col("Reco Remarks"))).Value = [[v] for v in verdict]
    rg.Range(rg.Cells(6, col("Considered in Table 6A1")), rg.Cells(5 + N, col("Considered in Table 6A1"))).Value = [[v] for v in six]
    rg.Range(rg.Cells(6, col("Remarks for accounting entries- For 6A1")), rg.Cells(5 + N, col("Remarks for accounting entries- For 6A1"))).Value = [[v] for v in sixrem]
    print("register stamped")
    ws = wbx.Worksheets("T6A1 Extract - 24-25"); ws.Activate()
    for pt in list(ws.PivotTables()): pt.TableRange2.Clear()
    old_last = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1
    ws.Range("A5:T%d" % max(old_last, 5)).ClearContents()
    for k in range(0, len(ex), 2000):
        chunk = ex[k:k + 2000]
        ws.Range("A%d:T%d" % (5 + k, 4 + k + len(chunk))).Value = [[((v - dt.datetime(1899, 12, 30)).days + (v - dt.datetime(1899, 12, 30)).seconds / 86400.0) if isinstance(v, dt.datetime) else v for v in row] for row in chunk]
    EN = 4 + len(ex)
    ws.Range("L5:L%d" % EN).NumberFormat = "dd-mm-yyyy"; ws.Range("C5:D%d" % EN).NumberFormat = "mmm-yy"; ws.Range("O5:R%d" % EN).NumberFormat = "#,##0"
    ws.Range("A4:S%d" % EN).AutoFilter(1)
    pc = wbx.PivotCaches().Create(SourceType=1, SourceData="'T6A1 Extract - 24-25'!R4C1:R%dC19" % EN)
    pt = pc.CreatePivotTable(TableDestination="'T6A1 Extract - 24-25'!R4C22", TableName="PT_6A1")
    pt.PivotFields("State").Orientation = 1; pt.PivotFields("GSTR-9 Remarks").Orientation = 2
    for f in ("IGST", "CGST", "SGST"): df = pt.AddDataField(pt.PivotFields(f), "Sum of " + f, -4157); df.NumberFormat = "#,##0"
    pt.RowAxisLayout(1); pt.TableStyle2 = "PivotStyleLight16"
    sm = wbx.Worksheets("ITC Summary")
    for c in ("G", "I", "P", "Q", "R"):
        sm.UsedRange.Replace(What="'T6A1 Extract - 24-25'!$%s$5:$%s$%d" % (c, c, old_last), Replacement="'T6A1 Extract - 24-25'!$%s$5:$%s$%d" % (c, c, EN), LookAt=2)
    print("extract rebuilt: rows 5..%d (was ..%d); pivot recreated; summary ranges repointed" % (EN, old_last))
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    hh = {rg.Cells(5, k).Value: k for k in range(1, 80) if rg.Cells(5, k).Value}
    cs = lambda h: sum(v[0] for v in rg.Range(rg.Cells(6, hh[h]), rg.Cells(5 + N, hh[h])).Value if isinstance(v[0], (int, float)))
    print("error cells:", e, "| Total GST golden %.2f | B_Total %.2f | 2B_Total %.2f | D_Total %.2f" % (rg.Cells(4, hh["Total GST"]).Value, cs("B_Total GST"), cs("2B_Total GST"), cs("D_Total GST")))
    from openpyxl.utils import column_index_from_string as CI
    print("ITC Summary total row: 6A1 blocks G/J/M/P/S =", [round(sum(sm.Cells(25, CI(c) + j).Value or 0 for j in range(3)), 2) for c in ("G", "J", "M", "P", "S")], "| Net-ITC diff", [round(sm.Cells(25, 49 + j).Value or 0, 2) for j in range(3)])
    print("ITCR vs 3B net diff: %.2f" % wbx.Worksheets("ITCR vs 3B Net ITC").Range("AG234").Value)
    wbx.Save(); wbx.Close(False)
finally: xl.Quit()
m = pickle.load(open(SP + r"\itc_b1_meta.pkl", "rb")); m.update({"verdict": verdict, "key2": key2, "consider": ["Consider" if l == "Consider" else "NA" for l in labels], "consider_labels": labels}); pickle.dump(m, open(SP + r"\itc_b1_meta.pkl", "wb"))
print("done %.0fs" % (time.time() - t0))
