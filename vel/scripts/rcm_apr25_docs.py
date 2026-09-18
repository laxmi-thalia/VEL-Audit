"""ITC Register 2025-26: replace the 49 Apr-25 RCM category lines (FY 24-25 RCM paid in Mar-25, availed Apr-25 = Table 6A1
component 5) with the Mar-25 DOCUMENTS from last year's RCM Register (VEL_GSTR 9_9C FY 24-25.xlsb, 3B Month '12 Mar 2025').
Pawan 18-09: supplier GSTIN / document / invoice must be real. Grouping and column conventions as rcm_itc_doclevel.py.
Delete-all-first (asserted), then insert per state at re-scanned positions, verify totals per state before saving."""
import shutil, time, datetime as dt, collections, warnings, openpyxl, win32com.client as win32, pythoncom
from pyxlsb import open_workbook, convert_date
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("MASTER IS OPEN IN EXCEL - close it (without saving) and rerun")
shutil.copy(P, "master2_snapshot_before_apr25rcm.xlsx")
E0 = dt.datetime(1899, 12, 30)
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v or 0)
    except Exception: return 0.0
def xd(v): return convert_date(v) if isinstance(v, float) and v > 30000 else (v if isinstance(v, dt.datetime) else None)
def ser(d): return (d - E0).days if isinstance(d, dt.datetime) else None
def fy(d):
    if not isinstance(d, dt.datetime): return None
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
GLN = {"2610080300": "CGST Output RCM", "2610080301": "SGST Output RCM", "2610080302": "IGST Output RCM"}
with open_workbook("VEL_2425.xlsb") as w:
    with w.get_sheet("RCM Register") as sh:
        raw = []
        for i, r in enumerate(sh.rows()):
            v = [c.v for c in r]
            if i > 8 and not any(x not in (None, "") for x in v[:6]): break
            raw.append(v)
h = {S(x): j for j, x in enumerate(raw[3]) if S(x)}; body = raw[4:]
mar = [r for r in body if S(r[h["3B Month"]]).startswith("12 Mar 2025")]
docs = collections.OrderedDict()
for r in mar:
    pdt = xd(r[h["Posting Date"]]); dn = r[h["Document Number"]]; dn = int(dn) if isinstance(dn, float) else dn
    key = ("%s|%s" % (pdt.strftime("%Y%m%d") if pdt else "", dn), S(r[h["Vendor Name"]]), S(r[h["Nature of Services"]]), S(r[h["Reference"]]))
    d = docs.get(key)
    if d is None:
        d = docs[key] = {"gstin": S(r[h["MY GSTN"]]), "bp": S(r[h["Business place"]]), "doctype": S(r[h["Document type"]]), "docno": dn, "pdt": pdt, "ref": S(r[h["Reference"]]),
                         "ddt": xd(r[h["Document Date"]]), "vcode": r[h["Vendor Code"]], "vgstin": S(r[h["GSTN"]]), "vname": S(r[h["Vendor Name"]]), "cat": S(r[h["Nature of Services"]]),
                         "rate": r[h["GST Rate"]], "pc": r[h["Profit Center"]], "tv": 0.0, "ig": 0.0, "cg": 0.0, "sg": 0.0}
    d["tv"] += num(r[h["Taxable Value as per SAP"]]); d["ig"] += num(r[h["IGST AS PER SAP"]]); d["cg"] += num(r[h["CGST AS PER SAP"]]); d["sg"] += num(r[h["SGST AS PER SAP"]])
print("LY Mar-25 legs:", len(mar), "| documents:", len(docs), "| tax %.2f | GSTIN on %d docs" % (sum(d["ig"] + d["cg"] + d["sg"] for d in docs.values()), sum(1 for d in docs.values() if d["vgstin"])))
bystate = collections.defaultdict(list)
for d in docs.values(): bystate[d["gstin"]].append(d)
exp = {g: round(sum(d["ig"] + d["cg"] + d["sg"] for d in ds), 2) for g, ds in bystate.items()}
wb = openpyxl.load_workbook(P, read_only=True); ws = wb["ITC Register 2025-26"]; hdr = [c.value for c in next(ws.iter_rows(min_row=5, max_row=5))]; H = {h_: i + 1 for i, h_ in enumerate(hdr) if h_}
itc = [(i + 6, r) for i, r in enumerate(ws.iter_rows(min_row=6, values_only=True)) if r[3]]; wb.close()
g2state = {}; g2bp = {}
for _, r in itc: g2state.setdefault(r[H["VEL GSTIN"] - 1], r[H["State Name"] - 1]); g2bp.setdefault(r[H["VEL GSTIN"] - 1], r[H["Business place"] - 1])
def itc_row(d):
    igst_doc = d["ig"] > 0 and d["cg"] == 0
    return {"Company": "VEL", "Business place": d["bp"] or g2bp.get(d["gstin"]), "State Name": g2state.get(d["gstin"], ""), "VEL GSTIN": d["gstin"], "3B Claim  Month": ser(dt.datetime(2025, 4, 1)),
            "Category": "RCM", "Document Type": d["doctype"] or "NA", "Document Number": d["docno"], "Posting Date": ser(d["pdt"]), "Posting Year": fy(d["pdt"]),
            "Invoice No.": d["ref"] or "NA", "Invoice Date": ser(d["ddt"]), "Invoice Year": fy(d["ddt"]), "Vendor Code": d["vcode"] if d["vcode"] not in (None, "") else "NA",
            "Vendor GSTIN": d["vgstin"] or "Missing", "Vendor Name/RCM Category": d["vname"] or d["cat"], "Tax Rate": d["rate"], "Taxable Value": round(d["tv"], 2),
            "IGST": round(d["ig"], 2), "CGST": round(d["cg"], 2), "SGST": round(d["sg"], 2), "Countif": None, "KEY2 (matched 2B key)": None,
            "Reco Remarks": "RCM self-invoice (FY 24-25 document, paid Mar-25, availed Apr-25) - Input GL dump starts 01-04-2025, not checkable",
            "Nature of Services": d["cat"], "G/L Account": 2610080302 if igst_doc else 2610080300, "G/L Account Text": "IGST Receivable" if igst_doc else "C-SGST Receivable",
            "SAP Period": ser(dt.datetime(d["pdt"].year, d["pdt"].month, 1)) if d["pdt"] else None, "Profit Center": d["pc"], "TYPE": "NA", "F.Y/Booking Year": "2024-25", "Company Code": 1000}
dels = [n for n, r in itc if r[H["Category"] - 1] == "RCM" and r[H["Document Number"] - 1] == "RCM"]
apr_cat = [(n, r) for n, r in itc if n in set(dels)]
old = collections.Counter()
for n, r in apr_cat: old[r[H["VEL GSTIN"] - 1]] += round(num(r[H["IGST"] - 1]) + num(r[H["CGST"] - 1]) + num(r[H["SGST"] - 1]), 2)
print("category lines to delete:", len(dels), "| old tax by state vs LY docs:", {g: (round(old.get(g, 0)), round(exp.get(g, 0))) for g in sorted(set(old) | set(exp)) if abs(old.get(g, 0) - exp.get(g, 0)) > 1})
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wbx = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wbx.Worksheets("ITC Register 2025-26")
    NC = max(H.values()); fcols = [c for c in range(1, NC + 1) if sh.Cells(6, c).HasFormula]; vcols = {h_: c for h_, c in H.items() if c not in fcols}
    fmt = {h_: sh.Cells(6, H[h_]).NumberFormat for h_ in ("3B Claim  Month", "Posting Date", "Invoice Date", "SAP Period")}
    RN0 = itc[-1][0]; cat = [v[0] for v in sh.Range(sh.Cells(6, H["Category"]), sh.Cells(RN0, H["Category"])).Value]; dno = [v[0] for v in sh.Range(sh.Cells(6, H["Document Number"]), sh.Cells(RN0, H["Document Number"])).Value]
    for n in dels: assert cat[n - 6] == "RCM" and dno[n - 6] == "RCM", ("delete target is not a category RCM line", n)
    runs = []
    for n in sorted(dels, reverse=True):
        if runs and runs[-1][0] == n + 1: runs[-1][0] = n
        else: runs.append([n, n])
    for a, b in runs: sh.Rows("%d:%d" % (a, b)).Delete(-4162)
    RN1 = RN0 - len(dels); print("deleted %d category lines (%d runs)" % (len(dels), len(runs)), flush=True)
    gst = [v[0] for v in sh.Range(sh.Cells(6, H["VEL GSTIN"]), sh.Cells(RN1, H["VEL GSTIN"])).Value]
    pos = {}
    for g in bystate:
        g_rows = [6 + i for i in range(len(gst)) if gst[i] == g]
        if g_rows: pos[g] = g_rows[0]                      # Apr-25 claims go first in the state's block (claim order)
        else: print("  WARNING: %s has %d LY RCM documents but no rows in the ITC register - not inserted" % (g, len(bystate[g])))
    for g, at in sorted(pos.items(), key=lambda kv: -kv[1]):
        new = [itc_row(d) for d in sorted(bystate[g], key=lambda d: (d["pdt"] or dt.datetime(1900, 1, 1), str(d["docno"])))]
        K = len(new); sh.Rows("%d:%d" % (at, at + K - 1)).Insert(-4121); src_row = at + K
        for c in fcols: sh.Range(sh.Cells(at, c), sh.Cells(at + K - 1, c)).FormulaR1C1 = sh.Cells(src_row, c).FormulaR1C1
        for h_, c in vcols.items(): sh.Range(sh.Cells(at, c), sh.Cells(at + K - 1, c)).Value = [[r.get(h_)] for r in new]
        for h_, f in fmt.items(): sh.Range(sh.Cells(at, H[h_]), sh.Cells(at + K - 1, H[h_])).NumberFormat = f
        print("  %s: inserted %d Apr-25 RCM documents at row %d" % (g, K, at), flush=True)
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    RN = sh.Cells(sh.Rows.Count, 4).End(-4162).Row; sh.Range(sh.Cells(5, 1), sh.Cells(RN, NC)).AutoFilter()
    cat = [v[0] for v in sh.Range(sh.Cells(6, H["Category"]), sh.Cells(RN, H["Category"])).Value]; dno = [v[0] for v in sh.Range(sh.Cells(6, H["Document Number"]), sh.Cells(RN, H["Document Number"])).Value]
    cmv = [v[0] for v in sh.Range(sh.Cells(6, H["3B Claim  Month"]), sh.Cells(RN, H["3B Claim  Month"])).Value2]; tg = [v[0] for v in sh.Range(sh.Cells(6, H["Total GST"]), sh.Cells(RN, H["Total GST"])).Value]; gs = [v[0] for v in sh.Range(sh.Cells(6, H["VEL GSTIN"]), sh.Cells(RN, H["VEL GSTIN"])).Value]
    got = collections.Counter(); left = 0
    for i in range(len(cat)):
        if cat[i] == "RCM" and dno[i] == "RCM": left += 1
        if cat[i] == "RCM" and isinstance(cmv[i], (int, float)) and (E0 + dt.timedelta(days=int(cmv[i]))).strftime("%Y-%m") == "2025-04": got[gs[i]] += round(float(tg[i] or 0), 2)
    bad = [(g, round(got[g], 2), exp.get(g)) for g in set(got) | set(exp) if abs(got[g] - exp.get(g, 0)) > 1]
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("error cells:", e, "| rows 6..%d | category lines left: %d | Apr-25 RCM tax by state mismatches vs LY docs: %s" % (RN, left, bad))
    print("Total GST golden now %.2f | Net-ITC diff %s" % (sh.Cells(4, H["Total GST"]).Value, [round(wbx.Worksheets("ITC Summary").Cells(25, 49 + j).Value or 0, 2) for j in range(3)]))
    assert e == 0 and left == 0 and not bad, "VERIFICATION FAILED - not saved"
    wbx.Save(); wbx.Close(False)
finally: xl.Quit()
print("done %.0fs" % (time.time() - t0))
