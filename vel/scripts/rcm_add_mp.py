"""Madhya Pradesh RCM rows from the monthly workings ('GSTR 3B Madhyapradesh <Mon> <Year>.xlsx' - space, not hyphen, which is why
the earlier passes missed MP). Pawan 18-09 (point 3: MP July path; point 4: months with a Statewise/Month-wise difference take the
monthly working, not Conso):
  * Jul-25: not in Conso at all -> INSERT the 48 monthly rows (posting grain: 300-301 combined -> CGST + SGST rows, taxable 50/50).
  * Jan-26: Conso 9,59,240 / 1,54,471 vs 3B 8,47,240 / 1,48,871; monthly working ties to 3B -> REPLACE the 37 Conso rows.
All via COM so every dependent range expands/shrinks with the insert/delete; new rows take formulas (R1C1) from the row above."""
import pickle, shutil, time, datetime as dt, collections, win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("MASTER IS OPEN IN EXCEL - close it (without saving) and rerun")
shutil.copy(P, "master2_snapshot_before_mp.xlsx")
D = pickle.load(open("rcm_monthly.pkl", "rb"))["data"]; G = "23AAECR0503Q1ZG"
GLN = {"2610080300": "CGST Output RCM", "2610080301": "SGST Output RCM", "2610080302": "IGST Output RCM"}
E0 = dt.datetime(1899, 12, 30)
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v)
    except Exception: return 0.0
def ser(d): return (d - E0).days if isinstance(d, dt.datetime) else None
def fy(d):
    if not isinstance(d, dt.datetime): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
def build(month, b3month, claim, source):
    rec = D[(G, month)]; hi = {}
    for i, h in enumerate(rec["header"]):
        if h and h not in hi: hi[h] = i
    out = []
    for r in rec["rows"]:
        d = lambda h: r[hi[h]] if h in hi and hi[h] < len(r) else None
        gl = S(d("G/L Account")); tv = num(d("Taxable Value")); ig, cg, sg = num(d("IGST")), num(d("CGST")), num(d("SGST"))
        base = {"3B Month": ser(b3month), "Business place": S(d("Business place")) or "MP01", "MY GSTN": G, "State": "Madhya Pradesh",
                "Fiscal Year": S(d("Fiscal Year")), "Year/Month": S(d("Year/Month") or d("Year/ Month")), "Document type": S(d("Document type") or d("Document Type")),
                "Document Number": d("Document Number"), "Posting Date": ser(d("Posting Date")), "Posting Year": fy(d("Posting Date")),
                "Assignment": S(d("Assignment")), "Reference": S(d("Reference")), "Document Date": ser(d("Document Date")), "Doc year": fy(d("Document Date")),
                "Vendor Code": d("Vendor Code"), "GSTN": S(d("Vendor GSTN")), "Vendor Name": S(d("Vendor Name")), "Posting Key": S(d("Posting Key")),
                "Local Currency": S(d("Local Currency")) or "INR", "Tax Code": S(d("Tax Code")), "Clearing Document": d("Clearing Document"),
                "Profit Center": d("Profit Center"), "Text": S(d("Text")), "Offsetting Account": d("Offsetting Account"),
                "GST Rate": num(d("Rate")) if S(d("Rate")) else (round((ig + cg + sg) / tv * 100, 2) if tv else None),
                "Nature of Services": S(d("Category")), "GSTR 3B Claim month": claim, "Source": source}
        if gl.startswith("2610080302") or (ig and not cg):
            row = dict(base); row.update({"G/L Account": 2610080302, "GL Name": GLN["2610080302"], "Amount in Local Currency": -ig, "Taxable Value as per SAP": tv, "IGST AS PER SAP": ig, "CGST AS PER SAP": 0.0, "SGST AS PER SAP": 0.0}); out.append(row)
        else:
            for acct, amt, tag in (("2610080300", cg, "C"), ("2610080301", sg, "S")):
                row = dict(base); row.update({"G/L Account": int(acct), "GL Name": GLN[acct], "Amount in Local Currency": -amt, "Taxable Value as per SAP": tv / 2, "IGST AS PER SAP": 0.0,
                                              "CGST AS PER SAP": cg if tag == "C" else 0.0, "SGST AS PER SAP": sg if tag == "S" else 0.0}); out.append(row)
    return out
JUL = build("Jul-25", dt.datetime(2025, 7, 4), "05 Aug 2025", "Monthly working Jul-25 (RCM sheet) - not in Conso RCM")
JAN = build("Jan-26", dt.datetime(2026, 1, 10), "11 Feb 2026", "Monthly working Jan-26 (RCM sheet) - replaces Conso RCM per Pawan 17-09")
for nm, rows in (("Jul-25", JUL), ("Jan-26", JAN)):
    print("%s: %d posting rows | taxable %.2f IGST %.2f CGST %.2f SGST %.2f" % (nm, len(rows), sum(r["Taxable Value as per SAP"] for r in rows), sum(r["IGST AS PER SAP"] for r in rows), sum(r["CGST AS PER SAP"] for r in rows), sum(r["SGST AS PER SAP"] for r in rows)))
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; rg = wb.Worksheets("RCM Register")
    RH = {rg.Cells(5, c).Value: c for c in range(1, 100) if rg.Cells(5, c).Value}; NC = max(RH.values())
    RN = max(r for r in range(6, rg.UsedRange.Rows.Count + 6) if rg.Cells(r, RH["MY GSTN"]).Value or rg.Cells(r, RH["Business place"]).Value)
    fcols = [c for c in range(1, NC + 1) if rg.Cells(6, c).HasFormula]; vcols = {h: c for h, c in RH.items() if c not in fcols}
    datecols = [RH[h] for h in ("3B Month", "Posting Date", "Document Date")]
    col_bp, col_m, col_src = RH["Business place"], RH["3B Month"], RH["Source"]
    bp = [v[0] for v in rg.Range(rg.Cells(6, col_bp), rg.Cells(RN, col_bp)).Value]; mo = [v[0] for v in rg.Range(rg.Cells(6, col_m), rg.Cells(RN, col_m)).Value2]; src = [v[0] for v in rg.Range(rg.Cells(6, col_src), rg.Cells(RN, col_src)).Value]
    def put(at, rows):
        K = len(rows); rg.Rows("%d:%d" % (at, at + K - 1)).Insert(-4121)
        for c in fcols: rg.Range(rg.Cells(at, c), rg.Cells(at + K - 1, c)).FormulaR1C1 = rg.Cells(at - 1, c).FormulaR1C1
        for h, c in vcols.items(): rg.Range(rg.Cells(at, c), rg.Cells(at + K - 1, c)).Value = [[r.get(h)] for r in rows]
        for c in datecols: rg.Range(rg.Cells(at, c), rg.Cells(at + K - 1, c)).NumberFormat = "dd-mm-yyyy"
        return K
    # Jan-26 first (higher rows): delete the MP Conso Jan-26 block, insert the monthly rows in its place
    jan_rows = [6 + i for i in range(len(bp)) if bp[i] == "MP01" and isinstance(mo[i], (int, float)) and (E0 + dt.timedelta(days=mo[i])).month == 1 and str(src[i]).startswith("Conso")]
    assert jan_rows and jan_rows[-1] - jan_rows[0] + 1 == len(jan_rows), ("MP Jan-26 Conso block not contiguous", jan_rows[:3], len(jan_rows))
    rg.Rows("%d:%d" % (jan_rows[0], jan_rows[-1])).Delete(-4162); print("deleted MP Jan-26 Conso rows:", len(jan_rows), "at", jan_rows[0])
    kj = put(jan_rows[0], JAN); print("inserted MP Jan-26 monthly rows:", kj, "at", jan_rows[0])
    # Jul-25: after the last Jul-25 row
    last_jul = max(6 + i for i in range(len(bp)) if isinstance(mo[i], (int, float)) and (E0 + dt.timedelta(days=mo[i])).month == 7)
    kj2 = put(last_jul + 1, JUL); print("inserted MP Jul-25 rows:", kj2, "at", last_jul + 1)
    RN2 = RN - len(jan_rows) + kj + kj2; rg.Range(rg.Cells(5, 1), rg.Cells(RN2, NC)).AutoFilter()
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    sub4 = rg.Cells(4, RH["Taxable Value as per SAP"]).Formula
    mw = wb.Worksheets("Month wise RCM vs 3B"); out = []
    for r in range(5, 246):
        if mw.Cells(r, 1).Value and "Madhya" in str(mw.Cells(r, 1).Value) and str(mw.Cells(r, 3).Value) in ("Jul-25", "Jan-26"):
            out.append((mw.Cells(r, 3).Value, round(mw.Cells(r, 4).Value or 0, 2), round(mw.Cells(r, 9).Value or 0, 2), round((mw.Cells(r, 8).Value or 0) - (mw.Cells(r, 13).Value or 0), 2)))
    fo = collections.Counter(str(v[0]) for v in rg.Range(rg.Cells(last_jul + 1, RH["Found in Output GL"]), rg.Cells(last_jul + kj2, RH["Found in Output GL"])).Value)
    print("error cells:", e, "| register rows 6..%d | row-4 subtotal: %s" % (RN2, sub4))
    print("Month wise MP (month, reg taxable, 3B taxable, tax diff):", out)
    print("Statewise RCM vs 3B total diff: %.2f | MP July rows Found in Output GL: %s" % (wb.Worksheets("Statewise RCM vs 3B").Range("N26").Value, dict(fo)))
    print("ITC golden %.2f" % wb.Worksheets("ITC Register 2025-26").Cells(4, 23).Value if wb.Worksheets("ITC Register 2025-26").Cells(4, 23).Value else "ITC golden n/a")
    wb.Save(); wb.Close(False)
finally: xl.Quit()
print("done %.0fs" % (time.time() - t0))
