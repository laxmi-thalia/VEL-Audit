"""ITC Register 2025-26: replace the category-level RCM lines (state x claim month x category, Document Number 'RCM') with
DOCUMENT-level lines sourced from the RCM Register (Pawan 18-09). FY 25-26 claims only (May-25..Mar-26; the Apr-25 claim lines
are FY 24-25 RCM = Table 6A1 component and stay as they are; Apr-26 claims belong to FY 26-27). One ITC line per RCM document
(GL Key) x vendor x category: legs summed (Conso single line carries CGST+SGST; monthly rows carry two legs at taxable/2).
Input-GL confirmation (same SAP document, RCM GL side 'Input') carried in Reco Remarks. COM: delete + insert per state block so
every dependent range follows; formulas copied R1C1 from the row above."""
import shutil, time, datetime as dt, collections, warnings, openpyxl, win32com.client as win32, pythoncom
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("MASTER IS OPEN IN EXCEL - close it (without saving) and rerun")
shutil.copy(P, "master2_snapshot_before_rcmdoc.xlsx")
E0 = dt.datetime(1899, 12, 30)
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v or 0)
    except Exception: return 0.0
def ser(d): return (d - E0).days if isinstance(d, dt.datetime) else None
def fy(d):
    if not isinstance(d, dt.datetime): return None
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
FYIDX = {4: 1, 5: 2, 6: 3, 7: 4, 8: 5, 9: 6, 10: 7, 11: 8, 12: 9, 1: 10, 2: 11, 3: 12}
def claim(v):
    """register 'GSTR 3B Claim month' -> (y, m) or None (Apr-26 = FY 26-27 claim excluded)"""
    if isinstance(v, dt.datetime): y, m = v.year, v.month
    else:
        s = S(v)
        try: d = dt.datetime.strptime(s, "%d %b %Y"); y, m = d.year, d.month
        except Exception:
            try: d = dt.datetime.strptime(s, "%d %B %Y"); y, m = d.year, d.month
            except Exception: return None
    return (y, m) if (y, m) >= (2025, 5) and (y, m) <= (2026, 3) else None
# ---- read register (posting grain) + ITC register conventions
wb = openpyxl.load_workbook(P, read_only=True)
rg = wb["RCM Register"]; rh = [c.value for c in next(rg.iter_rows(min_row=5, max_row=5))]; RH = {h: i for i, h in enumerate(rh) if h}
reg = [r for r in rg.iter_rows(min_row=6, values_only=True) if r[RH["Business place"]]]
ws = wb["ITC Register 2025-26"]; hdr = [c.value for c in next(ws.iter_rows(min_row=5, max_row=5))]; H = {h: i + 1 for i, h in enumerate(hdr) if h}
itc = [(i + 6, r) for i, r in enumerate(ws.iter_rows(min_row=6, values_only=True)) if r[3]]; wb.close()
g2state = {}; g2bp = {}
for _, r in itc: g2state.setdefault(r[H["VEL GSTIN"] - 1], r[H["State Name"] - 1]); g2bp.setdefault(r[H["VEL GSTIN"] - 1], r[H["Business place"] - 1])
# ---- Input GL flag per GL Key (RCM GL sheet, Side = Input)  - read once
wb2 = openpyxl.load_workbook(P, read_only=True, data_only=True); gl = wb2["RCM GL"]; gh = [c.value for c in next(gl.iter_rows(min_row=5, max_row=5))]; GH = {h: i for i, h in enumerate(gh) if h}
inkeys = {r[GH["GL Key"]] for r in gl.iter_rows(min_row=6, values_only=True) if r[GH["Side"]] == "Input"}
wb2.close()
# ---- group register legs into documents
docs = collections.OrderedDict()
skipped = collections.Counter()
for r in reg:
    cmo = claim(r[RH["GSTR 3B Claim month"]])
    if not cmo: skipped[S(r[RH["GSTR 3B Claim month"]])[:14] or "blank"] += 1; continue
    g = r[RH["MY GSTN"]]
    if not g: skipped["no GSTIN (HOIS)"] += 1; continue
    pdt = r[RH["Posting Date"]]; key = ("%s|%s" % (pdt.strftime("%Y%m%d") if isinstance(pdt, dt.datetime) else "", r[RH["Document Number"]]), S(r[RH["Vendor Name"]]), S(r[RH["Nature of Services"]]), S(r[RH["Reference"]]))
    d = docs.get(key)
    if d is None:
        d = docs[key] = {"gstin": g, "bp": r[RH["Business place"]], "claim": cmo, "doctype": S(r[RH["Document type"]]), "docno": r[RH["Document Number"]], "pdt": pdt, "ref": S(r[RH["Reference"]]),
                         "ddt": r[RH["Document Date"]], "vcode": r[RH["Vendor Code"]], "vgstin": S(r[RH["GSTN"]]), "vname": S(r[RH["Vendor Name"]]), "cat": S(r[RH["Nature of Services"]]), "rate": r[RH["GST Rate"]],
                         "pc": r[RH["Profit Center"]], "glkey": key[0], "tv": 0.0, "ig": 0.0, "cg": 0.0, "sg": 0.0}
    d["tv"] += num(r[RH["Taxable Value as per SAP"]]); d["ig"] += num(r[RH["IGST AS PER SAP"]]); d["cg"] += num(r[RH["CGST AS PER SAP"]]); d["sg"] += num(r[RH["SGST AS PER SAP"]])
print("register legs:", len(reg), "| documents (FY 25-26 claims):", len(docs), "| skipped legs:", dict(skipped))
bystate = collections.defaultdict(list)
for d in docs.values(): bystate[d["gstin"]].append(d)
def itc_row(d):
    y, m = d["claim"]; igst_doc = d["ig"] > 0 and d["cg"] == 0
    return {"Company": "VEL", "Business place": d["bp"], "State Name": g2state.get(d["gstin"], ""), "VEL GSTIN": d["gstin"], "3B Claim  Month": ser(dt.datetime(y, m, FYIDX[m])),
            "Category": "RCM", "Document Type": d["doctype"] or "NA", "Document Number": d["docno"], "Posting Date": ser(d["pdt"]), "Posting Year": fy(d["pdt"]),
            "Invoice No.": d["ref"] or "NA", "Invoice Date": ser(d["ddt"]), "Invoice Year": fy(d["ddt"]), "Vendor Code": d["vcode"] if d["vcode"] not in (None, "") else "NA",
            "Vendor GSTIN": d["vgstin"] or "Missing", "Vendor Name/RCM Category": d["vname"] or d["cat"], "Tax Rate": d["rate"], "Taxable Value": round(d["tv"], 2),
            "IGST": round(d["ig"], 2), "CGST": round(d["cg"], 2), "SGST": round(d["sg"], 2), "Countif": None, "KEY2 (matched 2B key)": None,
            "Reco Remarks": "RCM self-invoice - Input GL %s debit %s in the same SAP document" % ("1910060502" if igst_doc else "1910060500/501", "found" if d["glkey"] in inkeys else "NOT found"),
            "Nature of Services": d["cat"], "G/L Account": 2610080302 if igst_doc else 2610080300, "G/L Account Text": "IGST Receivable" if igst_doc else "C-SGST Receivable",
            "SAP Period": ser(dt.datetime(d["pdt"].year, d["pdt"].month, 1)) if isinstance(d["pdt"], dt.datetime) else None, "Profit Center": d["pc"], "TYPE": "NA",
            "F.Y/Booking Year": "2025-26", "Company Code": 1000}
# ---- existing category lines: delete ALL non-Apr-25 RCM lines first (exact rows, verified), then insert per state at re-scanned positions
dels = []; kept = 0
for n, r in itc:
    if r[H["Category"] - 1] != "RCM": continue
    cmv = r[H["3B Claim  Month"] - 1]
    if isinstance(cmv, dt.datetime) and (cmv.year, cmv.month) == (2025, 4): kept += 1
    else: dels.append(n)
tot_new = sum(len(v) for v in bystate.values())
print("category lines to delete:", len(dels), "| Apr-25 lines kept:", kept, "| new document lines:", tot_new)
exp_by_month = collections.Counter()
for d in docs.values(): exp_by_month[d["claim"]] += round(d["ig"] + d["cg"] + d["sg"], 2)
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wbx = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wbx.Worksheets("ITC Register 2025-26")
    NC = max(H.values()); fcols = [c for c in range(1, NC + 1) if sh.Cells(6, c).HasFormula]; vcols = {h: c for h, c in H.items() if c not in fcols}
    fmt = {h: sh.Cells(6, H[h]).NumberFormat for h in ("3B Claim  Month", "Posting Date", "Invoice Date", "SAP Period")}
    RN0 = itc[-1][0]
    cat = [v[0] for v in sh.Range(sh.Cells(6, H["Category"]), sh.Cells(RN0, H["Category"])).Value]; dno = [v[0] for v in sh.Range(sh.Cells(6, H["Document Number"]), sh.Cells(RN0, H["Document Number"])).Value]
    for n in dels: assert cat[n - 6] == "RCM" and dno[n - 6] == "RCM", ("delete target is not a category RCM line", n, cat[n - 6], dno[n - 6])
    runs = []
    for n in sorted(dels, reverse=True):
        if runs and runs[-1][0] == n + 1: runs[-1][0] = n
        else: runs.append([n, n])
    for a, b in runs: sh.Rows("%d:%d" % (a, b)).Delete(-4162)
    RN1 = RN0 - len(dels); print("deleted %d category lines in %d runs; rows now 6..%d" % (len(dels), len(runs), RN1), flush=True)
    gst = [v[0] for v in sh.Range(sh.Cells(6, H["VEL GSTIN"]), sh.Cells(RN1, H["VEL GSTIN"])).Value]; cat = [v[0] for v in sh.Range(sh.Cells(6, H["Category"]), sh.Cells(RN1, H["Category"])).Value]
    assert sum(1 for c in cat if c == "RCM") == kept, ("kept count mismatch", sum(1 for c in cat if c == "RCM"), kept)
    pos = {}
    for g in bystate:
        rcm_rows = [6 + i for i in range(len(gst)) if gst[i] == g and cat[i] == "RCM"]; g_rows = [6 + i for i in range(len(gst)) if gst[i] == g]
        if rcm_rows: pos[g] = rcm_rows[-1] + 1
        elif g_rows: pos[g] = g_rows[0]
        else: print("  WARNING: %s has %d RCM documents but no rows in the ITC register - not inserted" % (g, len(bystate[g])))
    for g, at in sorted(pos.items(), key=lambda kv: -kv[1]):     # bottom-up so earlier positions stay valid
        new = [itc_row(d) for d in sorted(bystate[g], key=lambda d: (d["claim"], d["pdt"] or dt.datetime(1900, 1, 1), str(d["docno"])))]
        K = len(new); sh.Rows("%d:%d" % (at, at + K - 1)).Insert(-4121)
        src_row = at - 1 if at > 6 else at + K   # formulas from an existing row (above, else the first row below the insert)
        for c in fcols: sh.Range(sh.Cells(at, c), sh.Cells(at + K - 1, c)).FormulaR1C1 = sh.Cells(src_row, c).FormulaR1C1
        for h, c in vcols.items(): sh.Range(sh.Cells(at, c), sh.Cells(at + K - 1, c)).Value = [[r.get(h)] for r in new]
        for h, f in fmt.items(): sh.Range(sh.Cells(at, H[h]), sh.Cells(at + K - 1, H[h])).NumberFormat = f
        print("  %s: inserted %d document lines at row %d" % (g, K, at), flush=True)
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    RN = sh.Cells(sh.Rows.Count, 4).End(-4162).Row; sh.Range(sh.Cells(5, 1), sh.Cells(RN, NC)).AutoFilter()
    # verification: per claim month, RCM tax on the sheet == register legs (documents); no category lines left except Apr-25
    cat = [v[0] for v in sh.Range(sh.Cells(6, H["Category"]), sh.Cells(RN, H["Category"])).Value]; dno = [v[0] for v in sh.Range(sh.Cells(6, H["Document Number"]), sh.Cells(RN, H["Document Number"])).Value]
    cmv = [v[0] for v in sh.Range(sh.Cells(6, H["3B Claim  Month"]), sh.Cells(RN, H["3B Claim  Month"])).Value2]; tg = [v[0] for v in sh.Range(sh.Cells(6, H["Total GST"]), sh.Cells(RN, H["Total GST"])).Value]
    got = collections.Counter(); left = 0
    for i in range(len(cat)):
        if cat[i] != "RCM": continue
        d = E0 + dt.timedelta(days=int(cmv[i])); 
        if dno[i] == "RCM": left += 1
        else: got[(d.year, d.month)] += round(float(tg[i] or 0), 2)
    bad = [(k, round(got[k], 2), round(exp_by_month[k], 2)) for k in sorted(set(got) | set(exp_by_month)) if abs(got[k] - exp_by_month[k]) > 1]
    e = 0
    for w in wbx.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("error cells:", e, "| rows 6..%d | category lines left: %d (expected %d Apr-25) | claim-month mismatches vs register: %s" % (RN, left, kept, bad))
    print("Total GST golden now %.2f | ITC Summary Net-ITC diff %s | ITCR vs 3B net diff %.2f" % (sh.Cells(4, H["Total GST"]).Value, [round(wbx.Worksheets("ITC Summary").Cells(25, 49 + j).Value or 0, 2) for j in range(3)], wbx.Worksheets("ITCR vs 3B Net ITC").Range("AG234").Value))
    assert e == 0 and left == kept and not bad, "VERIFICATION FAILED - not saved"
    wbx.Save(); wbx.Close(False)
finally: xl.Quit()
print("done %.0fs" % (time.time() - t0))
