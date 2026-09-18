"""Read-only survey: FY 26-27 state working files (Audit data of FY 2026-27 / <month> / <state> / GSTR-3B*.xlsx) -> 'Inputs' sheet
(header row = the row holding 'Document Number'). Collect all rows; report per month/state counts, Invoice Year '25-26' rows
(= FY 25-26 dated invoices claimed in FY 26-27) and compare with the current 'ITC Register 2026-27' (260 rows from 2B)."""
import os, re, pickle, collections, warnings, openpyxl, datetime as dt
warnings.filterwarnings("ignore")
A = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Audit data of FY 2026-27/"
MONTHS = ["01 Apr 2026", "02 May 2026", "03 June 2026", "04 July 2026"]
def S(v): return "" if v is None else str(v).strip()
out = []; miss = []; hdrs = collections.Counter()
for mf in MONTHS:
    for d in sorted(os.listdir(A + mf)):
        folder = A + mf + "/" + d + "/"
        if not os.path.isdir(folder): continue
        c = [x for x in os.listdir(folder) if re.match(r"gstr[ -]?3b", x.lower()) and x.lower().endswith(".xlsx") and not x.startswith("~$")]
        if not c: miss.append((mf, d, "no 3B xlsx")); continue
        try: wb = openpyxl.load_workbook(folder + c[0], read_only=True, data_only=True)
        except Exception as e: miss.append((mf, d, "open failed")); continue
        sn = next((s for s in wb.sheetnames if s.strip().lower() in ("inputs", "input")), None)
        if not sn: miss.append((mf, d, "no Inputs sheet: %s" % wb.sheetnames[:8])); wb.close(); continue
        ws = wb[sn]; rows = list(ws.iter_rows(values_only=True)); hi = next((i for i, r in enumerate(rows[:40]) if r and any(S(v) == "Document Number" for v in r)), None)
        if hi is None: miss.append((mf, d, "no header in Inputs")); wb.close(); continue
        H = {S(v): j for j, v in enumerate(rows[hi]) if S(v)}; hdrs[tuple(sorted(H))] += 1
        n = 0
        for r in rows[hi + 1:]:
            dn = r[H["Document Number"]] if H["Document Number"] < len(r) else None
            if dn in (None, "") or not str(dn).replace(".0", "").isdigit(): continue
            rec = {k: (r[j] if j < len(r) else None) for k, j in H.items()}; rec["_month"] = mf; rec["_state"] = d; out.append(rec); n += 1
        wb.close(); print("%s / %-20s %s rows %d" % (mf, d, sn, n), flush=True)
pickle.dump({"rows": out, "missing": miss}, open("fy2627_inputs.pkl", "wb"))
print("\nTOTAL rows:", len(out), "| missing:", miss); print("header variants:", len(hdrs))
for k, v in hdrs.items(): print("  x%d %s" % (v, list(k)[:28]))
iy = collections.Counter(S(r.get("Invoice Year")) for r in out); print("Invoice Year:", dict(iy))
py = [r for r in out if S(r.get("Invoice Year")) == "25-26"]; tax = lambda r: sum(float(r.get(c) or 0) for c in ("IGST", "CGST", "SGST"))
print("FY 25-26 dated rows claimed in FY 26-27 (Inputs): %d | tax %.2f | by month: %s" % (len(py), sum(tax(r) for r in py), {m: round(sum(tax(r) for r in py if r["_month"] == m)) for m in MONTHS}))
print("GST CREDIT values on those:", dict(collections.Counter(S(r.get("GST CREDIT")) for r in py)))
wb = openpyxl.load_workbook(r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx", read_only=True, data_only=True); n27 = wb["ITC Register 2026-27"]; hd = [c.value for c in next(n27.iter_rows(min_row=4, max_row=4))]; NH = {h: i for i, h in enumerate(hd) if h}
cur = [r for r in n27.iter_rows(min_row=5, values_only=True) if r[3]]
print("current ITC Register 2026-27: %d rows | tax %.2f | by claim month: %s" % (len(cur), sum(float(r[NH[c]] or 0) for r in cur for c in ("IGST", "CGST", "SGST")), dict(collections.Counter(S(r[NH["3B Claim  Month"]]) for r in cur))))
norm = lambda s: re.sub(r"[ \-/.'_]", "", S(s).upper())
keys = {(S(r.get("GSTN")).upper(), norm(r.get("Reference"))) for r in py}
print("current 26-27 rows matched in Inputs (GSTIN + invoice):", sum(1 for r in cur if (S(r[NH["Vendor GSTIN"]]).upper(), norm(r[NH["Invoice No."]])) in keys), "of", len(cur))
