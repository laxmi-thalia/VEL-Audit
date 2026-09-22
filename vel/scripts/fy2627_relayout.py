"""ITC Register 2026-27 rebuilt IN PLACE in the exact layout of ITC Register 2025-26 (Pawan 22-09: "use the SAME FORMAT entirely").
Header row 5 (rows 2-3 title/description, row 4 SUBTOTALs), columns A..BV identical to the 25-26 sheet (names, order, header
styling, widths, number formats, formulas re-based to this sheet's rows); the sheet's own extra columns follow after BV so nothing is
lost (Correct GSTIN, GSTR 9_Reporting, Reasons (GSTR 9), Available in 2B, 2B Inv, 2B Period (client), Final Remarks, Query, POS
Remarks, Source). External references (ITC Summary whole-column refs, GSTR-2B Apr25-Aug26 remark lookup, INDEX) are repointed by
header name. Single COM session; verifies 0 error cells, the 25-26 golden, ITC Summary values unchanged, the 26-27 totals, xlsb."""
import re, time, shutil, datetime as dt, collections, warnings, openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_relayout.xlsx")
S = lambda v: "" if v is None else str(v).strip()
E0 = dt.datetime(1899, 12, 30)
def xv(v):   # value for COM: naive datetimes -> Excel serials
    if isinstance(v, dt.datetime): return (v - E0).days + (v - E0).seconds / 86400.0
    if isinstance(v, dt.date): return (dt.datetime(v.year, v.month, v.day) - E0).days
    return v
S27, S25, SB2 = "ITC Register 2026-27", "ITC Register 2025-26", "GSTR-2B Apr25-Aug26"
# ---------------- read both sheets (formulas and cached values) --------------------------------------------------------
t0 = time.time()
wf = openpyxl.load_workbook(P, read_only=True); wv = openpyxl.load_workbook(P, read_only=True, data_only=True)
h25 = [S(c.value) for c in next(wf[S25].iter_rows(min_row=5, max_row=5))]
while h25 and not h25[-1]: h25.pop()
H25 = {h: i + 1 for i, h in enumerate(h25) if h}; N25 = 5 + sum(1 for r in wv[S25].iter_rows(min_row=6, values_only=True) if r[3])
r25f = list(next(wf[S25].iter_rows(min_row=6, max_row=6))); f25 = {L(i + 1): c.value for i, c in enumerate(r25f) if isinstance(c.value, str) and c.value.startswith("=")}
nf25 = {L(i + 1): getattr(c, "number_format", "General") for i, c in enumerate(r25f)}
row4 = {L(i + 1): c.value for i, c in enumerate(next(wf[S25].iter_rows(min_row=4, max_row=4))) if c.value is not None}
row2 = next(wv[S25].iter_rows(min_row=2, max_row=2, values_only=True))[0]
h27 = [S(c.value) for c in next(wf[S27].iter_rows(min_row=4, max_row=4))]
while h27 and not h27[-1]: h27.pop()
old = {}   # old letter -> header (both 'Reasons' disambiguated)
for i, h in enumerate(h27):
    if h == "Reasons": h = "Reasons (GSTR 9)" if L(i + 1) == "Y" else "Reasons"
    old[L(i + 1)] = h
OLD = {h: l for l, h in old.items()}
r27f = list(next(wf[S27].iter_rows(min_row=5, max_row=5))); f27 = {L(i + 1): c.value for i, c in enumerate(r27f) if isinstance(c.value, str) and c.value.startswith("=")}
desc27 = next(wv[S27].iter_rows(min_row=2, max_row=2, values_only=True))[0]
vals = [r for r in wv[S27].iter_rows(min_row=5, max_row=5 + 5000, values_only=True) if r[3]]
N = 5 + len(vals); V = {old[L(i + 1)]: [r[i] for r in vals] for i in range(len(h27))}
# external references
ext = []   # (sheet, address, formula)
for ws in wf.worksheets:
    if ws.title in (S27, SB2): continue
    for ri, row in enumerate(ws.iter_rows(), 1):
        for ci, c in enumerate(row, 1):
            if isinstance(c.value, str) and "ITC Register 2026-27" in c.value and c.value.startswith("="): ext.append((ws.title, "%s%d" % (L(ci), ri), c.value))
b2h = [S(c.value) for c in next(wf[SB2].iter_rows(min_row=2, max_row=2))]; B2H = {h: i + 1 for i, h in enumerate(b2h) if h}
b2f = wf[SB2].cell(3, B2H["Reco Remarks"]).value; NB = 2 + sum(1 for r in wv[SB2].iter_rows(min_row=3, values_only=True) if r[0])
wf.close(); wv.close()
print("read %.0fs | 25-26 %d cols rows 6..%d | 26-27 %d cols rows 5..%d -> new 6..%d | external refs %d" % (time.time() - t0, len(h25), N25, len(h27), 4 + len(vals), N, len(ext)))
# ---------------- column plan ------------------------------------------------------------------------------------------
RENAME = {"State Name": "STATE NAME", "VEL GSTIN": "VEL GSTN", "Vendor Name/RCM Category": "Vendor Name/RCM Category", "GSTR 2B/6A Period": "GSTR 2B/6A PERIOD",
          "2B Year": "2B YEAR", "TYPE": "Type", "Consider 8A reco": "Consider in 8A reco", "F.Y/Booking Year": "Posting Year", "Reasons": "Reasons"}
FORMULA_FROM_25 = {"Tax Rate", "Total GST", "KEY", "B_IGST", "B_CGST", "B_SGST", "B_Total GST", "2B_IGST", "2B_CGST", "2B_SGST", "2B_Total GST",
                   "D_IGST", "D_CGST", "D_SGST", "D_Total GST", "Invoice as per 2B", "2B Period", "My GSTN as per 2B", "Supplier GSTN as per 2B",
                   "Vendor", "My GSTN", "As per State", "As per Amounts", "POS Check", "POS Query", "Expense GL Element", "PO Number", "Expense Description", "ZFI06 status"}
EXTRAS = [("Correct GSTIN", "F"), ("GSTR 9_Reporting", "V"), ("Reasons (GSTR 9)", "V"), ("Available in 2B", "F"), ("2B Inv", "F"), ("2B Period", "F"),
          ("Final Remarks", "F"), ("Query", "V"), ("Remarks", "F"), ("Source", "V")]
EXTRA_HDR = {"2B Period": "2B Period (client lookup)", "Remarks": "POS Remarks", "Reasons (GSTR 9)": "Reasons (GSTR 9)"}
new_hdr = list(h25) + [EXTRA_HDR.get(h, h) for h, _ in EXTRAS]
NEW = {h: i + 1 for i, h in enumerate(new_hdr)}
# old letter -> new letter (for remapping the extras' own-sheet references and the external refs)
o2n = {}
for l, h in old.items():
    if h in NEW: o2n[l] = L(NEW[h])
    else:
        tgt = next((k for k, v in RENAME.items() if v == h), None)
        if tgt: o2n[l] = L(NEW[tgt])
for h, _ in EXTRAS: o2n[OLD[h]] = L(NEW[EXTRA_HDR.get(h, h)])
o2n.update({"AY": "Z", "AZ": "AA", "BA": "AB", "BB": "AC", "BC": "AD", "BD": "AE", "BE": "AF", "BF": "AG", "BG": "AH", "BH": "AI", "BI": "AJ", "BJ": "AK", "BK": "AL", "BL": "AM", "BM": "AN", "BN": "AO"})
missing = [l for l in old if l not in o2n]; assert not missing, ("unmapped old columns", [(l, old[l]) for l in missing])
Q = re.compile(r"'[^']+'!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?")   # sheet-qualified ranges (protected)
R = re.compile(r"(?<![A-Za-z0-9_!'])(\$?)([A-Z]{1,3})(\$?)(\d+)")
def remap_own(f):
    keep = []
    def prot(m): keep.append(m.group(0)); return "\x00%d\x00" % (len(keep) - 1)
    g = Q.sub(prot, f)
    def rep(m):
        col = m.group(2); row = int(m.group(4))
        if col not in o2n: raise SystemExit("cannot remap %s in %s" % (col, f))
        return "%s%s%s%d" % (m.group(1), o2n[col], m.group(3), row + 1)
    g = R.sub(rep, g)
    return re.sub(r"\x00(\d+)\x00", lambda m: keep[int(m.group(1))], g)
def remap_ext(f):
    """References to the 26-27 sheet: letters via o2n, data rows shifted by +1 (header 4 -> 5, data 5.. -> 6..); A1 hyperlinks untouched."""
    def cell(dc, col, dr, row):
        r = int(row) if row else None
        return "%s%s%s%s" % (dc, o2n.get(col, col), dr, (str(r + 1) if r and r >= 5 else row) if row else "")
    pat = r"'ITC Register 2026-27'!(\$?)([A-Z]{1,3})(\$?)(\d*)(?::(\$?)([A-Z]{1,3})(\$?)(\d*))?"
    return re.sub(pat, lambda m: "'ITC Register 2026-27'!" + cell(m.group(1), m.group(2), m.group(3), m.group(4)) + ((":" + cell(m.group(5), m.group(6), m.group(7), m.group(8))) if m.group(6) else ""), f)
# ---------------- build the value block + per-column formulas ----------------------------------------------------------
nrows = len(vals); block = [[None] * len(new_hdr) for _ in range(nrows)]; colf = {}
def put(h, seq):
    j = NEW[h] - 1
    for i, v in enumerate(seq): block[i][j] = xv(v)
for h in h25:
    if h in FORMULA_FROM_25:
        colf[h] = f25[L(H25[h])].replace(str(N25), str(N)); continue
    if h == "Matching of 12B of FY 25-26 and 12C of 24-25": put(h, ["NA"] * nrows); continue
    if h == "2B pull":
        seen = set(); pull = []
        for c, k in zip(V["Countif"], V["KEY2 (matched 2B key)"]):
            if c == "Consider" and S(k): pull.append("No" if k in seen else "Yes"); seen.add(k)
            else: pull.append(None)
        put(h, pull); continue
    if h == "2B Year": colf[h] = remap_own(f27[OLD["2B YEAR"]]); continue
    if h == "3B Claim  Month":   # 26-27 holds text '01 Apr 2026' -> real dates like the 25-26 sheet
        def cm(v):
            if isinstance(v, dt.datetime) or not S(v): return v
            for fmt in ("%d %b %Y", "%d %B %Y", "%d-%m-%Y", "%Y-%m-%d"):
                try: return dt.datetime.strptime(S(v), fmt)
                except ValueError: pass
            raise SystemExit("unparsed 3B Claim Month: %r" % v)
        put(h, [cm(v) for v in V[h]]); continue
    src = h if h in V else RENAME.get(h)
    if src in V: put(h, V[src])
    # else: no source in the 26-27 data (Nature of Services, Type for GSTR9, Material Description, Eligibility, Considered in Table 6A1, ...) -> blank
for h, kind in EXTRAS:
    nh = EXTRA_HDR.get(h, h)
    if kind == "V": put(nh, V[h])
    else: colf[nh] = remap_own(f27[OLD[h]])
print("formula columns %d | value columns %d | extras %d" % (len(colf), sum(1 for h in new_hdr if h not in colf), len(EXTRAS)))
# ---------------- COM ----------------------------------------------------------------------------------------------------
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    s27, s25, sb2, isum = wb.Worksheets(S27), wb.Worksheets(S25), wb.Worksheets(SB2), wb.Worksheets("ITC Summary")
    ur = isum.UsedRange; before = ur.Value
    reg25 = s25.Cells(4, H25["Total GST"]).Value
    assert s27.Cells(s27.Rows.Count, 4).End(-4162).Row == 4 + nrows, "26-27 row count drift"
    s27.AutoFilterMode = False; s27.Cells.Clear()
    s27.Cells(2, 1).Value = row2; s27.Cells(3, 1).Value = desc27
    s27.Cells(2, 1).Font.Bold = True
    s27.Range(s27.Cells(5, 1), s27.Cells(5, len(new_hdr))).Value = [new_hdr]
    s27.Range(s27.Cells(6, 1), s27.Cells(N, len(new_hdr))).Value = block
    for h, f in colf.items():
        c = L(NEW[h]); s27.Range("%s6:%s%d" % (c, c, N)).Formula = f
    # header formats + widths + number formats from the 25-26 sheet; extras styled like the 25-26 ZFI06-status header (blue)
    s25.Range("A5:%s5" % L(len(h25))).Copy(); s27.Range("A5:%s5" % L(len(h25))).PasteSpecial(-4122); xl.CutCopyMode = False
    for h in h25:
        c = L(H25[h]); s27.Columns(c).ColumnWidth = s25.Columns(c).ColumnWidth
        if nf25.get(c) and nf25[c] != "General": s27.Range("%s6:%s%d" % (c, c, N)).NumberFormat = nf25[c]
    for j in range(len(h25) + 1, len(new_hdr) + 1):
        hc = s27.Cells(5, j); hc.Font.Bold = True; hc.Font.Color = 0xFFFFFF; hc.Interior.Color = 0xB09784; s27.Columns(j).ColumnWidth = 14
    s27.Columns(NEW["Source"]).ColumnWidth = 44; s27.Columns(NEW["Reco Remarks"]).ColumnWidth = 60
    s27.Range("%s6:%s%d" % (L(NEW["2B Period (client lookup)"]), L(NEW["2B Period (client lookup)"]), N)).NumberFormat = "mmm-yy"
    for c, f in row4.items():
        s27.Cells(4, CI(c)).Formula = str(f).replace(str(N25), str(N)) if isinstance(f, str) and f.startswith("=") else f
        s27.Cells(4, CI(c)).Font.Bold = True
    s27.Range(s27.Cells(5, 1), s27.Cells(N, len(new_hdr))).AutoFilter()
    s27.Activate(); xl.ActiveWindow.FreezePanes = False; s27.Range("A6").Select(); xl.ActiveWindow.FreezePanes = True
    # external references
    for sh, addr, f in ext: wb.Worksheets(sh).Range(addr).Formula = remap_ext(f)
    # 2B sheet remark lookup (26-27 branch) -> new letters and rows
    nb2 = remap_ext(b2f).replace("$5:$", "$6:$").replace("$%d" % (4 + nrows), "$%d" % N)
    rc = L(B2H["Reco Remarks"]); sb2.Range("%s3:%s%d" % (rc, rc, NB)).Formula = nb2
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    # ---------------- verify ----------------
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    after = isum.UsedRange.Value; same = before == after
    if not same:
        diffs = [(i, j, a, b_) for i, (ra, rb) in enumerate(zip(before, after)) for j, (a, b_) in enumerate(zip(ra, rb)) if a != b_ and not (isinstance(a, float) and isinstance(b_, float) and abs(a - b_) < 0.01)]
        print("ITC Summary differences:", len(diffs), diffs[:8])
    col = lambda h: [v[0] for v in s27.Range("%s6:%s%d" % (L(NEW[h]), L(NEW[h]), N)).Value]
    num = lambda v: v if isinstance(v, (int, float)) else 0.0
    tot = {h: round(sum(num(v) for v in col(h)), 2) for h in ("Total GST", "B_Total GST", "2B_Total GST", "D_Total GST")}
    print("error cells %d | 25-26 golden %.2f | ITC Summary unchanged %s" % (e, s25.Cells(4, H25["Total GST"]).Value, same))
    print("26-27 totals:", tot, "| Countif", dict(collections.Counter(col("Countif"))), "| remarks", dict(collections.Counter(S(v).split(" – ")[0] for v in col("Reco Remarks"))))
    print("26-27 headers A..BV identical to 25-26:", [S(s27.Cells(5, j).Value) for j in range(1, len(h25) + 1)] == h25, "| extras:", [S(s27.Cells(5, j).Value) for j in range(len(h25) + 1, len(new_hdr) + 1)])
    b2d = collections.Counter(S(v[0]).split(" – ")[0] + (" (26-27)" if "(ITCR 26-27)" in S(v[0]) else "") for v in sb2.Range("%s3:%s%d" % (rc, rc, NB)).Value)
    print("2B sheet remarks:", dict(sorted(b2d.items())))
    assert e == 0 and same and abs(s25.Cells(4, H25["Total GST"]).Value - reg25) < 0.01 and abs(tot["Total GST"] - 24280458.10) < 1 and abs(tot["B_Total GST"] - 24280458.10) < 1, "VERIFICATION FAILED"
    wb.Save()
    if xlsb_free(): wb.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    wb.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
