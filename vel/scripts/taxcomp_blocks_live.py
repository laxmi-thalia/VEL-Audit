"""Pawan 24-09: the four difference blocks on 'Tax comp report' (S:U 4D1, V:X 2B glitch, Y:AA 4A5, AB:AD CN in 4A5) become
LIVE formulas instead of typed values.

Routing is read from the Reasons text (Q), which the tax-comparison engine writes one segment per reason, each segment
opening with the head(s) it concerns ("IGST Rs. ..." / "C/SGST Rs. ... each") and naming the block. For every (row, head):
  - one block named   -> that block cell = the row's own difference for the head (L / M / N), live; the other blocks 0
  - two blocks named  -> the smaller quoted amount is written as the split (a value, with a comment), the other block takes
                         the live remainder (difference minus the split) so the head still ties to L / M / N
  - "nets off / no impact" segments route nowhere; amounts under Rs 1 are suppressed (rounding noise)
Zeros are hidden by the number format, so the sheet reads as before. Row-5 totals become SUBTOTAL.
Safety: the typed values are read first; after recalc every cell is compared and the file is saved only if no cell moved by
more than Rs 1 and the workbook has 0 error cells - otherwise the rows are listed and nothing is saved."""
import re, shutil, time, warnings, openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for f in (P, B):
    try: open(f, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + f)
shutil.copy(P, "master2_snapshot_before_taxcomp_live.xlsx")
SHEET = "Tax comp report"; R0, R1 = 8, 211; S = lambda v: "" if v is None else str(v); n = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
BLOCKS = {"S": ["Not reported in 4D(1)"],
          "V": ["Mismatch in 2B"],
          "Y": ["Excess reported in 4A(5)", "Short reported in 4A(5)", "Not reported in 4A(5)", "Debit not excess reported in 4A5"],
          "AB": ["CN netted", "CN not considered in 4A(5)"]}
DIFF = ["L", "M", "N"]; HEADS = ["IGST", "CGST", "SGST"]
amt = lambda s: float(s.replace(",", ""))
def parse(q):
    """-> {head_index: [(block_col, quoted_amount_or_None), ...]}"""
    out = {0: [], 1: [], 2: []}
    for seg in q.split("\n"):
        head = seg.split("[")[0]
        if "nets off" in seg.lower() or "no impact" in seg.lower(): continue
        blk = next((b for b, kws in BLOCKS.items() if any(k.lower() in head.lower() for k in kws)), None)
        if not blk: continue
        mi = re.search(r"IGST Rs\. (-?[\d,]+(?:\.\d+)?)", head); mc = re.search(r"C/SGST Rs\. (-?[\d,]+(?:\.\d+)?)", head)
        if mi: out[0].append((blk, amt(mi.group(1))))
        if mc: out[1].append((blk, amt(mc.group(1)))); out[2].append((blk, amt(mc.group(1))))
        if not mi and not mc:
            for h in range(3): out[h].append((blk, None))
    return out
wb = openpyxl.load_workbook(P, read_only=True, data_only=True); ws = wb[SHEET]
assert ws.cell(7, CI("Q")).value == "Reasons" and ws.cell(7, CI("S")).value == "IGST" and ws.cell(7, CI("AD")).value == "SGST", "layout"
typed = {}; plan = {}; splits = []
for r in range(R0, R1 + 1):
    for c in range(CI("S"), CI("AD") + 1): typed[(r, c)] = round(n(ws.cell(r, c).value), 2)
    q = S(ws.cell(r, CI("Q")).value)
    if not q: continue
    routed = parse(q)
    for h in range(3):
        items = routed[h]
        if not items: continue
        d = "%s%d" % (DIFF[h], r)
        if len(items) == 1:
            plan[(r, CI(items[0][0]) + h)] = "=IF(ABS(%s)>=1,%s,0)" % (d, d)
        else:
            assert len(items) == 2 and all(a is not None for _, a in items), (r, h, items)
            (b1, a1), (b2, a2) = sorted(items, key=lambda t: abs(t[1]))     # smaller quoted amount is the split
            c_split = CI(b1) + h; plan[(r, c_split)] = a1; splits.append((r, L(c_split), a1))
            plan[(r, CI(b2) + h)] = "=IF(ABS(%s)>=1,%s-%s%d,0)" % (d, d, L(c_split), r)
wb.close()
print("typed non-zero cells: %d | routed cells: %d (%d formulas, %d split values at %s)" % (sum(1 for v in typed.values() if v), len(plan), sum(1 for v in plan.values() if isinstance(v, str)), len(splits), splits))
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = w.Worksheets(SHEET)
    tot_before = [round(n(sh.Cells(5, c).Value), 2) for c in range(CI("S"), CI("AD") + 1)]
    for r in range(R0, R1 + 1):
        for c in range(CI("S"), CI("AD") + 1):
            v = plan.get((r, c), 0)
            if isinstance(v, str): sh.Cells(r, c).Formula = v
            else:
                sh.Cells(r, c).Value = v
                if v: sh.Cells(r, c).AddComment("Split per the Reasons text (two reasons on one head); the other block carries the live remainder. Value, not formula.")
    for c in range(CI("S"), CI("AD") + 1): sh.Cells(5, c).Formula = "=SUBTOTAL(9,%s%d:%s%d)" % (L(c), R0, L(c), R1)
    sh.Range("S%d:AD%d" % (R0, R1)).NumberFormat = "#,##0.00;-#,##0.00;"
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    moved = [(r, L(c), old, round(n(sh.Cells(r, c).Value), 2)) for (r, c), old in typed.items() if abs(round(n(sh.Cells(r, c).Value), 2) - old) > 1]
    tot_after = [round(n(sh.Cells(5, c).Value), 2) for c in range(CI("S"), CI("AD") + 1)]
    e = 0
    for s_ in w.Worksheets:
        if s_.Name == "T6A1 Extract - 24-25": continue
        try: e += s_.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("block totals before %s\n             after  %s" % (tot_before, tot_after))
    print("cells moved by more than Rs 1: %d %s | error cells %d | golden %.2f" % (len(moved), moved[:12], e, w.Worksheets("ITC Register 2025-26").Range("V4").Value))
    if moved or e:
        w.Close(False); raise SystemExit("NOT SAVED - see the rows above")
    w.Save(); w.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); print("re-open in a fresh instance: OK (%.0fs)" % (time.time() - t0)); w.SaveAs(B, FileFormat=50); w.Close(False); print("xlsb exported")
finally: xl.Quit()
