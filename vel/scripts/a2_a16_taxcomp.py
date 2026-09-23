"""A2 + A16 (Pawan 23-09, Priyesh M2) on 'Tax comp report'.

This year's sheet already carries every column last year's did (Q Reasons, R Impact, AH..AO the 9C-note block). What is
missing is the CONTENT Priyesh wants carried across "in full": last year's reason vocabulary and its GSTR-9C note templates.
This script (a) writes the FY 24-25 vocabulary - every distinct Reasons (Q) text, every distinct 9C-note (AI) text and the
three Note-1/Note-2 formula templates - as a reference block to the right of the table, (b) attaches dropdown suggestions
(non-blocking) on Q and AI so this year's rows are labelled from the same set, and (c) widens the totals-row SUBTOTALs to
row 5000 so rows added later stay inside them. Nothing is decided for the CA: the 30 'Pending' months keep their state."""
import collections, shutil, time, warnings, openpyxl, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
warnings.filterwarnings("ignore")
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_a2.xlsx")
S = lambda v: "" if v is None else str(v).strip()
# ---- last year's vocabulary and templates
wv = openpyxl.load_workbook("ly_9c_fy2425.xlsx", read_only=True, data_only=True); ly = wv["Tax comp report"]
rows = [r for r in ly.iter_rows(min_row=7, values_only=True) if r[0] is not None and r[1] is not None]
reasons = [k for k, _ in collections.Counter(S(r[CI("Q") - 1]) for r in rows if S(r[CI("Q") - 1])).most_common()]
notes = [k for k, _ in collections.Counter(S(r[CI("AI") - 1]) for r in rows if S(r[CI("AI") - 1])).most_common()]
impacts = [k for k, _ in collections.Counter(S(r[CI("R") - 1]) for r in rows if S(r[CI("R") - 1])).most_common()]
wv.close()
wf = openpyxl.load_workbook("ly_9c_fy2425.xlsx", read_only=True); lyf = wf["Tax comp report"]
templates = []
for r in lyf.iter_rows(min_row=7, max_row=240):
    for c in ("AL", "AM"):
        v = r[CI(c) - 1].value
        if isinstance(v, str) and v.startswith("=CONCATENATE") and v not in [t[1] for t in templates]:
            templates.append(("%s row %d" % (c, r[0].row), v))
wf.close()
print("LY vocabulary: %d reasons, %d notes, %d impacts, %d note templates" % (len(reasons), len(notes), len(impacts), len(templates)))
# ---- COM
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); xl.Calculation = -4135; ws = w.Worksheets("Tax comp report")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    tot_before = [ws.Cells(5, CI(c)).Value for c in ("S", "V", "Y", "AB", "AE")]
    # (c) widen the totals-row SUBTOTALs
    widened = 0
    for c in range(1, ws.Cells(7, ws.Columns.Count).End(-4159).Column + 1):
        f = ws.Cells(5, c).Formula
        if isinstance(f, str) and f.upper().startswith("=SUBTOTAL("):
            import re
            ws.Cells(5, c).Formula = re.sub(r"(\$?[A-Z]{1,3}\$?)(\d+)\)", lambda m: m.group(1) + "5000)", f); widened += 1
    # (a) vocabulary block at the far right, clear of the LY pivot area (AR..BA)
    c0 = CI("BD"); hdrs = ["Reasons vocabulary – FY 24-25 tax comp (col Q)", "Impact vocabulary – FY 24-25 (col R)",
                           "GSTR-9C note vocabulary – FY 24-25 (col AI)", "GSTR-9C Note 1 / Note 2 formula templates – FY 24-25 (cols AL / AM)"]
    ws.Cells(6, c0).Value = "Carried forward from last year's Tax comp report (A2, Priyesh 23-09: 'we take the entire thing'). Pick from these; add new ones below the list."
    ws.Cells(6, c0).Font.Italic = True
    for k, hd in enumerate(hdrs):
        h = ws.Cells(7, c0 + k); h.Value = hd; h.Font.Bold = True; h.Font.Color = 0xFFFFFF; h.Interior.Color = 0x4F3F33; ws.Columns(c0 + k).ColumnWidth = 60
    for i, v in enumerate(reasons): ws.Cells(8 + i, c0).Value = v
    for i, v in enumerate(impacts): ws.Cells(8 + i, c0 + 1).Value = v
    for i, v in enumerate(notes): ws.Cells(8 + i, c0 + 2).Value = v
    for i, (src, f) in enumerate(templates): ws.Cells(8 + i, c0 + 3).Value = "'" + f + "   [LY " + src + "]"
    # (b) non-blocking dropdown suggestions on Q, R and AI for the data rows
    for col, vc, n in (("Q", c0, len(reasons)), ("R", c0 + 1, len(impacts)), ("AI", c0 + 2, len(notes))):
        rng = ws.Range("%s8:%s5000" % (col, col)); rng.Validation.Delete()
        rng.Validation.Add(3, 1, 1, "=$%s$8:$%s$%d" % (L(vc), L(vc), 7 + n))   # xlValidateList, xlValidAlertInformation
        rng.Validation.ShowError = False; rng.Validation.IgnoreBlank = True; rng.Validation.InCellDropdown = True
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for sh in w.Worksheets:
        try: e += sh.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    tot_after = [ws.Cells(5, CI(c)).Value for c in ("S", "V", "Y", "AB", "AE")]
    print("data rows 8..%d | SUBTOTALs widened %d | totals before %s after %s | error cells %d | vocabulary at %s..%s" % (last, widened, [round(x or 0, 2) for x in tot_before], [round(x or 0, 2) for x in tot_after], e, L(c0), L(c0 + 3)))
    assert e == 0 and all(abs((a or 0) - (b or 0)) < 0.01 for a, b in zip(tot_before, tot_after)), "VERIFICATION FAILED"
    w.Save(); w.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); w = xl.Workbooks.Open(P); print("re-open in a fresh instance: OK (%.0fs)" % (time.time() - t0))
    if xlsb_free(): w.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    w.Close(False)
finally: xl.Quit()
