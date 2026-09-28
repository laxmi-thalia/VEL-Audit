"""C3 (Rashid 28-09): ITC Register 2025-26 - `Reco Remarks` on every Not-consider line = the remark of its document's
Consider line (values, byte-identical text, so a filter on a remark shows the whole document). Grouped by the sheet's
own KEY (vendor GSTIN & invoice, the Countif grouping). KEY2 / B_ / 2B_ / D_ untouched (they test Countif="Consider").
Supersedes the 22-09 'Consider line only' ruling for this file. cascade_fix.consolidate() must follow before any re-run.
"""
import time
from collections import Counter

import win32com.client as w

P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2) (1) (1).xlsb"
SHEET, HDR, FIRST = "ITC Register 2025-26", 5, 6


def col_of(ws, *names):
    heads = {str(ws.Cells(HDR, j).Value).strip(): j for j in range(1, 120) if ws.Cells(HDR, j).Value}
    for n in names:
        if n in heads:
            return heads[n]
    raise SystemExit(f"none of {names} on row {HDR}")


def errors_in(ws):
    try:
        return ws.UsedRange.SpecialCells(-4123, 16).Count
    except Exception:  # noqa: BLE001
        return 0


xl = w.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False
t0 = time.time()
try:
    wb = xl.Workbooks.Open(P, UpdateLinks=0)
    ws = wb.Worksheets(SHEET)
    ck, cc, cr, ct = col_of(ws, "KEY"), col_of(ws, "Countif"), col_of(ws, "Reco Remarks"), col_of(ws, "Total GST")
    ur = ws.UsedRange
    last = ur.Row + ur.Rows.Count - 1
    while last > FIRST and ws.Cells(last, ct).Value in (None, ""):
        last -= 1
    if ws.FilterMode:
        ws.ShowAllData()
    ws.Range(ws.Rows(FIRST), ws.Rows(last)).Hidden = False
    err0 = errors_in(ws)
    gold0 = xl.WorksheetFunction.Sum(ws.Range(ws.Cells(FIRST, ct), ws.Cells(last, ct)))
    key = [r[0] for r in ws.Range(ws.Cells(FIRST, ck), ws.Cells(last, ck)).Value]
    cnt = [str(r[0] or "").strip() for r in ws.Range(ws.Cells(FIRST, cc), ws.Cells(last, cc)).Value]
    rem_f = [r[0] for r in ws.Range(ws.Cells(FIRST, cr), ws.Cells(last, cr)).Formula]
    rem = [r[0] for r in ws.Range(ws.Cells(FIRST, cr), ws.Cells(last, cr)).Value]
    assert not any(str(f).startswith("=") for f in rem_f), "Reco Remarks holds formulas - stop"
    print(f"opened, {last - FIRST + 1} lines in {time.time()-t0:.0f}s; Countif: {Counter(cnt).most_common(4)}")

    consider_remark = {}
    for k, c, r in zip(key, cnt, rem):
        if c == "Consider" and k not in (None, ""):
            consider_remark.setdefault(k, r)
    stats = Counter()
    targets = {}   # row -> text
    for i, (k, c, r) in enumerate(zip(key, cnt, rem), FIRST):
        if not c.startswith("Not consider"):
            continue
        doc = consider_remark.get(k)
        if doc in (None, ""):
            stats["not consider: no Consider remark found (left as is)"] += 1
            continue
        if r not in (None, "") and str(r).strip():
            stats["not consider: already had a remark" + (" (same)" if str(r).strip() == str(doc).strip() else " (different, kept)")] += 1
            continue
        targets[i] = doc
        stats["filled: " + str(doc).split(" – ")[0]] += 1
    # TRAP (cost 45 min of CPU): the 2B sheet's remark lookups depend on this column - with automatic calculation every
    # small write triggers a full recalc. Manual calculation, ONE column write, one Calculate at the end.
    xl.Calculation = -4135   # xlCalculationManual
    ws.Range(ws.Cells(FIRST, cr), ws.Cells(last, cr)).NumberFormat = "@"
    newcol = [[targets.get(i, r if r not in (None, "") else None)] for i, r in enumerate(rem, FIRST)]
    ws.Range(ws.Cells(FIRST, cr), ws.Cells(last, cr)).Value = newcol
    print(f"{len(targets)} cells filled (column written once), {time.time()-t0:.0f}s")
    xl.Calculation = -4105   # xlCalculationAutomatic
    xl.Calculate()
    print(f"recalculated, {time.time()-t0:.0f}s")
    after = [r[0] for r in ws.Range(ws.Cells(FIRST, cr), ws.Cells(last, cr)).Value]
    mism = sum(1 for i, t in targets.items() if str(after[i - FIRST]) != str(t))
    err1 = errors_in(ws)
    gold1 = xl.WorksheetFunction.Sum(ws.Range(ws.Cells(FIRST, ct), ws.Cells(last, ct)))
    blank_nc = sum(1 for c, r in zip(cnt, after) if c.startswith("Not consider") and not (r and str(r).strip()))
    for k, v in sorted(stats.items(), key=lambda t: -t[1]):
        print("  ", k, v)
    print(f"read-back mismatches {mism}; Not-consider lines still blank {blank_nc}; errors {err0} -> {err1}; Total GST {gold0:,.2f} -> {gold1:,.2f}")
    assert mism == 0 and err1 == err0 and abs(gold1 - gold0) < 0.005
    tie = {}
    for c, r in zip(cnt, after):
        h = str(r or "").split(" – ")[0] or "blank"
        tie[(h, "Consider" if c == "Consider" else "Not consider" if c.startswith("Not consider") else "other")] = tie.get((h, "Consider" if c == "Consider" else "Not consider" if c.startswith("Not consider") else "other"), 0) + 1
    print("remark no x Countif:", sorted(tie.items()))
    wb.Save()
    print(f"saved in {time.time()-t0:.0f}s")
    wb.Close(False)
finally:
    xl.Quit()
xl2 = w.DispatchEx("Excel.Application")
xl2.Visible = False
xl2.DisplayAlerts = False
try:
    wb2 = xl2.Workbooks.Open(P, UpdateLinks=0, ReadOnly=True)
    print("fresh reopen OK, sheets", wb2.Worksheets.Count)
    wb2.Close(False)
finally:
    xl2.Quit()
