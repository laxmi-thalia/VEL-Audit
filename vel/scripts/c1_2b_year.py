"""C1 (Rashid 28-09): ITC Register 2025-26 `2B Year` follows the reco's `2B Period` where there is one.

Live formula per line: FY of `2B Period` (AQ, the KEY2 lookup into the 2B sheet) when it is filled, else the value
the cell already held (the client's own year, kept as the fallback literal). Fills the blanks and corrects the
lines whose client year disagreed with the matched 2B period. Works on the .xlsb directly through Excel.
"""
import datetime as dt
import os
import sys
import time
from collections import Counter

import win32com.client as w

P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2) (1) (1).xlsb"
SHEET = "ITC Register 2025-26"
HDR = 5
FIRST = 6


def fy(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 40000 < v < 60000:
        v = dt.datetime(1899, 12, 30) + dt.timedelta(days=int(v))
    if not hasattr(v, "year"):
        return None
    y = v.year if v.month >= 4 else v.year - 1
    return f"{y}-{str(y + 1)[2:]}"


def col_of(ws, name):
    for j in range(1, 120):
        if ws.Cells(HDR, j).Value == name:
            return j
    raise SystemExit(f"header {name!r} not found on row {HDR}")


def errors_in(ws):
    try:
        return ws.UsedRange.SpecialCells(-4123, 16).Count  # xlCellTypeFormulas, xlErrors
    except Exception:  # noqa: BLE001 - no error cells
        return 0


def L(j):
    s = ""
    while j:
        j, r = divmod(j - 1, 26)
        s = chr(65 + r) + s
    return s


xl = w.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False
t0 = time.time()
try:
    wb = xl.Workbooks.Open(P, UpdateLinks=0)
    ws = wb.Worksheets(SHEET)
    cq, cy, ct = col_of(ws, "2B Period"), col_of(ws, "2B Year"), col_of(ws, "Total GST")
    # TRAP: the sheet carries an AutoFilter with most rows hidden - End(xlUp) stops at the last VISIBLE row and an
    # array write skips hidden rows. Take the bound from UsedRange and show every row before writing.
    ur = ws.UsedRange
    last = ur.Row + ur.Rows.Count - 1
    while last > FIRST and ws.Cells(last, ct).Value in (None, ""):
        last -= 1
    filtered = bool(ws.FilterMode)
    if filtered:
        ws.ShowAllData()
    hidden = sum(1 for r in range(FIRST, last + 1, 997) if ws.Rows(r).Hidden)   # coarse probe
    ws.Range(ws.Rows(FIRST), ws.Rows(last)).Hidden = False
    print(f"opened in {time.time()-t0:.0f}s; 2B Period={L(cq)} 2B Year={L(cy)} Total GST={L(ct)} rows {FIRST}..{last}; "
          f"filter cleared={filtered}, hidden-row probe hits={hidden}")
    err_before = errors_in(ws)
    gold_before = xl.WorksheetFunction.Sum(ws.Range(ws.Cells(FIRST, ct), ws.Cells(last, ct)))
    old = [r[0] for r in ws.Range(ws.Cells(FIRST, cy), ws.Cells(last, cy)).Value]
    per = [r[0] for r in ws.Range(ws.Cells(FIRST, cq), ws.Cells(last, cq)).Value]
    assert not any(isinstance(v, str) and v.startswith("=") for v in old), "2B Year already holds formulas"

    q = L(cq)
    out = []
    stats = Counter()
    for i, (o, p) in enumerate(zip(old, per), FIRST):
        lit = "" if o is None else str(o).strip()
        lit_x = lit.replace('"', '""')
        f = (f'=IF(${q}{i}="","{lit_x}",IF(MONTH(${q}{i})>=4,YEAR(${q}{i})&"-"&RIGHT(YEAR(${q}{i})+1,2),'
             f'YEAR(${q}{i})-1&"-"&RIGHT(YEAR(${q}{i}),2)))')
        out.append([f])
        new = fy(p) if p not in (None, "") else lit
        if p not in (None, ""):
            if not lit:
                stats[("filled from 2B Period", new)] += 1
            elif lit != new:
                stats[(f"corrected {lit} -> {new}",)] += 1
            else:
                stats[("already agreed",)] += 1
        else:
            stats[("no 2B Period: kept " + ("client value" if lit else "blank"),)] += 1
    step = 4000
    for s in range(0, len(out), step):
        ws.Range(ws.Cells(FIRST + s, cy), ws.Cells(FIRST + s + len(out[s:s + step]) - 1, cy)).Formula = out[s:s + step]
    print(f"formulas written in {time.time()-t0:.0f}s")
    # every line must reference its OWN row (the filtered-write trap put row 12's formula on row 2827)
    got = [r[0] for r in ws.Range(ws.Cells(FIRST, cy), ws.Cells(last, cy)).Formula]
    bad = sum(1 for i, f in enumerate(got, FIRST) if f'${q}{i}=' not in str(f))
    assert bad == 0, f"{bad} lines carry a formula for another row - aborting without saving"
    xl.Calculate()
    err_after = errors_in(ws)
    gold_after = xl.WorksheetFunction.Sum(ws.Range(ws.Cells(FIRST, ct), ws.Cells(last, ct)))
    newvals = [r[0] for r in ws.Range(ws.Cells(FIRST, cy), ws.Cells(last, cy)).Value]
    print("2B Year after:", Counter(str(v or "").strip() or "blank" for v in newvals).most_common(6))
    for k, v in sorted(stats.items(), key=lambda t: -t[1]):
        print("  ", " ".join(k), v)
    print(f"error cells on the register: before {err_before} after {err_after}; Total GST {gold_before:,.2f} -> {gold_after:,.2f}")
    assert err_after == err_before and abs(gold_after - gold_before) < 0.005
    wb.Save()
    print(f"saved {os.path.basename(P)} in {time.time()-t0:.0f}s")
    wb.Close(False)
finally:
    xl.Quit()

# fresh instance re-open: the file must open cleanly
xl2 = w.DispatchEx("Excel.Application")
xl2.Visible = False
xl2.DisplayAlerts = False
try:
    wb2 = xl2.Workbooks.Open(P, UpdateLinks=0, ReadOnly=True)
    ws2 = wb2.Worksheets(SHEET)
    print("fresh-instance reopen OK; AZ6 =", ws2.Range("AZ6").Formula[:60], "->", ws2.Range("AZ6").Value)
    wb2.Close(False)
finally:
    xl2.Quit()
