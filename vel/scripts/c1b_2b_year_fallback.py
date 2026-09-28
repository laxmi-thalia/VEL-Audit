"""C1b (Rashid 28-09): `2B Year` also follows the client's own `GSTR 2B/6A Period` when the reco has no 2B Period.

Order per line: FY of `2B Period` (AQ, reco) -> FY of `GSTR 2B/6A Period` (AY, client, when it is a date) -> the value
the cell held before C1 (kept as the literal fallback). Fixes the lines whose client year said 2024-25 against a
FY 25-26 period (and the reverse), and fills the blanks that only have the client's period.
"""
import datetime as dt
import re
import time
from collections import Counter

import win32com.client as w

P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2) (1) (1).xlsb"
SHEET, HDR, FIRST = "ITC Register 2025-26", 5, 6
LIT_RE = re.compile(r'^=IF\(\$AQ\d+="","((?:[^"]|"")*)",')


def fy(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 40000 < v < 60000:
        v = dt.datetime(1899, 12, 30) + dt.timedelta(days=int(v))
    if not hasattr(v, "year"):
        return None
    if getattr(v, "hour", 0) == 18 and v.minute == 30:
        v = v + dt.timedelta(hours=6)
    y = v.year if v.month >= 4 else v.year - 1
    return f"{y}-{str(y + 1)[2:]}"


def col_of(ws, name):
    for j in range(1, 120):
        if ws.Cells(HDR, j).Value == name:
            return j
    raise SystemExit(f"header {name!r} missing")


def L(j):
    s = ""
    while j:
        j, r = divmod(j - 1, 26)
        s = chr(65 + r) + s
    return s


def errors_in(ws):
    try:
        return ws.UsedRange.SpecialCells(-4123, 16).Count
    except Exception:  # noqa: BLE001
        return 0


def fy_expr(c):
    return f'IF(MONTH({c})>=4,YEAR({c})&"-"&RIGHT(YEAR({c})+1,2),YEAR({c})-1&"-"&RIGHT(YEAR({c}),2))'


xl = w.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False
t0 = time.time()
try:
    wb = xl.Workbooks.Open(P, UpdateLinks=0)
    ws = wb.Worksheets(SHEET)
    cq, cy, ca, ct = (col_of(ws, h) for h in ("2B Period", "2B Year", "GSTR 2B/6A Period", "Total GST"))
    ur = ws.UsedRange
    last = ur.Row + ur.Rows.Count - 1
    while last > FIRST and ws.Cells(last, ct).Value in (None, ""):
        last -= 1
    if ws.FilterMode:
        ws.ShowAllData()
    ws.Range(ws.Rows(FIRST), ws.Rows(last)).Hidden = False
    q, a = L(cq), L(ca)
    assert (q, L(cy), a) == ("AQ", "AZ", "AY"), (q, L(cy), a)
    err0 = errors_in(ws)
    gold0 = xl.WorksheetFunction.Sum(ws.Range(ws.Cells(FIRST, ct), ws.Cells(last, ct)))
    cur_f = [r[0] for r in ws.Range(ws.Cells(FIRST, cy), ws.Cells(last, cy)).Formula]
    cur_v = [r[0] for r in ws.Range(ws.Cells(FIRST, cy), ws.Cells(last, cy)).Value]
    per = [r[0] for r in ws.Range(ws.Cells(FIRST, cq), ws.Cells(last, cq)).Value]
    cli = [r[0] for r in ws.Range(ws.Cells(FIRST, ca), ws.Cells(last, ca)).Value]
    print(f"opened, {last - FIRST + 1} lines, {time.time()-t0:.0f}s")
    out, stats = [], Counter()
    for i, (f, v, p, c) in enumerate(zip(cur_f, cur_v, per, cli), FIRST):
        m = LIT_RE.match(str(f))
        assert m, f"row {i}: 2B Year does not carry the C1 formula: {str(f)[:60]!r}"
        lit_x = m.group(1)                       # already ""-escaped
        lit = lit_x.replace('""', '"')
        out.append([f'=IF(${q}{i}<>"",{fy_expr(f"${q}{i}")},IF(ISNUMBER(${a}{i}),{fy_expr(f"${a}{i}")},"{lit_x}"))'])
        shown = (str(v).strip() if v else "")
        new = fy(p) if p not in (None, "") else (fy(c) if fy(c) else lit)
        if p in (None, "") and fy(c):
            if not lit:
                stats[("filled from client period", new)] += 1
            elif lit != new:
                stats[(f"client year {lit} -> period year {new}",)] += 1
            else:
                stats[("client year already agrees with client period",)] += 1
    step = 4000
    for s in range(0, len(out), step):
        ws.Range(ws.Cells(FIRST + s, cy), ws.Cells(FIRST + s + len(out[s:s + step]) - 1, cy)).Formula = out[s:s + step]
    got = [r[0] for r in ws.Range(ws.Cells(FIRST, cy), ws.Cells(last, cy)).Formula]
    bad = sum(1 for i, g in enumerate(got, FIRST) if f"${q}{i}<>" not in str(g) or f"${a}{i})" not in str(g))
    assert bad == 0, f"{bad} misaligned formulas - not saving"
    xl.Calculate()
    err1 = errors_in(ws)
    gold1 = xl.WorksheetFunction.Sum(ws.Range(ws.Cells(FIRST, ct), ws.Cells(last, ct)))
    after = [r[0] for r in ws.Range(ws.Cells(FIRST, cy), ws.Cells(last, cy)).Value]
    print("2B Year after:", Counter(str(x or "").strip() or "blank" for x in after).most_common(6))
    for k, n in sorted(stats.items(), key=lambda t: -t[1]):
        print("  ", " ".join(k), n)
    print(f"error cells {err0} -> {err1}; Total GST {gold0:,.2f} -> {gold1:,.2f}")
    assert err1 == err0 and abs(gold1 - gold0) < 0.005
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
    print("fresh reopen OK; AZ6 ->", wb2.Worksheets(SHEET).Range("AZ6").Value)
    wb2.Close(False)
finally:
    xl2.Quit()
