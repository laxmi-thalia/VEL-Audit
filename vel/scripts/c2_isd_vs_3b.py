"""C2 (Rashid 28-09): new sheet `2B ISD vs 3B 4A(4)` - ISD credit in the 2B (Octa `2B ISD Apr25-Aug26`) against
GSTR-3B Table 4A(4) (`3B Data`), per GSTIN per month, all live SUMIFS; per-GSTIN block on the right.

Sign: Diff = 2B ISD - 3B 4A(4); positive = credit in the 2B not claimed in 3B that month. Cumulative diff within
the GSTIN tells timing apart from a real gap. Rows mirror `3B Data` (one line per GSTIN-month, FY 25-26).
"""
import time

import win32com.client as w

P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2) (1) (1).xlsb"
NAME = "2B ISD vs 3B 4A(4)"
AFTER = "ITCR vs 3B Net ITC"
ISD, B3 = "'2B ISD Apr25-Aug26'", "'3B Data'"
ISD_R, B3_R = "$3:$2000", "$3:$400"          # generous bounds: both sheets are small
HDR, FIRST = 5, 6
MONEY = "#,##0.00"

HEADERS = ["State", "GSTIN", "Month", "Month (date)",
           "2B ISD IGST", "2B ISD CGST", "2B ISD SGST", "2B ISD Total",
           "3B 4A(4) IGST", "3B 4A(4) CGST", "3B 4A(4) SGST", "3B 4A(4) Total",
           "Diff IGST (2B - 3B)", "Diff CGST (2B - 3B)", "Diff SGST (2B - 3B)", "Diff Total (2B - 3B)",
           "Cumulative diff (GSTIN)", "Remarks (live)", "DPS Remarks (type here)"]
SUM_HEADERS = ["State", "GSTIN", "2B ISD Total", "3B 4A(4) Total", "Diff (2B - 3B)", "Remarks (live)"]
SUM_COL = 21   # U


def L(j):
    s = ""
    while j:
        j, r = divmod(j - 1, 26)
        s = chr(65 + r) + s
    return s


def rng(sheet, col, bound):
    a, b = bound.split(":")
    return f"{sheet}!${col}{a}:${col}{b}"


def isd(col, i):
    return (f'SUMIFS({rng(ISD, col, ISD_R)},{rng(ISD, "A", ISD_R)},$B{i},'
            f'{rng(ISD, "B", ISD_R)},">="&$D{i},{rng(ISD, "B", ISD_R)},"<"&EDATE($D{i},1))')


def b3(col, i):
    return f'SUMIFS({rng(B3, col, B3_R)},{rng(B3, "B", B3_R)},$B{i},{rng(B3, "C", B3_R)},$C{i})'


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
    src = wb.Worksheets("3B Data")
    n = 0
    while src.Cells(3 + n, 2).Value not in (None, ""):
        n += 1
    gstins, seen = [], set()
    for k in range(n):
        g, st = src.Cells(3 + k, 2).Value, src.Cells(3 + k, 1).Value
        if g not in seen:
            seen.add(g)
            gstins.append((st, g))
    print(f"opened in {time.time()-t0:.0f}s; 3B Data rows {n}, GSTINs {len(gstins)}")
    for s in wb.Worksheets:
        if s.Name == NAME:
            s.Delete()
    ws = wb.Worksheets.Add(After=wb.Worksheets(AFTER))
    ws.Name = NAME
    last = FIRST + n - 1

    ws.Range("A1").Value = "2B ISD vs GSTR-3B Table 4A(4) - FY 2025-26"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A2").Value = ("ISD credit per the GSTR-2B ISD section (sheet '2B ISD Apr25-Aug26', Octa export) against ITC claimed "
                            "in GSTR-3B Table 4A(4) (sheet '3B Data'), one line per GSTIN per month. Diff = 2B ISD less 3B 4A(4): "
                            "positive = credit in the 2B not claimed that month. Cumulative diff runs within the GSTIN, so a timing "
                            "gap that clears later reads as timing. FY 26-27 ISD rows in the 2B sheet are outside this year's 3B and "
                            "are not in these totals. All amounts are live SUMIFS.")
    ws.Range("A2").Font.Italic = True
    for j, h in enumerate(HEADERS, 1):
        c = ws.Cells(HDR, j)
        c.Value = h
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.Interior.Color = 0x4F3F33   # BGR of 333F4F
    for j, h in enumerate(SUM_HEADERS, SUM_COL):
        c = ws.Cells(HDR, j)
        c.Value = h
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.Interior.Color = 0x8497B0   # BGR of B09784 (the block fill)
    ws.Cells(HDR - 1, SUM_COL).Value = "Per GSTIN (FY 2025-26)"
    ws.Cells(HDR - 1, SUM_COL).Font.Bold = True

    rows = []
    for k in range(n):
        i = FIRST + k
        r3 = 3 + k
        rows.append([
            f"={B3}!A{r3}", f"={B3}!B{r3}", f"={B3}!C{r3}",
            f'=DATE(2000+VALUE(RIGHT($C{i},2)),MATCH(LEFT($C{i},3),{{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}},0),1)',
            f"={isd('J', i)}", f"={isd('K', i)}", f"={isd('L', i)}", f"=E{i}+F{i}+G{i}",
            f"={b3('S', i)}", f"={b3('T', i)}", f"={b3('U', i)}", f"=I{i}+J{i}+K{i}",
            f"=E{i}-I{i}", f"=F{i}-J{i}", f"=G{i}-K{i}", f"=H{i}-L{i}",
            f'=SUMIFS($P${FIRST}:$P${last},$B${FIRST}:$B${last},$B{i},$D${FIRST}:$D${last},"<="&$D{i})',
            (f'=IF(ABS(P{i})<10,"Matched",IF(P{i}>0,"ISD in 2B not claimed in 3B - "&TEXT(P{i},"#,##0"),'
             f'"Claimed in 3B more than 2B ISD - "&TEXT(-P{i},"#,##0"))&IF(ABS(Q{i})<10," - timing, cleared cumulatively",""))'),
            None,
        ])
    ws.Range(ws.Cells(FIRST, 1), ws.Cells(last, len(HEADERS))).Formula = rows
    for j in range(5, 18):
        ws.Range(ws.Cells(FIRST, j), ws.Cells(last, j)).NumberFormat = MONEY
        c = ws.Cells(HDR - 1, j)
        if j != 17:
            c.Formula = f"=SUBTOTAL(9,{L(j)}{FIRST}:{L(j)}{last})"
            c.NumberFormat = MONEY
            c.Font.Bold = True
    ws.Range(ws.Cells(FIRST, 4), ws.Cells(last, 4)).NumberFormat = "mmm-yy"

    # per-GSTIN block
    srow = []
    for k, (st, g) in enumerate(gstins):
        i = FIRST + k
        srow.append([
            f"=INDEX($A${FIRST}:$A${last},MATCH($V{i},$B${FIRST}:$B${last},0))", g,
            f"=SUMIFS($H${FIRST}:$H${last},$B${FIRST}:$B${last},$V{i})",
            f"=SUMIFS($L${FIRST}:$L${last},$B${FIRST}:$B${last},$V{i})",
            f"=W{i}-X{i}",
            f'=IF(ABS(Y{i})<10,"Matched for the year",IF(Y{i}>0,"ISD in 2B not claimed in 3B - "&TEXT(Y{i},"#,##0"),"Claimed in 3B more than 2B ISD - "&TEXT(-Y{i},"#,##0")))',
        ])
    slast = FIRST + len(gstins) - 1
    ws.Range(ws.Cells(FIRST, SUM_COL), ws.Cells(slast, SUM_COL + len(SUM_HEADERS) - 1)).Formula = srow
    for j in range(SUM_COL + 2, SUM_COL + 5):
        ws.Range(ws.Cells(FIRST, j), ws.Cells(slast, j)).NumberFormat = MONEY
        c = ws.Cells(HDR - 1, j)
        c.Formula = f"=SUM({L(j)}{FIRST}:{L(j)}{slast})"
        c.NumberFormat = MONEY
        c.Font.Bold = True

    widths = {1: 18, 2: 18, 3: 9, 4: 10, 17: 16, 18: 52, 19: 30, 21: 18, 22: 18, 26: 40}
    for j in range(1, SUM_COL + len(SUM_HEADERS)):
        ws.Columns(j).ColumnWidth = widths.get(j, 15)
    ws.Range(f"A{HDR}:{L(len(HEADERS))}{last}").AutoFilter()
    ws.Activate()
    xl.ActiveWindow.FreezePanes = False
    xl.ActiveWindow.SplitRow = 0
    xl.ActiveWindow.SplitColumn = 0
    ws.Range(f"A{FIRST}").Select()
    xl.ActiveWindow.FreezePanes = True

    xl.Calculate()
    err = errors_in(ws)
    h4, l4, p4 = ws.Range("H4").Value, ws.Range("L4").Value, ws.Range("P4").Value
    w4, x4 = ws.Range("W4").Value, ws.Range("X4").Value
    print(f"built {n} lines + {len(gstins)} GSTIN lines in {time.time()-t0:.0f}s; error cells {err}")
    print(f"2B ISD total {h4:,.2f} | 3B 4A(4) total {l4:,.2f} | diff {p4:,.2f} | per-GSTIN block {w4:,.2f} / {x4:,.2f}")
    vals = ws.Range(ws.Cells(FIRST, 1), ws.Cells(last, 18)).Value
    rem = {}
    for r in vals:
        key = str(r[17]).split(" - ")[0]
        rem[key] = rem.get(key, 0) + 1
    print("remarks:", rem)
    big = sorted(((r[0], r[2], r[15], r[16]) for r in vals if abs(r[15] or 0) >= 10), key=lambda t: -abs(t[2]))[:12]
    print("largest month gaps (state, month, diff, cumulative):", [(a, b, round(c, 2), round(d, 2)) for a, b, c, d in big])
    gv = ws.Range(ws.Cells(FIRST, SUM_COL), ws.Cells(slast, SUM_COL + 5)).Value
    print("per-GSTIN with a year gap:", [(r[0], round(r[4], 2)) for r in gv if abs(r[4] or 0) >= 10])
    assert err == 0 and abs(h4 - w4) < 0.01 and abs(l4 - x4) < 0.01
    wb.Worksheets("ITC Register 2025-26").Activate()
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
    print("fresh reopen OK; sheets:", wb2.Worksheets.Count, "| new sheet H4 =", wb2.Worksheets(NAME).Range("H4").Value)
    wb2.Close(False)
finally:
    xl2.Quit()
