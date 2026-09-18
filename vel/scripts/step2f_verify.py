import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\VEL_Step2_FORMULA_BUILD.xlsx"
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(P)
    xl.CalculateFullRebuild()
    total_err = 0
    for shn in ["SR vs GSTR-1 (CA format)", "Month-on-Month", "Reconciliation Summary", "SR Data", "GSTR-1 Data", "Matched"]:
        ws = wb.Worksheets(shn)
        try:
            e = ws.UsedRange.SpecialCells(-4123, 16).Count
        except Exception:
            e = 0
        total_err += e
        if e: print("ERROR CELLS in", shn, ":", e)
    print("formula ERROR cells across all sheets:", total_err)

    ws = wb.Worksheets("SR vs GSTR-1 (CA format)")
    G = ["B2B", "B2C", "Credit note", "Debit note", "Advance received", "Advance adjusted", "Total liability"]
    # expected totals from the previous (value-based, verified) build
    EXP = {
        26: [8927118980.77, 64753.60, -486125232.19, 28170523.64, 381672455.66, -146306990.57, 8704594490.91],
        50: [8933996915.45, 64753.72, -476126481.99, 28170523.64, 381672455.56, -146307000.01, 8721471166.37],
        74: [6877934.68, 0.12, 9998750.20, 0.00, -0.10, -9.44, 16876675.46],
    }
    ok = True
    for row, exp in EXP.items():
        got = [round(ws.Cells(row, 3 + i * 4).Value or 0, 2) for i in range(7)]
        for g, e, name in zip(got, exp, G):
            if abs(g - e) > 0.02:
                ok = False
                print("  MISMATCH row", row, name, ": got", g, "expected", e)
    print("block totals vs verified values:", "ALL MATCH" if ok else "*** MISMATCH ***")

    # nonzero difference rows
    print("\nDifference rows (taxable) with |value| >= 1:")
    for r in range(55, 74):
        st = ws.Cells(r, 2).Value
        nz = [(G[i], round(ws.Cells(r, 3 + i * 4).Value or 0, 2)) for i in range(7)
              if abs(ws.Cells(r, 3 + i * 4).Value or 0) >= 1]
        if nz:
            print("  ", st, nz)

    mm = wb.Worksheets("Month-on-Month")
    last = mm.Cells(mm.Rows.Count, 1).End(-4162).Row
    n = 0
    for r in range(2, last + 1):
        if abs(mm.Cells(r, 7).Value or 0) >= 1:
            n += 1
    print("\nMonth-on-Month rows:", last - 1, "| with taxable diff >= Rs 1:", n, "(expected 7)")

    su = wb.Worksheets("Reconciliation Summary")
    print("\nSummary (live):")
    for r in range(3, 13):
        a = su.Cells(r, 1).Value
        if a:
            print("  ", a, "=", su.Cells(r, 2).Value,
                  ("| " + str(round(su.Cells(r, 3).Value, 2))) if su.Cells(r, 3).Value is not None else "")
    wb.Close(SaveChanges=True)
finally:
    if wb is None:
        pass
    xl.Quit()
