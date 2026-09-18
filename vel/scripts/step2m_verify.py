import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot = 0
    for n in ["SR_2025-26", "SR vs GSTR-1", "Month-on-Month 1", "Step 2 Summary", "GSTR-1 Data", "Step 1 Flags"]:
        try: e = wb.Worksheets(n).UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", n, e)
    print("formula ERROR cells (all sheets):", tot)
    ws = wb.Worksheets("SR vs GSTR-1")
    G = ["B2B", "B2C", "Credit note", "Debit note", "Advance received", "Advance adjusted", "Total liability"]
    EXP = {27: [8927118980.77, 64753.60, -486125232.19, 28170523.64, 381672455.66, -146306990.57, 8704594490.91],
           51: [8933996915.45, 64753.72, -476126481.99, 28170523.64, 381672455.56, -146307000.01, 8721471166.37],
           75: [6877934.68, 0.12, 9998750.20, 0.00, -0.10, -9.44, 16876675.46]}
    ok = True
    for row, exp in EXP.items():
        for i, (e, name) in enumerate(zip(exp, G)):
            got = round(ws.Cells(row, 3 + i * 4).Value or 0, 2)
            if abs(got - e) > 0.02: ok = False; print("  MISMATCH row", row, name, got, "vs", e)
    print("CA-format totals vs verified values:", "ALL MATCH" if ok else "*** MISMATCH ***")
    print("Difference rows with |taxable| >= 1:")
    for r in range(56, 75):
        nz = [(G[i], round(ws.Cells(r, 3 + i * 4).Value or 0, 2)) for i in range(7) if abs(ws.Cells(r, 3 + i * 4).Value or 0) >= 1]
        if nz: print("  ", ws.Cells(r, 2).Value, nz)
    mm = wb.Worksheets("Month-on-Month 1")
    last = mm.Cells(mm.Rows.Count, 1).End(-4162).Row
    diffs = [(mm.Cells(r, 1).Value, mm.Cells(r, 3).Value, round(mm.Cells(r, 12).Value, 2)) for r in range(4, last + 1)
             if mm.Cells(r, 3).Value and abs(mm.Cells(r, 12).Value or 0) >= 1]
    print("Month rows with taxable diff >= 1:", len(diffs)); [print("  ", d) for d in diffs]
    bt = sum((mm.Cells(r, 4).Value or 0) for r in range(4, last + 1) if mm.Cells(r, 3).Value)
    print("MoM books taxable total:", round(bt, 2), "(register 8,721,471,166.37)")
    su = wb.Worksheets("Step 2 Summary")
    for r in (2, 4, 5, 6, 8, 9, 13, 14, 15):
        print("  ", su.Cells(r, 1).Value, "=", su.Cells(r, 2).Value, ("| " + str(round(su.Cells(r, 3).Value, 2))) if isinstance(su.Cells(r, 3).Value, (int, float)) else "")
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
