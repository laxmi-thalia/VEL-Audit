import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot = 0
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells:", tot, "| visible sheets:", [w.Name for w in wb.Worksheets if w.Visible == -1])
    ws = wb.Worksheets("SR vs GSTR-1")
    print("\nDifference block with auto-explanation (rows with any nonzero):")
    for r in range(55, 74):
        d = ws.Cells(r, 27).Value or 0   # AA = Total liability taxable diff
        af, ag, ah = ws.Cells(r, 32).Value or 0, ws.Cells(r, 33).Value or 0, ws.Cells(r, 34).Value or 0
        if abs(d) >= 1 or abs(ah) >= 1:
            print("  %-18s diff %15.2f | books-only %14.2f | g1-only %14.2f | RESIDUAL %10.2f | %s"
                  % (ws.Cells(r, 2).Value, d, af, ag, ah, str(ws.Cells(r, 35).Value or "")[:45]))
    print("  TOTAL row 74: residual =", round(ws.Cells(74, 34).Value or 0, 2))
    su = wb.Worksheets("Summary")
    for r in range(7, 12):
        a = su.Cells(r, 1).Value
        if a: print("  Summary:", a, "=", su.Cells(r, 2).Value, "|", su.Cells(r, 3).Value)
    ex = wb.Worksheets("Exceptions")
    import collections
    c = collections.Counter()
    for r in range(2, ex.Cells(ex.Rows.Count, 1).End(-4162).Row + 1): c[ex.Cells(r, 1).Value] += 1
    print("  Exceptions statuses:", dict(c))
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
