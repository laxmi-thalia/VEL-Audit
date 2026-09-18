import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_Step4_Reco_GL_vs_SR_DRAFT.xlsx"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot = 0
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells:", tot)
    ws = wb.Worksheets("GL vs Sales Register")
    tr = 23
    print("TOTALS: GL CGST %.2f | SR CGST %.2f | diff %.2f" % (ws.Cells(tr, 3).Value, ws.Cells(tr, 7).Value, ws.Cells(tr, 11).Value))
    print("        GL IGST %.2f | SR IGST %.2f | diff %.2f" % (ws.Cells(tr, 5).Value, ws.Cells(tr, 9).Value, ws.Cells(tr, 13).Value))
    print("        EXPLAINED %.2f | RESIDUAL %.2f" % (ws.Cells(tr, 15).Value or 0, ws.Cells(tr, 16).Value or 0))
    print("\nstates with CGST diff or residual >= 1:")
    for r in range(6, tr):
        k, res = ws.Cells(r, 11).Value or 0, ws.Cells(r, 16).Value or 0
        if abs(k) >= 1 or abs(res) >= 1:
            print("  %-18s diff %12.2f | explained %12.2f | RESIDUAL %10.2f"
                  % (ws.Cells(r, 2).Value, k, ws.Cells(r, 15).Value or 0, res))
    su = wb.Worksheets("Summary")
    for r in (11, 12, 13, 14): print("  Summary:", su.Cells(r, 1).Value, "=", su.Cells(r, 2).Value)
    ex = wb.Worksheets("Exceptions")
    print("\nExceptions sample:")
    for r in range(2, min(12, ex.Cells(ex.Rows.Count, 1).End(-4162).Row + 1)):
        print("  ", ex.Cells(r, 1).Value, "|", ex.Cells(r, 2).Value, "|", ex.Cells(r, 3).Value, "|", ex.Cells(r, 7).Value, "|", ex.Cells(r, 8).Value)
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
