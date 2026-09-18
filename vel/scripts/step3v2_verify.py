import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_Step3_Reco_GSTR1_vs_3B_DRAFT.xlsx"
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot = 0
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells (all sheets):", tot)
    ws = wb.Worksheets("GSTR-1 vs GSTR-3B")
    g1, b3, d = ws.Range("C50").Value, ws.Range("G50").Value, ws.Range("K50").Value
    print("Totals: GSTR-1 %.2f | 3B %.2f | diff %.2f   (expected 8,704,594,490.91 / 8,669,828,675.34 / 34,765,815.57)" % (g1, b3, d))
    print("States with |diff taxable| >= 1:")
    for r in range(31, 50):
        if abs(ws.Cells(r, 11).Value or 0) >= 1 or abs(ws.Cells(r, 13).Value or 0) >= 1:
            print("   %-18s taxable diff %16.2f | CGST diff %12.2f" % (ws.Cells(r, 2).Value, ws.Cells(r, 11).Value or 0, ws.Cells(r, 13).Value or 0))
    ex = wb.Worksheets("Exceptions")
    print("\nEvidence sheet (live values):")
    last = ex.Cells(ex.Rows.Count, 1).End(-4162).Row
    keys = ["Taxable value", "CGST", "9% check", "BLANK IRN and a value", "their taxable value", "found in the books",
            "zero taxable and blank IRN", "Bihar Aug-25", "Bihar Jan-26", "Telangana (3 months", "Sum of the above", "Annual GSTR-1", "Unexplained residual",
            "CGST difference May-25"]
    for r in range(1, last + 1):
        a = ex.Cells(r, 1).Value
        if isinstance(a, str) and any(k in a for k in keys):
            vals = [ex.Cells(r, c).Value for c in range(2, 7)]
            vals = [round(v, 4) if isinstance(v, float) else v for v in vals]
            print("   r%-3d %-60s %s" % (r, a[:60], vals))
    # month rows
    for mo_r in range(1, last + 1):
        a = ex.Cells(mo_r, 1).Value
        if a in ("May-25", "Jun-25", "Aug-25"):
            print("   Telangana %s: G1 %.2f | 3B %.2f | diff %.2f | CGST diff %.2f | ratio %s" % (a, ex.Cells(mo_r, 2).Value, ex.Cells(mo_r, 3).Value, ex.Cells(mo_r, 4).Value, ex.Cells(mo_r, 5).Value, ex.Cells(mo_r, 6).Value))
    # in-books checks must all be 0 ; IRN lengths 64 ; in-GSTR-1 checks 0
    zeros = ok64 = bad = 0
    for r in range(1, last + 1):
        h = ex.Cells(r, 8).Value
        if isinstance(h, (int, float)) and isinstance(ex.Cells(r, 1).Value, (int, float)):
            if h == 0: zeros += 1
            else: bad += 1
        g = ex.Cells(r, 7).Value
        if g == 64: ok64 += 1
    print("   listed documents: 'in books'/'in GSTR-1' checks = 0 on %d rows, non-zero on %d | IRN length 64 on %d Telangana rows" % (zeros, bad, ok64))
    am = wb.Worksheets("Amendment Check"); print("\nAmendment Check: FY25-26 amendment rows =", am.Range("B3").Value)
    su = wb.Worksheets("Summary")
    for r in range(3, 10):
        if su.Cells(r, 1).Value: print("   ", su.Cells(r, 1).Value, "=", su.Cells(r, 2).Value)
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
