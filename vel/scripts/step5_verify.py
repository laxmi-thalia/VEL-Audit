import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_Step5_HSN_Rate_Summary_DRAFT.xlsx"
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
    hs = wb.Worksheets("HSN Summary")
    last = hs.Cells(hs.Rows.Count, 1).End(-4162).Row
    print("HSN Summary totals: qty %.2f | taxable %.2f | IGST %.2f | CGST %.2f | SGST %.2f"
          % tuple(hs.Cells(last, j).Value or 0 for j in range(7, 12)))
    rw = wb.Worksheets("Rate-wise Summary")
    lr = rw.Cells(rw.Rows.Count, 1).End(-4162).Row
    print("Rate-wise totals:   taxable %.2f | total tax %.2f" % (rw.Cells(lr, 4).Value or 0, rw.Cells(lr, 8).Value or 0))
    print("EXPECTED: taxable 848,610,5710.82? no -> register ex-advances = 8486105710.82 | qty 53,395,416.32 | CGST 588,721,533.06")
    ck = wb.Worksheets("Checks")
    for r in range(3, 9):
        if ck.Cells(r, 1).Value: print("  Check:", ck.Cells(r, 1).Value, "=", ck.Cells(r, 2).Value)
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
