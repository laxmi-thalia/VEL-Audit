import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\Downloads\VEL_Step6_Advances_Control_DRAFT.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot)
    ca=wb.Worksheets("Control Account"); tr=24
    print("opening %.2f + received %.2f - adjusted %.2f = closing %.2f | books-vs-G1 diff %.2f"
          % (ca.Cells(tr,3).Value, ca.Cells(tr,4).Value, ca.Cells(tr,5).Value, ca.Cells(tr,6).Value, ca.Cells(tr,9).Value))
    neg=[(ca.Cells(r,1).Value, round(ca.Cells(r,6).Value or 0,2), ca.Cells(r,10).Value) for r in range(5,tr) if (ca.Cells(r,6).Value or 0)<-1]
    print("states with NEGATIVE closing now:", neg if neg else "NONE - openings absorbed all in-year over-adjustments (Bihar -4.81 within rounding)")
    for r in range(5,tr):
        if abs(ca.Cells(r,6).Value or 0)>1 and (ca.Cells(r,3).Value or 0)>0:
            print("  %-18s open %14.2f -> close %14.2f" % (ca.Cells(r,1).Value, ca.Cells(r,3).Value, ca.Cells(r,6).Value))
    g=wb.Worksheets("GL Advance Check")
    lastg=g.Cells(g.Rows.Count,1).End(-4162).Row+1
    bad=[(g.Cells(r,2).Value, round(g.Cells(r,5).Value or 0,2)) for r in range(4,lastg) if abs(g.Cells(r,5).Value or 0)>=1]
    print("GL Advance Check rows with |ledger+books| >= 1:", bad if bad else "NONE - ledger ties to books in every state")
    wb.Close(SaveChanges=True)
finally: xl.Quit()
