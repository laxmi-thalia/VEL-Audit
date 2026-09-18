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
    print("TOTALS: received %.2f | adjusted %.2f | books net %.2f | GSTR-1 net %.2f | diff %.2f"
          % (ca.Cells(tr,4).Value, ca.Cells(tr,5).Value, ca.Cells(tr,8).Value, ca.Cells(tr,7).Value, ca.Cells(tr,9).Value))
    print("states with books-vs-GSTR-1 diff >= 1 or negative closing:")
    for r in range(5,tr):
        d=ca.Cells(r,9).Value or 0; cl=ca.Cells(r,6).Value or 0
        if abs(d)>=1 or cl<-1:
            print("  %-18s recd %14.2f adj %14.2f closing %14.2f | g1net %14.2f diff %10.2f"
                  % (ca.Cells(r,1).Value, ca.Cells(r,4).Value or 0, ca.Cells(r,5).Value or 0, cl, ca.Cells(r,7).Value or 0, d))
    mm=wb.Worksheets("Month-on-Month")
    last=mm.Cells(mm.Rows.Count,1).End(-4162).Row
    n=sum(1 for r in range(4,last+1) if abs(mm.Cells(r,8).Value or 0)>=1 or abs(mm.Cells(r,9).Value or 0)>=1)
    print("month rows with a books-vs-GSTR-1 diff >= 1:", n)
    wb.Close(SaveChanges=True)
finally: xl.Quit()
