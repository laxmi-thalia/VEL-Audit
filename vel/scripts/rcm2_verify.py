import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
fails=[]
def ck(n,g,e,t=0.05):
    ok = g is not None and abs(g-e)<=t
    if not ok: fails.append(n)
    print("  %-42s %18s %s"%(n,format(g,",.2f") if isinstance(g,(int,float)) else g,"OK" if ok else "*** FAIL exp %s"%e))
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot); fails.extend(["errs"] if tot else [])
    sw=wb.Worksheets("Statewise RCM vs 3B")
    lr=sw.Cells(sw.Rows.Count,1).End(-4162).Row
    ck("SW reg taxable GT",sw.Cells(lr,2).Value,115152683.11)
    ck("SW reg IGST GT",sw.Cells(lr,3).Value,8096738.00)
    ck("SW 3B taxable GT",sw.Cells(lr,8).Value,125601384.80)
    ck("SW 3B IGST GT",sw.Cells(lr,9).Value,8692445.00)
    ck("SW diff taxable GT",sw.Cells(lr,14).Value,-10448701.69,0.1)
    big=[]
    for r in range(6,lr):
        d=sw.Cells(r,14).Value or 0
        if abs(d)>=1000: big.append((str(sw.Cells(r,1).Value),round(d,2)))
    print("  states with |taxable diff| >= 1000:",len(big))
    for b in sorted(big,key=lambda x:-abs(x[1]))[:8]: print("    ",b[0],format(b[1],",.2f"))
    mw=wb.Worksheets("Month wise RCM vs 3B")
    lm=mw.Cells(mw.Rows.Count,1).End(-4162).Row
    ck("MW reg taxable GT",mw.Cells(lm,2).Value,115152683.11)
    ck("MW 3B taxable GT",mw.Cells(lm,7).Value,125601384.80)
    ck("MW diff GT ties SW",mw.Cells(lm,13).Value,(sw.Cells(lr,14).Value or 0),0.1)
    ck("sales golden T3",wb.Worksheets("SR_2025-26").Range("T3").Value,8721471166.37)
    ck("RCM reg SAP subtotal intact",wb.Worksheets("RCM Register").Cells(4,44).Value,115152683.11)
    n=sum(1 for w in wb.Worksheets if w.Name!="INDEX" for s in w.Shapes if s.Name=="btnIndex")
    print("  buttons:",n,"| sheets:",wb.Worksheets.Count)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if not fails else "FAIL: %s"%fails)
finally:
    xl.Quit()
