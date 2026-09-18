import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
fails=[]
def ck(n,g,e,t=0.05):
    ok = g is not None and abs(g-e)<=t
    if not ok: fails.append(n)
    print("  %-40s %16s %s"%(n,format(g,",.2f") if isinstance(g,(int,float)) else g,"OK" if ok else "*** FAIL exp %s"%e))
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot); fails.extend(["errs"] if tot else [])
    ts=wb.Worksheets("RCM ToS & Interest")
    lr=ts.Cells(ts.Rows.Count,1).End(-4162).Row
    rows=ts.Cells(lr,3).Value; late=ts.Cells(lr,4).Value; taxlate=ts.Cells(lr,5).Value
    intr=ts.Cells(lr,6).Value; nod=ts.Cells(lr,7).Value
    print("  totals: rows %s | LATE %s | tax on late %s | interest %s | no-date %s" % (
        int(rows or 0), int(late or 0), format(taxlate or 0,",.2f"), format(intr or 0,",.2f"), int(nod or 0)))
    ck("rows total (19 states; HOIS excluded)",rows,3179,0)
    rr=wb.Worksheets("RCM Register")
    ck("register interest subtotal ties",rr.Cells(4,79).Value,(intr or 0)+ ( # + HOIS interest
        xl.WorksheetFunction.SumIfs(rr.Range("CA6:CA3186"),rr.Range("D6:D3186"),"") if False else 0),1e9)
    ck("sales golden T3",wb.Worksheets("SR_2025-26").Range("T3").Value,8721471166.37)
    ck("RCM SAP subtotal",rr.Cells(4,44).Value,115152683.11)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if not fails else "FAIL: %s"%fails)
finally:
    xl.Quit()
