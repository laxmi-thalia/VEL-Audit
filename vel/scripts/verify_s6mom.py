import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
fails=[]
def ck(n,g,e,t=0.05):
    ok = g is not None and abs(g-e)<=t
    if not ok: fails.append(n)
    print("  %-40s %18s %s"%(n,format(g,",.2f") if isinstance(g,(int,float)) else g,"OK" if ok else "*** FAIL exp %s"%e))
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot); fails.extend(["errs"] if tot else [])
    m=wb.Worksheets("S6 Month-on-Month")
    ck("MoM opening total",m.Cells(232,4).Value,197901651.69)
    ck("MoM received total",m.Cells(232,5).Value,381672455.56)
    ck("MoM adjusted total",m.Cells(232,6).Value,146307000.00)
    ck("MoM closing total (Mar-26)",m.Cells(232,7).Value,433267107.25)
    ck("MoM books-vs-G1 total",m.Cells(232,10).Value,-9.53,0.2)
    a6=wb.Worksheets("S6 Advances Control")
    ck("annual closing agrees",a6.Cells(23,6).Value,433267107.25)
    # every state's Mar-26 closing equals its annual closing
    ann={}
    for r in range(4,23): ann[str(a6.Cells(r,2).Value)]=a6.Cells(r,6).Value or 0
    bad=0
    for r in range(4,232):
        if str(m.Cells(r,3).Value)=="Mar-26":
            g=str(m.Cells(r,2).Value)
            if abs((m.Cells(r,7).Value or 0)-ann.get(g,0))>0.05: bad+=1; print("  state closing mismatch:",g)
    ck("state closings tie annual",float(bad),0,0)
    n=0
    for ws in wb.Worksheets:
        if ws.Name=="INDEX": continue
        for shp in ws.Shapes:
            if shp.Name=="btnIndex" and shp.Hyperlink.SubAddress=="'INDEX'!A1": n+=1
    ck("INDEX buttons (now 30 sheets)",float(n),30,0)
    ix=wb.Worksheets("INDEX")
    links=sum(1 for r in range(1,70) if str(ix.Cells(r,2).Formula or "").startswith("=HYPERLINK"))
    ck("INDEX links",float(links),30,0)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if not fails else "FAIL: %s"%fails)
finally:
    xl.Quit()
