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
    s2=wb.Worksheets("S2 Reco (CA format)")
    ck("S2 total diff (AA74)",s2.Range("AA74").Value,16876675.46)
    ck("S2 explained books-side (AF74)",s2.Range("AF74").Value,11430864.28,50)
    ck("S2 explained G1-side (AG74)",s2.Range("AG74").Value,5445820.62,50)
    ck("S2 unexplained residual (AH74)",s2.Range("AH74").Value,-9.42,0.5)
    print("  (step2 file golden: AF74+AG74=16,876,684.90, AH74=-9.42)")
    print("  AF+AG =",format((s2.Range("AF74").Value or 0)+(s2.Range("AG74").Value or 0),",.2f"))
    ck("register untouched",wb.Worksheets("SR_2025-26").Range("R2").Value,8721471166.37)
    ck("S6 closing untouched",wb.Worksheets("S6 Advances Control").Cells(24,6).Value,433267107.29)
    n=0
    for ws in wb.Worksheets:
        if ws.Name=="INDEX": continue
        for shp in ws.Shapes:
            if shp.Name=="btnIndex" and shp.Hyperlink.SubAddress=="'INDEX'!A1": n+=1
    ck("INDEX buttons",float(n),29,0)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if not fails else "FAIL: %s"%fails)
finally:
    xl.Quit()
