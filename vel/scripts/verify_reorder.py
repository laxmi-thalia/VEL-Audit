import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
fails=[]
def ck(n,g,e,t=0.05):
    ok = g is not None and abs(g-e)<=t
    if not ok: fails.append(n)
    print("  %-46s %18s %s"%(n,format(g,",.2f") if isinstance(g,(int,float)) else g,"OK" if ok else "*** FAIL exp %s"%e))
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot); fails.extend(["errs"] if tot else [])
    mm=wb.Worksheets("S2 Month-on-Month")
    ck("S2 MoM residual sum (V)",xl.WorksheetFunction.Sum(mm.Range("V4:V1371")),-9.42,0.5)
    ck("S2 MoM SR taxable sum (E)",xl.WorksheetFunction.Sum(mm.Range("E4:E1371")),8721471166.37)
    s2=wb.Worksheets("S2 Reco (CA format)")
    ck("S2 diff AA74",s2.Range("AA74").Value,16876675.46)
    ck("S2 AH74 residual",s2.Range("AH74").Value,-9.42,0.5)
    m3=wb.Worksheets("S3 Month-on-Month")
    ck("S3 MoM diff taxable total (N232)",m3.Range("N232").Value,34765815.57)
    s31=wb.Worksheets("S3 1 vs 3B")
    ck("S3 1vs3B G1 taxable (C23)",s31.Range("C23").Value,8704594490.91)
    ck("S3 1vs3B 3B taxable (H23)",s31.Range("H23").Value,8669828675.34)
    ck("S3 1vs3B diff taxable (M23)",s31.Range("M23").Value,34765815.57)
    sv=wb.Worksheets("S3 SR vs 3B")
    ck("S3 SRvs3B SR taxable (C23)",sv.Range("C23").Value,8721471166.37)
    ck("S3 SRvs3B diff taxable (M23)",sv.Range("M23").Value,51642491.03)
    sm=wb.Worksheets("S3 SR vs 3B MoM")
    ck("S3 SRvs3B MoM diff total (N232)",sm.Range("N232").Value,51642491.03)
    s4=wb.Worksheets("S4 GL vs SR")
    ck("S4 GL CGST total (D23)",s4.Range("D23").Value,610374814.40)
    ck("S4 SR CGST total (H23)",s4.Range("H23").Value,609904424.06)
    ck("S4 diff CGST (L23)",s4.Range("L23").Value,470390.34)
    ck("S4 residual (P23)",s4.Range("P23").Value,-0.66,0.2)
    ck("S4 diff IGST (K23)",s4.Range("K23").Value,0.0)
    pv=wb.Worksheets("S2 Pivot Month-on-Month")
    ck("pivot alive",float(pv.PivotTables().Count),1,0)
    ck("register R2",wb.Worksheets("SR_2025-26").Range("R2").Value,8721471166.37)
    ck("S6 closing",wb.Worksheets("S6 Advances Control").Cells(23,6).Value,433267107.25)
    s7=wb.Worksheets("S7 CN Time-bar"); l7=s7.Cells(s7.Rows.Count,2).End(-4162).Row
    beyond=sum(1 for r in range(4,l7+1) if "BEYOND" in str(s7.Cells(r,8).Value or ""))
    ck("S7 beyond",float(beyond),22,0)
    n=sum(1 for w in wb.Worksheets if w.Name!="INDEX" for s in w.Shapes if s.Name=="btnIndex")
    ck("buttons",float(n),31,0)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if not fails else "FAIL: %s"%fails)
finally:
    xl.Quit()
