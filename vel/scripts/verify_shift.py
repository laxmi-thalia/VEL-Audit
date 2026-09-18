# All row refs +1 after the reserved button row insert
import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
fails=[]
def ck(n,g,e,t=0.05):
    ok = g is not None and abs(g-e)<=t
    if not ok: fails.append(n)
    print("  %-44s %18s %s"%(n,format(g,",.2f") if isinstance(g,(int,float)) else g,"OK" if ok else "*** FAIL exp %s"%e))
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot); fails.extend(["errs"] if tot else [])
    ck("register taxable (T3)",wb.Worksheets("SR_2025-26").Range("T3").Value,8721471166.37)
    ck("gate blocking (Flags B5)",float(wb.Worksheets("Step 1 Flags").Range("B5").Value),3,0)
    s2=wb.Worksheets("S2 Reco (CA format)")
    ck("S2 diff (AA75)",s2.Range("AA75").Value,16876675.46)
    ck("S2 residual (AH75)",s2.Range("AH75").Value,-9.42,0.5)
    mm=wb.Worksheets("S2 Month-on-Month")
    ck("S2 MoM residual (V5:V1372)",xl.WorksheetFunction.Sum(mm.Range("V5:V1372")),-9.42,0.5)
    ck("S3 1vs3B diff (M24)",wb.Worksheets("S3 1 vs 3B").Range("M24").Value,34765815.57)
    ck("S3 SRvs3B diff (M24)",wb.Worksheets("S3 SR vs 3B").Range("M24").Value,51642491.03)
    ck("S3 MoM diff total (N233)",wb.Worksheets("S3 Month-on-Month").Range("N233").Value,34765815.57)
    ck("S3 SRvs3B MoM diff (N233)",wb.Worksheets("S3 SR vs 3B MoM").Range("N233").Value,51642491.03)
    s4=wb.Worksheets("S4 GL vs SR")
    ck("S4 GL CGST (D24)",s4.Range("D24").Value,610374814.40)
    ck("S4 residual (P24)",s4.Range("P24").Value,-0.66,0.2)
    ck("S5 HSN taxable",wb.Worksheets("S5 HSN Summary").Cells(wb.Worksheets("S5 HSN Summary").Cells(1048576,1).End(-4162).Row,7).Value,8486105710.82)
    a6=wb.Worksheets("S6 Advances Control")
    ck("S6 closing (F24)",a6.Range("F24").Value,433267107.25)
    m6=wb.Worksheets("S6 Month-on-Month")
    l6=m6.Cells(1048576,1).End(-4162).Row
    ck("S6 MoM closing total",m6.Cells(l6,7).Value,433267107.25)
    s7=wb.Worksheets("S7 CN Time-bar"); l7=s7.Cells(1048576,2).End(-4162).Row
    beyond=sum(1 for r in range(5,l7+1) if "BEYOND" in str(s7.Cells(r,8).Value or ""))
    ck("S7 beyond",float(beyond),22,0)
    s9=wb.Worksheets("S9 Sales Reco")
    ck("S9 Difference A (U23)",s9.Range("U23").Value,3771605598.21,1)
    pv=wb.Worksheets("S2 Pivot Month-on-Month")
    pt=pv.PivotTables("PT_Long"); pt.PivotCache().Refresh()
    ck("pivot diff taxable",pt.GetData("'Diff Taxable '"),16876675.46)
    n=sum(1 for w in wb.Worksheets if w.Name!="INDEX" for s in w.Shapes if s.Name=="btnIndex")
    ck("buttons (32 non-INDEX sheets)",float(n),32,0)
    ok=0
    for w in wb.Worksheets:
        if w.Name=="INDEX": continue
        for s in w.Shapes:
            if s.Name=="btnIndex" and s.Top<21 and s.Left<10: ok+=1
    ck("buttons inside reserved row",float(ok),32,0)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if not fails else "FAIL: %s"%fails)
finally:
    xl.Quit()
