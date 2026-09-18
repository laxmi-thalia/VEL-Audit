import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
fails=[]
def ck(n,g,e,t=0.6):
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
    s9=wb.Worksheets("S9 Sales Reco")
    ck("U8  Revenue from Operations (FS)",s9.Cells(8,21).Value,12493076764.58)
    ck("U20 Total sales as per GST",s9.Cells(20,21).Value,8721471166.37)
    ck("U22 Difference (A)",s9.Cells(22,21).Value,3771605598.21)
    ck("U34 advance opening (-)",s9.Cells(34,21).Value,-197901651.69)
    ck("U35 advance closing (+)",s9.Cells(35,21).Value,433267107.25)
    ck("U38 section-2 total",s9.Cells(38,21).Value,235365455.56)
    ck("U13 Sales (INV)",s9.Cells(13,21).Value,8721471166.37-235365455.56-( (s9.Cells(14,21).Value or 0)+(s9.Cells(16,21).Value or 0) ),2)
    # cross-checks with other sheets
    ck("S6 closing tie",wb.Worksheets("S6 Advances Control").Cells(23,6).Value,433267107.25)
    ck("register tie",wb.Worksheets("SR_2025-26").Range("R2").Value,8721471166.37)
    n=0
    for ws in wb.Worksheets:
        if ws.Name=="INDEX": continue
        for shp in ws.Shapes:
            if shp.Name=="btnIndex" and shp.Hyperlink.SubAddress=="'INDEX'!A1": n+=1
    ck("INDEX buttons",float(n),30,0)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if not fails else "FAIL: %s"%fails)
finally:
    xl.Quit()
