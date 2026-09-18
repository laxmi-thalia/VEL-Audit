import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
fails=[]
def ck(n,g,e,t=0.05):
    ok = g is not None and abs(g-e)<=t
    if not ok: fails.append("%s: %r vs %r"%(n,g,e))
    print("  %-40s %18s %s"%(n,format(g,",.2f") if isinstance(g,(int,float)) else g,"OK" if ok else "*** FAIL exp %s"%e))
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot); fails.extend(["errors %d"%tot] if tot else [])
    m=wb.Worksheets("S3 SR vs 3B MoM")
    ck("MoM total diff taxable",m.Cells(232,6).Value,51642491.03)
    ck("state-level sheet still ties",wb.Worksheets("S3 SR vs 3B").Cells(23,5).Value,51642491.03)
    # rows with a material diff
    big=[(str(m.Cells(r,1).Value),str(m.Cells(r,3).Value),m.Cells(r,6).Value) for r in range(4,232) if abs(m.Cells(r,6).Value or 0)>=1]
    print("  state-months with |diff| >= Re 1:",len(big))
    for b in big[:12]: print("    ",b[0],b[1],format(b[2],",.2f"))
    ck("register untouched",wb.Worksheets("SR_2025-26").Range("R2").Value,8721471166.37)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if not fails else "FAIL:\n"+"\n".join(fails))
finally:
    xl.Quit()
