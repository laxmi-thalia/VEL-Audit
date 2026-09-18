import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
fails=[]
def ck(name,got,exp,tol=0.05):
    ok = got is not None and abs(got-exp)<=tol
    if not ok: fails.append("%s: %r vs %r"%(name,got,exp))
    print("  %-46s %18s %s"%(name,format(got,",.2f") if isinstance(got,(int,float)) else got,"OK" if ok else "*** FAIL exp %s"%exp))
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot); fails.extend(["errors %d"%tot] if tot else [])
    ws=wb.Worksheets("SR_2025-26"); ck("register taxable R2",ws.Range("R2").Value,8721471166.37)
    s2=wb.Worksheets("S2 Reco (CA format)"); ck("S2 diff",s2.Range("AA74").Value,16876675.46)
    s3=wb.Worksheets("S3 1 vs 3B"); last=s3.Cells(s3.Rows.Count,2).End(-4162).Row
    ck("S3 diff",s3.Cells(last,5).Value,34765815.57)
    s4=wb.Worksheets("S4 GL vs SR"); l4=s4.Cells(s4.Rows.Count,2).End(-4162).Row
    ck("S4 residual",s4.Cells(l4,10).Value,-0.66,0.2)
    # S4 Exceptions now formulas - values must be unchanged
    ex=wb.Worksheets("S4 Exceptions")
    exp_h={2:365858,3:43888,4:20076,5:-5837,6:634499.7,7:46406,8:67548.81,9:-634499.7,10:-67548.81}
    for r,v in exp_h.items(): ck("S4 Exc row %d contribution"%r,ex.Cells(r,8).Value,v,0.02)
    # S9 advances line live
    s9=wb.Worksheets("S9 Sales Reco")
    advr=None
    for r in range(6,20):
        if str(s9.Cells(r,1).Value or "").startswith("Unadjusted advances"): advr=r; break
    hdr={str(s9.Cells(3,j).Value or ""):j for j in range(2,21)}
    ck("S9 adv Arunachal",s9.Cells(advr,hdr["Arunachal Pradesh"]).Value,-3807022.22)
    ck("S9 adv Bihar",s9.Cells(advr,hdr["Bihar"]).Value,-729166.67)
    ck("S9 adv Gujarat",s9.Cells(advr,hdr["Gujarat"]).Value,-51069933.29)
    ck("S9 adv MP",s9.Cells(advr,hdr["Madhya Pradesh"]).Value,81524711.12)
    ck("S9 adv total (=control acct net)",s9.Cells(advr,21).Value,235365455.60,0.2)
    for r in range(6,20):
        if str(s9.Cells(r,1).Value or "")=="Difference (A)":
            ck("S9 Difference (A) total",s9.Cells(r,21).Value,3771605598.21,1); break
    a6=wb.Worksheets("S6 Advances Control"); l6=a6.Cells(a6.Rows.Count,1).End(-4162).Row
    ck("S6 closing",a6.Cells(l6,6).Value,433267107.29)
    s7=wb.Worksheets("S7 CN Time-bar"); l7=s7.Cells(s7.Rows.Count,2).End(-4162).Row
    beyond=sum(1 for r in range(4,l7+1) if "BEYOND" in str(s7.Cells(r,8).Value or ""))
    ck("S7 beyond-window CNs",float(beyond),3,0)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","ALL CHECKS PASS" if not fails else "FAILURES:\n"+"\n".join(fails))
finally:
    xl.Quit()
