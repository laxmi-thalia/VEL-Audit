import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
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
    ix=wb.Worksheets("INDEX")
    links=bad=0
    for r in range(1,60):
        for c in (2,):
            v=ix.Cells(r,c).Formula
            if isinstance(v,str) and v.startswith("=HYPERLINK"):
                links+=1
                t=str(ix.Cells(r,c).Value or "")
                if t.startswith("#") or not t: bad+=1
    print("hyperlinks:",links,"| broken:",bad)
    print("register:",format(wb.Worksheets("SR_2025-26").Range("R2").Value,",.2f"))
    print("sheet1 is:",wb.Worksheets(1).Name)
    wb.Close(SaveChanges=True)
    print("VERDICT:","PASS" if tot==0 and bad==0 and links==29 else "REVIEW (links=%d)"%links)
finally:
    xl.Quit()
