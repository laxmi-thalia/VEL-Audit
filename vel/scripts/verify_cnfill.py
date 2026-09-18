import win32com.client as win32, pythoncom
from collections import Counter
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
    s7=wb.Worksheets("S7 CN Time-bar"); l7=s7.Cells(s7.Rows.Count,2).End(-4162).Row
    cnt=Counter(); beyond=[]
    for r in range(4,l7+1):
        v=str(s7.Cells(r,8).Value or "")
        if not v: continue
        key=("BEYOND" if "BEYOND" in v else "not captured" if "not captured" in v else v[:40])
        cnt[key]+=1
        if "BEYOND" in v: beyond.append((str(s7.Cells(r,1).Value),str(s7.Cells(r,2).Value),str(s7.Cells(r,6).Value)[:11]))
    print("S7 statuses:",dict(cnt))
    print("BEYOND CNs:")
    for b in beyond: print("  ",b)
    print("register R2:",format(wb.Worksheets("SR_2025-26").Range("R2").Value,",.2f"))
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
