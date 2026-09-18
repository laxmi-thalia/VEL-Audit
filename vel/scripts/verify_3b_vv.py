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
    b3=wb.Worksheets("3B Data")
    cnt=Counter(); flagged=[]
    for r in range(2,230):
        s=str(b3.Cells(r,10).Value or ""); cnt[s]+=1
        if s and s!="Matched with books":
            flagged.append((str(b3.Cells(r,1).Value),str(b3.Cells(r,3).Value),s,b3.Cells(r,9).Value))
    for k,v in cnt.most_common(): print(" ",k,"->",v)
    for f_ in flagged: print("  FLAG:",f_[0],f_[1],"|",f_[2],"| diff",format(f_[3] or 0,",.2f"))
    ok = tot==0 and cnt["Matched with books"]==226
    print("register check:",format(wb.Worksheets("SR_2025-26").Range("R2").Value,",.2f"))
    wb.Close(SaveChanges=True)
    print("\nVERDICT:","PASS" if ok else "REVIEW")
finally:
    xl.Quit()
