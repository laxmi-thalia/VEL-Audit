"""For every HYPERLINK on INDEX: parse the '#'-target and resolve it via COM exactly
the way Excel's click handler does (Application.Range on the sub-address)."""
import re, win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False

try:
    wb=xl.Workbooks.Open(P)
    xl.Calculation = -4135
    ix=wb.Worksheets("INDEX")
    ok=bad=0
    for r in range(1,60):
        f=ix.Cells(r,2).Formula
        if isinstance(f,str) and f.startswith("=HYPERLINK"):
            m=re.match(r'=HYPERLINK\("#(.+?)","', f)
            if not m: bad+=1; print("UNPARSEABLE row",r,f[:60]); continue
            target=m.group(1)
            try:
                rng=xl.Application.Range(target)   # what the click resolves
                sheet=rng.Worksheet.Name
                ok+=1
            except Exception as e:
                bad+=1; print("BROKEN row",r,"target",target,"->",e)
    print("targets resolved OK:",ok,"| broken:",bad)
    wb.Close(SaveChanges=False)
    print("VERDICT:","ALL 29 LINKS NAVIGATE" if ok==29 and bad==0 else "REVIEW")
finally:
    xl.Quit()
