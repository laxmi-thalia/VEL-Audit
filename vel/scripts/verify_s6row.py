import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P)
    a6=wb.Worksheets("S6 Advances Control"); l6=a6.Cells(a6.Rows.Count,1).End(-4162).Row
    print("S6 totals row:", l6, "| closing:", format(a6.Cells(l6,6).Value or 0,",.2f"))
    wb.Close(SaveChanges=False)
finally:
    xl.Quit()
