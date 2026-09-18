import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\VEL_2425.xlsb"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P, ReadOnly=True, UpdateLinks=0)
    print(f"sheets: {wb.Sheets.Count}")
    for i,w in enumerate(wb.Sheets,1):
        try:
            ur=w.UsedRange; dims=f"{ur.Rows.Count}x{ur.Columns.Count}"
        except Exception: dims="?"
        print(f"  {i:2d}. {w.Name!r:52s} {dims:>14s}  visible={w.Visible}")
finally:
    if wb is not None: wb.Close(SaveChanges=False)
    xl.Quit()
