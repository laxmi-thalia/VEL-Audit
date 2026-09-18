import win32com.client as win32, os, pythoncom
src = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\july_src.xls"
dst = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\July2025_converted.xlsx"
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb = xl.Workbooks.Open(src, ReadOnly=True, UpdateLinks=0)
    print("sheets:", [s.Name for s in wb.Sheets])
    if os.path.exists(dst): os.remove(dst)
    wb.SaveAs(dst, FileFormat=51); wb.Close(SaveChanges=False)
    print("WROTE", os.path.getsize(dst), "bytes")
finally:
    xl.Quit()
