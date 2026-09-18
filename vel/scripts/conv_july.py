import win32com.client as win32, os, pythoncom
src = r"\server\GST FOLDER\GST Returns\GST Audit & Annual Return\FY 2025-26\1. Corporate Clients\VEL\Clients Data\05.08.2026 Main Data\04 July 2025\Cleartax Sales Reg July 2025.xls"
dst = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\July2025_converted.xlsx"
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application")
xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(src, ReadOnly=True, UpdateLinks=0)
    print("sheets:", [s.Name for s in wb.Sheets])
    if os.path.exists(dst): os.remove(dst)
    wb.SaveAs(dst, FileFormat=51)   # xlOpenXMLWorkbook
    wb.Close(SaveChanges=False)
    print("WROTE:", dst, os.path.getsize(dst), "bytes")
finally:
    xl.Quit()
