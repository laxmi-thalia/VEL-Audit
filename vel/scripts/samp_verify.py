import win32com.client as win32, pythoncom, collections
P=r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    ws=wb.Worksheets("SR_2025-26")
    try: errs=ws.Range("A5:BC27006").SpecialCells(-4123,16).Count
    except Exception: errs=0
    hdr=[ws.Cells(4,c).Value for c in range(1,56)]
    print("formula ERROR cells:", errs)
    print("AX..BC headers:", hdr[49:55])
    print("R2 taxable:", ws.Range("R2").Value)
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
