import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    ws=wb.Worksheets("SR_2025-26")
    try: e=ws.Range("A5:BD27006").SpecialCells(-4123,16).Count
    except Exception: e=0
    print("register formula errors:", e, "| sheets:", [s.Name for s in wb.Sheets])
    g=wb.Worksheets("Step 1 Flags")
    print("Blocking:", g.Range("B4").Value, "| Review:", g.Range("B5").Value, "| GATE:", g.Range("B7").Value)
    wb.Close(SaveChanges=True)
finally: xl.Quit()
