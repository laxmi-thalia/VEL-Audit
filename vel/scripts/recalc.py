import win32com.client as win32, pythoncom, os
P=r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P)
    xl.CalculateFullRebuild()
    ws=wb.Worksheets("SR_2025-26")
    errs=0; samples=[]
    try:
        rng=ws.Range("A5:AW27006")
        f=rng.SpecialCells(-4123, 16)   # xlCellTypeFormulas, xlErrors
        errs=f.Count
        for c in list(f)[:5]: samples.append((c.Address, str(c.Text)))
    except Exception:
        errs=0
    print("formula ERROR cells:", errs, samples)
    for a in ["Q2","R2","S2","T2","U2","V2","AI2"]:
        print(f"  {a} (SUBTOTAL) = {ws.Range(a).Value}")
    for a in ["P5","V5","W5","Y5","AF5","AN5","AO5","AP5","AQ5","AV5"]:
        print(f"  {a} = {ws.Range(a).Value}")
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
