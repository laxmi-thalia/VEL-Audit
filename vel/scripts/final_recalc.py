import win32com.client as win32, pythoncom, collections
P=r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    ws=wb.Worksheets("SR_2025-26")
    try:
        f=ws.Range("A5:AW27006").SpecialCells(-4123,16); errs=f.Count
    except Exception: errs=0
    print("formula ERROR cells:", errs)
    vals=ws.Range("AF5:AF27006").Value
    c=collections.Counter(v[0] for v in vals)
    print("GOOD/SERVICE split:", dict(c))
    print("sample AG (HSN) types:", [type(ws.Range(f"AG{r}").Value).__name__ for r in (5,6,7,100)])
    print("sample AG values     :", [ws.Range(f"AG{r}").Value for r in (5,6,7,100)])
    print("Q2 SUBTOTAL taxable  :", ws.Range("Q2").Value)
    wb.Close(SaveChanges=True)
finally:
    if wb is None: pass
    xl.Quit()
