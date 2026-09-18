import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\Downloads\VEL_Step3_Reco_GSTR1_vs_3B_DRAFT.xlsx"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    ws=wb.Worksheets("GSTR-1 vs GSTR-3B")
    try: errs=ws.Range("A1:O50").SpecialCells(-4123,16).Count
    except Exception: errs=0
    print("formula ERROR cells:", errs)
    print(f"\n{'state':22s} {'GSTR-1':>18s} {'GSTR-3B':>18s} {'diff taxable':>16s} {'diff CGST':>14s}")
    for r in range(31,51):
        st=ws.Cells(r,2).Value or "TOTAL"
        a=ws.Cells(r,3).Value or 0; b=ws.Cells(r,7).Value or 0
        d=ws.Cells(r,11).Value or 0; dc=ws.Cells(r,13).Value or 0
        flag="   <<<" if abs(d)>=1 or abs(dc)>=1 else ""
        print(f"{str(st):22s} {a:>18,.2f} {b:>18,.2f} {d:>16,.2f} {dc:>14,.2f}{flag}")
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
