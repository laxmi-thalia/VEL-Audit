import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\VEL_2425.xlsb"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P, ReadOnly=True, UpdateLinks=0)
    ws=wb.Sheets("GSTR-1 vs GSTR-3B")
    print("rows 41-53, block-2 columns (GSTIN, State, G1 taxable, 3B taxable, Diff taxable, Remark):")
    for r in range(31,54):
        g=ws.Cells(r,1).Value; st=ws.Cells(r,2).Value
        a=ws.Cells(r,3).Value; b=ws.Cells(r,7).Value; d=ws.Cells(r,11).Value; rem=ws.Cells(r,15).Value
        if g is None and st is None: continue
        f=lambda v: ("" if v is None else (f"{v:,.2f}" if isinstance(v,(int,float)) else str(v)))
        print(f"  r{r:2d} {str(st)[:20]:20s} G1={f(a):>18} 3B={f(b):>18} diff={f(d):>14}  {str(rem)[:70] if rem else ''}")
    print()
    print("full remark cells in column 15:")
    for r in range(29,54):
        v=ws.Cells(r,15).Value
        if v: print(f"  r{r}: {v}")
finally:
    if wb is not None: wb.Close(SaveChanges=False)
    xl.Quit()
