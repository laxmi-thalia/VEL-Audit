import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    ws=wb.Worksheets("SR vs GSTR-1 (CA format)")
    try: errs=ws.Range("A1:AD74").SpecialCells(-4123,16).Count
    except Exception: errs=0
    print("formula ERROR cells:", errs)
    G=["B2B","B2C","Credit note","Debit note","Advance received","Advance adjusted","Total liability"]
    for lbl,row in [("As per GSTR-1",26),("Sales Register",50),("DIFFERENCE",74)]:
        vals=[round(ws.Cells(row,3+i*4).Value or 0,2) for i in range(7)]
        print(f"\n{lbl} - taxable value by group:")
        for g,v in zip(G,vals): print(f"    {g:20s} {v:>18,.2f}")
    print("\nDifference rows that are non-zero (taxable, any group):")
    for r in range(55,74):
        gst=ws.Cells(r,1).Value; st=ws.Cells(r,2).Value
        nz=[(G[i],round(ws.Cells(r,3+i*4).Value or 0,2)) for i in range(7) if abs(ws.Cells(r,3+i*4).Value or 0)>=1]
        if nz: print(f"    {st:20s} " + ", ".join(f"{g}={v:,.2f}" for g,v in nz))
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
