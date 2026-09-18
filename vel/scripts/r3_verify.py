import win32com.client as win32, pythoncom, collections
P=r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    ws=wb.Worksheets("SR_2025-26")
    try: e=ws.Range("A5:BD27006").SpecialCells(-4123,16).Count
    except Exception: e=0
    print("register formula errors:", e, "| R2 taxable:", ws.Range("R2").Value)
    flags=[r[0] for r in ws.Range("BA5:BA27006").Value]
    nz=[f for f in flags if f]; print("rows with a live flag:", len(nz))
    c=collections.Counter()
    for f in nz:
        for part in str(f).split("; "):
            if part: c[part]+=1
    for k,v in c.most_common(): print("   %-42s %d"%(k,v))
    g=wb.Worksheets("Step 1 Flags")
    print("GATE -> Blocking:", g.Range("B4").Value, "| Review:", g.Range("B5").Value, "|", g.Range("B7").Value)
    # live test: fix the undated CRN's date in memory and watch the flag clear, then undo
    r=None
    for i,f in enumerate(flags, start=5):
        if f and "Document Date missing" in f: r=i; break
    if r:
        old=ws.Cells(r,5).Value; ws.Cells(r,5).Value="2025-10-17"; xl.Calculate()
        print("live test row %d: after filling Document Date -> flag = %r" % (r, ws.Cells(r,53).Value))
        ws.Cells(r,5).Value=old; xl.Calculate()
        print("                  restored -> flag = %r | gate blocking = %s" % (ws.Cells(r,53).Value, g.Range("B4").Value))
    wb.Close(SaveChanges=True)
finally: xl.Quit()
