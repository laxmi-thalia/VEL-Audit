import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    g = wb.Worksheets("Step 1 Flags"); ws = wb.Worksheets("SR_2025-26")
    try: e = g.UsedRange.SpecialCells(-4123, 16).Count
    except Exception: e = 0
    first = 22
    def listed(): return sum(1 for r in range(first, first + 400) if g.Cells(r, 2).Value not in (None, ""))
    print("flags sheet errors:", e, "| listed rows (live):", listed(), "| gate blocking:", g.Range("B4").Value)
    print("first 3:", [(g.Cells(r, 2).Value, str(g.Cells(r, 3).Value)[:45], g.Cells(r, 6).Value) for r in range(first, first + 3)])
    # live test: fill the Document Date on the first listed row -> that flag drops; the list re-packs
    row = int(g.Cells(first, 2).Value)
    old = ws.Cells(row, 5).Value; ws.Cells(row, 5).Value = "2025-10-17"; xl.Calculate()
    print("after fixing register row %d's date: listed = %d | its flag now = %r" % (row, listed(), ws.Cells(row, 53).Value))
    ws.Cells(row, 5).Value = old; xl.Calculate()
    print("restored: listed =", listed(), "| flag =", repr(ws.Cells(row, 53).Value))
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
