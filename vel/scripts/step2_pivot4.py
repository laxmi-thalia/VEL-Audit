import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
xlDatabase, xlRowField, xlSum, xlTabularRow = 1, 1, -4157, 1
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(P)
    xl.CalculateFullRebuild()
    for w in list(wb.Worksheets):
        if w.Name == "Pivot - Month on Month": w.Delete()
    src = wb.Worksheets("Month-on-Month")
    last = src.Cells(src.Rows.Count, 1).End(-4162).Row
    rng = src.Range("A3:P%d" % last)
    pv = wb.Worksheets.Add(Before=src); pv.Name = "Pivot - Month on Month"
    cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=rng)
    pt = cache.CreatePivotTable(TableDestination=pv.Range("A4"), TableName="PT_Long")
    for pos, fld in enumerate(["State", "Month", "GSTR-1 Type"], 1):
        f = pt.PivotFields(fld); f.Orientation = xlRowField; f.Position = pos
    for fld, cap in [("Books Taxable", "Books Taxable "), ("Books IGST", "Books IGST "), ("Books CGST", "Books CGST "), ("Books SGST", "Books SGST "),
                     ("GSTR-1 Taxable", "GSTR-1 Taxable "), ("GSTR-1 IGST", "GSTR-1 IGST "), ("GSTR-1 CGST", "GSTR-1 CGST "), ("GSTR-1 SGST", "GSTR-1 SGST "),
                     ("Diff Taxable", "Diff Taxable "), ("Diff IGST", "Diff IGST "), ("Diff CGST", "Diff CGST "), ("Diff SGST", "Diff SGST ")]:
        d = pt.AddDataField(pt.PivotFields(fld), cap, xlSum); d.NumberFormat = "#,##0.00"
    pt.RowAxisLayout(xlTabularRow)                 # one flat row per State | Month | Type
    for fld in ["State", "Month", "GSTR-1 Type"]:
        try: pt.PivotFields(fld).Subtotals = [False] * 12
        except Exception: pass
    try: pt.PivotFields("State").RepeatLabels = True; pt.PivotFields("Month").RepeatLabels = True
    except Exception: pass
    order = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26"]
    try:
        for i, mo in enumerate(order, 1): pt.PivotFields("Month").PivotItems(mo).Position = i
    except Exception as e: print("month order partial:", e)
    pv.Range("A1").Value = "Sales Register vs GSTR-1 — State | Month | GSTR-1 Type, with Books / GSTR-1 / Difference for Taxable, IGST, CGST, SGST"
    pv.Range("A2").Value = "PivotTable over the 'Month-on-Month' sheet (long format). Collapse a state or filter a type as needed."
    pv.Range("A1").Font.Bold = True
    pt.RefreshTable()
    # hide zero-only rows? keep all; CA can filter. Check error cells on the long sheet
    try: e = src.Range("E4:P%d" % last).SpecialCells(-4123, 16).Count
    except Exception: e = 0
    tot_sr = xl.WorksheetFunction.Sum(src.Range("E4:E%d" % last))
    tot_g1 = xl.WorksheetFunction.Sum(src.Range("I4:I%d" % last))
    nz = sum(1 for r in range(4, last + 1) if abs(src.Cells(r, 13).Value or 0) >= 1)
    print("long sheet errors:", e, "| SR total:", round(tot_sr, 2), "| GSTR-1 total:", round(tot_g1, 2), "| rows with |diff|>=1:", nz)
    print("pivot at", pt.TableRange1.Address, "| sheets:", [w.Name for w in wb.Worksheets])
    wb.Save()
finally:
    if wb is not None: wb.Close(SaveChanges=False)
    xl.Quit()
