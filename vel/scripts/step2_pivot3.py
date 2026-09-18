import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
xlDatabase, xlRowField, xlColumnField, xlSum = 1, 1, 2, -4157
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
wb = None
try:
    wb = xl.Workbooks.Open(P)
    for w in list(wb.Worksheets):
        if w.Name == "Pivot - Month on Month": w.Delete()
    src = wb.Worksheets("Month-on-Month")
    last = src.Cells(src.Rows.Count, 1).End(-4162).Row
    rng = src.Range("A3:O%d" % last)
    pv = wb.Worksheets.Add(Before=src); pv.Name = "Pivot - Month on Month"
    cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=rng)
    pt = cache.CreatePivotTable(TableDestination=pv.Range("A4"), TableName="PT_MoM")
    f = pt.PivotFields("State"); f.Orientation = xlRowField; f.Position = 1
    m = pt.PivotFields("Month"); m.Orientation = xlColumnField
    for fld, cap in [("Books Taxable", "Books "), ("GSTR-1 Taxable", "GSTR-1 "), ("Diff Taxable", "Difference ")]:
        d = pt.AddDataField(pt.PivotFields(fld), cap, xlSum); d.NumberFormat = "#,##0.00"
    # the per-state Total rows on the source sheet have a blank Month -> hide that column
    try: pt.PivotFields("Month").PivotItems("(blank)").Visible = False
    except Exception as e: print("could not hide (blank):", e)
    # chronological month order
    order = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26"]
    try:
        for i, mo in enumerate(order, 1): pt.PivotFields("Month").PivotItems(mo).Position = i
    except Exception as e: print("month ordering partial:", e)
    pv.Range("A1").Value = "Sales Register vs GSTR-1 — month on month (PivotTable over the 'Month-on-Month' sheet)"
    pv.Range("A2").Value = "Difference = Books minus GSTR-1. Drag 'GSTIN' in for a GSTIN view; Diff IGST/CGST/SGST are available as extra values."
    pv.Range("A1").Font.Bold = True; pv.Range("A1").Font.Size = 12
    pt.RefreshTable()
    print("pivot at", pt.TableRange1.Address, "| sheets:", [w.Name for w in wb.Worksheets])
    wb.Save()
finally:
    if wb is not None: wb.Close(SaveChanges=False)
    xl.Quit()
