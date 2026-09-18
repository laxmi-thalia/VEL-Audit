import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
xlDatabase, xlRowField, xlColumnField, xlDataField, xlSum = 1,1,2,4,-4157
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P)
    src=wb.Worksheets("Month-on-Month")
    last=src.Cells(src.Rows.Count,1).End(-4162).Row
    rng=src.Range(f"A1:J{last}")
    try: wb.Worksheets("Pivot - Month on Month").Delete()
    except Exception: pass
    pv=wb.Worksheets.Add(); pv.Name="Pivot - Month on Month"
    pv.Move(After=wb.Worksheets("Reconciliation Summary"))
    pv=wb.Worksheets("Pivot - Month on Month")
    cache=wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=rng)
    pt=cache.CreatePivotTable(TableDestination=pv.Range("A4"), TableName="PT_MoM")
    pt.PivotFields("State").Orientation=xlRowField; pt.PivotFields("State").Position=1
    pt.PivotFields("GSTR-1 Month").Orientation=xlColumnField
    for fld,cap in [("Taxable Books","Books"),("Taxable GSTR-1","GSTR-1"),("Taxable DIFF","Difference")]:
        d=pt.AddDataField(pt.PivotFields(fld), f"{cap} ", xlSum); d.NumberFormat="#,##0.00"
    pt.PivotFields("Doc Type").Orientation=3          # xlPageField
    try:
        sc=wb.SlicerCaches.Add2(pt,"Doc Type"); sc.Slicers.Add(pv,None,"DocType","Doc Type",30,700,120,180)
        sc2=wb.SlicerCaches.Add2(pt,"State");   sc2.Slicers.Add(pv,None,"StateSl","State",30,880,180,260)
        slic="slicers added"
    except Exception as ex: slic=f"slicer skipped: {ex}"
    pv.Range("A1").Value="VEL — Sales Register vs GSTR-1, month on month"
    pv.Range("A2").Value="Use the Doc Type / State slicers. Difference = Books minus GSTR-1."
    pv.Range("A1").Font.Bold=True; pv.Range("A1").Font.Size=13
    pt.RefreshTable()
    print("pivot rows:", pt.TableRange1.Rows.Count, "| cols:", pt.TableRange1.Columns.Count, "|", slic)
    wb.Save(); wb.Close(SaveChanges=False)
finally:
    xl.Quit()
