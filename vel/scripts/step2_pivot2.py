import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
xlDatabase,xlRowField,xlColumnField,xlPageField,xlSum = 1,1,2,3,-4157
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P)
    for w in list(wb.Worksheets):
        if w.Name=="Pivot - Month on Month": w.Delete()
    src=wb.Worksheets("Month-on-Month")
    last=src.Cells(src.Rows.Count,1).End(-4162).Row
    rng=src.Range("A1:J%d"%last)
    pv=wb.Worksheets.Add(Before=src)
    pv.Name="Pivot - Month on Month"
    cache=wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=rng)
    pt=cache.CreatePivotTable(TableDestination=pv.Range("A5"), TableName="PT_MoM")
    f=pt.PivotFields("State"); f.Orientation=xlRowField; f.Position=1
    pt.PivotFields("GSTR-1 Month").Orientation=xlColumnField
    for fld,cap in [("Taxable Books","Books "),("Taxable GSTR-1","GSTR-1 "),("Taxable DIFF","Difference ")]:
        d=pt.AddDataField(pt.PivotFields(fld), cap, xlSum); d.NumberFormat="#,##0.00"
    pt.PivotFields("Doc Type").Orientation=xlPageField
    msg=""
    try:
        sc=wb.SlicerCaches.Add2(pt,"Doc Type"); sc.Slicers.Add(pv,None,"DocTypeSl","Doc Type",8,760,110,150)
        s2=wb.SlicerCaches.Add2(pt,"State");   s2.Slicers.Add(pv,None,"StateSl","State",8,890,190,260)
        msg="slicers: Doc Type + State"
    except Exception as ex: msg="slicer skipped: %s"%ex
    pv.Range("A1").Value="VEL - Sales Register vs GSTR-1, month on month"
    pv.Range("A2").Value="Difference = Books minus GSTR-1. Use the Doc Type / State slicers to filter both sides together."
    pv.Range("A1").Font.Bold=True; pv.Range("A1").Font.Size=13
    pt.RefreshTable()
    print("pivot placed:", pt.TableRange1.Address, "|", msg)
    print("sheets:", [w.Name for w in wb.Worksheets])
    wb.Save()
finally:
    if wb is not None: wb.Close(SaveChanges=False)
    xl.Quit()
