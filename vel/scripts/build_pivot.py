"""Create 'S2 Pivot Month-on-Month' in the master: real PivotTable over S2 Month-on-Month
(A3:U1371), same definition as the step-2 file's PT_Long: rows State/Month/Type, 12 sum fields."""
import win32com.client as win32, pythoncom, time
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f=open(P,'r+b'); f.close()
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
NM="S2 Pivot Month-on-Month"
try:
    wb=xl.Workbooks.Open(P)
    xl.CalculateFullRebuild()   # pivot cache snapshots values - make them current first
    for w in list(wb.Worksheets):
        if w.Name==NM: xl.DisplayAlerts=False; w.Delete()
    anchor=wb.Worksheets("S2 Reco (CA format)")
    ws=wb.Worksheets.Add(None, anchor)   # after the CA-format reco
    ws.Name=NM
    ws.Cells(1,1).Value="Sales Register vs GSTR-1 - PivotTable view (double-click any figure to drill into the underlying state-month rows)"
    ws.Cells(1,1).Font.Bold=True
    ws.Cells(2,1).Value="Source: 'S2 Month-on-Month' (live formulas). After editing the register, right-click the pivot > Refresh."
    src="'S2 Month-on-Month'!R3C1:R1371C21"
    pc=wb.PivotCaches().Create(SourceType=1, SourceData=src)   # xlDatabase
    pt=pc.CreatePivotTable(TableDestination="'%s'!R4C1"%NM, TableName="PT_Long")
    pt.RowAxisLayout(1)  # tabular
    for fname in ("State","Month","GSTR-1 Type"):
        fld=pt.PivotFields(fname); fld.Orientation=1  # xlRowField
    pt.PivotFields("Month").Subtotals=[False]*12
    pt.PivotFields("GSTR-1 Type").Subtotals=[False]*12
    for fname in ("Books Taxable","Books IGST","Books CGST","Books SGST",
                  "GSTR-1 Taxable","GSTR-1 IGST","GSTR-1 CGST","GSTR-1 SGST",
                  "Diff Taxable","Diff IGST","Diff CGST","Diff SGST"):
        df=pt.AddDataField(pt.PivotFields(fname), "%s " % fname, -4157)  # xlSum
        df.NumberFormat="#,##0.00"
    pt.TableStyle2="PivotStyleLight16"
    pt.ColumnGrand=True; pt.RowGrand=True
    ws.Columns("A:C").AutoFit()
    gt=pt.GetData("'Diff Taxable '")
    print("pivot grand total Diff Taxable:", format(gt,",.2f"))
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
