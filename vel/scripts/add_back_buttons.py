"""Back-navigation: a floating '<< INDEX' button (shape + hyperlink) on every sheet,
top area ~column H so it never covers the A-column titles. Idempotent. Also a defined
name INDEX_HOME so F5 / Name Box can jump from anywhere."""
import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
try:
    f=open(P,'r+b'); f.close()
except Exception as e:
    print("LOCKED - close the master first:", e); raise SystemExit(1)
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P)
    xl.Calculation=-4135
    made=0
    for ws in wb.Worksheets:
        if ws.Name=="INDEX": continue
        for shp in list(ws.Shapes):
            if shp.Name=="btnIndex": shp.Delete()
        shp=ws.Shapes.AddShape(5, 470, 2, 86, 17)   # msoShapeRoundedRectangle
        shp.Name="btnIndex"
        shp.Fill.ForeColor.RGB=0x794E1F            # BGR of 1F4E79
        shp.Line.Visible=False
        tr=shp.TextFrame2.TextRange
        tr.Text="<< INDEX"
        tr.Font.Bold=True; tr.Font.Size=9; tr.Font.Fill.ForeColor.RGB=0xFFFFFF
        ws.Hyperlinks.Add(Anchor=shp, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
        made+=1
    try: wb.Names("INDEX_HOME").Delete()
    except Exception: pass
    wb.Names.Add("INDEX_HOME", "='INDEX'!$A$1")
    wb.Close(SaveChanges=True)
    print("buttons added on", made, "sheets | defined name INDEX_HOME set")
finally:
    xl.Quit()
