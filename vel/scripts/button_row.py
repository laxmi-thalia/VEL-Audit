"""Reserve row 1 on every visible sheet (except INDEX) for the << INDEX button.
Row insert via Excel -> all formulas/freezes/filters/pivot shift correctly."""
import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f=open(P,'r+b'); f.close()
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P)
    done=0
    for ws in wb.Worksheets:
        if ws.Name=="INDEX" or ws.Visible!=-1: continue
        ws.Activate()
        # guard: skip if row 1 already looks like a reserved button row (empty + short height marker)
        if str(ws.Cells(1,1).Value or "")=="" and abs(ws.Rows(1).RowHeight-21.0)<0.01:
            pass
        ws.Rows(1).Insert(-4121)      # shift down
        ws.Rows(1).RowHeight=21.0
        btn=None
        for s in ws.Shapes:
            if s.Name=="btnIndex": btn=s
        if btn is None:
            btn=ws.Shapes.AddShape(5, 2.0, 2.0, 86.0, 17.0); btn.Name="btnIndex"
            btn.Fill.ForeColor.RGB=0x794E1F; btn.Line.Visible=False
            t=btn.TextFrame2.TextRange; t.Text="<< INDEX"; t.Font.Bold=True; t.Font.Size=9; t.Font.Fill.ForeColor.RGB=0xFFFFFF
            ws.Hyperlinks.Add(Anchor=btn, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
        btn.Left=2.0; btn.Top=2.0; btn.Width=86.0; btn.Height=17.0; btn.Placement=2
        done+=1
    print("row inserted + button parked on", done, "sheets")
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
