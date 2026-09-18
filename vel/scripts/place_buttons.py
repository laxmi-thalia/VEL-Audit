"""Reposition the << INDEX buttons so they never cover content:
- data sheets (headers across row 1): one blank column after the last header, inside the frozen row
- title sheets (only A1 text in row 1): past the title's rendered width
Sized to row-1 height so nothing below is touched. Placement=2 (moves with cells, doesn't size)."""
import win32com.client as win32, pythoncom, time
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f=open(P,'r+b'); f.close()
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P); xl.Calculation=-4135
    report=[]
    for ws in wb.Worksheets:
        if ws.Name=="INDEX": continue
        btn=None
        for s in ws.Shapes:
            if s.Name=="btnIndex": btn=s
        if btn is None: report.append((ws.Name,"MISSING")); continue
        last=ws.Cells(1, ws.Columns.Count).End(-4159).Column   # xlToLeft: last used col in row 1
        a1=str(ws.Cells(1,1).Value or "")
        if last<=1:
            left=max(470.0, len(a1)*6.5+30.0)                  # clear of the A1 title overflow
        else:
            left=float(ws.Cells(1,last+2).Left)                # one blank col after the last header
        h=min(15.0, float(ws.Rows(1).Height)-1.0)
        btn.Left=left; btn.Top=1.0; btn.Width=86.0; btn.Height=max(12.0,h)
        btn.Placement=2
        btn.TextFrame2.TextRange.Font.Size=8
        report.append((ws.Name, "last col %d -> left %.0f" % (last,left)))
    xl.Calculation=-4105
    wb.Close(SaveChanges=True)
    for r_ in report: print("%-28s %s" % r_)
finally:
    xl.Quit()
