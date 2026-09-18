"""Re-add the << INDEX buttons after ANY openpyxl edit of the master (openpyxl drops shapes on save).
Row 1 of every sheet is RESERVED for the button (inserted 2026-08-27) - park at top-left (2,2).
Run: venv python readd_buttons.py"""
import win32com.client as win32, pythoncom, time
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
f = open(P, 'r+b'); f.close()
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
def put(ws):
    for s in list(ws.Shapes):
        if s.Name == "btnIndex": s.Delete()
    ws.Rows(1).RowHeight = 21
    s = ws.Shapes.AddShape(5, 2.0, 2.0, 86.0, 17.0); s.Name = "btnIndex"
    s.Fill.ForeColor.RGB = 0x794E1F; s.Line.Visible = False; s.Placement = 2
    t = s.TextFrame2.TextRange; t.Text = "<< INDEX"; t.Font.Bold = True; t.Font.Size = 9; t.Font.Fill.ForeColor.RGB = 0xFFFFFF
    ws.Hyperlinks.Add(Anchor=s, Address="", SubAddress="'INDEX'!A1", ScreenTip="Back to INDEX")
try:
    wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    fails = []
    for ws in wb.Worksheets:
        if ws.Name == "INDEX": continue
        try:
            ws.Activate(); put(ws)
        except Exception:
            fails.append(ws.Name)
    for nm in fails:
        ws = wb.Worksheets(nm); ws.Activate(); time.sleep(0.3); ws.Range("A1").Select(); put(ws)
    try: wb.Names("INDEX_HOME").Delete()
    except Exception: pass
    wb.Names.Add("INDEX_HOME", "='INDEX'!$A$1")
    wb.Worksheets("INDEX").Activate(); xl.Calculation = -4105
    wb.Close(SaveChanges=True); print("buttons restored in reserved row 1 (retries: %s)" % fails)
finally:
    xl.Quit()
