"""C4 (Rashid 28-09): ITC Register 2025-26 `GSTR 2B/6A Period` shows the date only (dd-mm-yy), no time. Format only."""
import win32com.client as w
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2) (1) (1).xlsb"
xl = w.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P, UpdateLinks=0)
    ws = wb.Worksheets("ITC Register 2025-26")
    col = next(j for j in range(1, 120) if ws.Cells(5, j).Value == "GSTR 2B/6A Period")
    ur = ws.UsedRange; last = ur.Row + ur.Rows.Count - 1
    before = ws.Cells(6, col).NumberFormat
    ws.Range(ws.Cells(6, col), ws.Cells(last, col)).NumberFormat = "dd-mm-yy"
    print(f"column {col}: format {before!r} -> {ws.Cells(6, col).NumberFormat!r}; shows {ws.Cells(6, col).Text!r}; value unchanged {ws.Cells(6, col).Value!r}")
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
