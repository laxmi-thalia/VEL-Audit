import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\VEL_2425.xlsb"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P, ReadOnly=True, UpdateLinks=0)
    ws=wb.Sheets("GSTR-1 vs GSTR-3B"); ur=ws.UsedRange
    print("LAYOUT 'GSTR-1 vs GSTR-3B':", ur.Rows.Count,"x",ur.Columns.Count)
    for i,row in enumerate(ur.Value,1):
        cells=[(j,str(v).strip()) for j,v in enumerate(row,1) if v is not None and str(v).strip()!=""]
        if not cells: continue
        txt=[c for c in cells if not str(c[1]).replace('.','',1).replace('-','',1).replace('e','',1).isdigit()]
        show=[f"c{j}:{v[:30]}" for j,v in (txt if txt else cells)[:9]]
        print(f"  r{i:2d} ({len(cells)}): "+" | ".join(show))
finally:
    if wb is not None: wb.Close(SaveChanges=False)
    xl.Quit()
