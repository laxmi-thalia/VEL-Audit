import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\VEL_2425.xlsb"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P, ReadOnly=True, UpdateLinks=0)
    ws=wb.Sheets("Sales Reco"); ur=ws.UsedRange
    print("Sales Reco:", ur.Rows.Count, "x", ur.Columns.Count)
    for i,row in enumerate(ur.Value,1):
        cells=[(j,str(v).strip()) for j,v in enumerate(row,1) if v is not None and str(v).strip()!=""]
        if not cells: continue
        show=[]
        for j,v in cells[:10]:
            try: v="{:,.0f}".format(float(v)) if v.replace('.','',1).replace('-','',1).replace('e','').replace('+','').isdigit() or 'e+' in v else v
            except Exception: pass
            show.append("c%d:%s" % (j, v[:34]))
        print("r%02d: %s" % (i, " | ".join(show)))
finally:
    if wb is not None: wb.Close(False)
    xl.Quit()
