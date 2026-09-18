import win32com.client as win32, pythoncom, collections
P=r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\VEL_2425.xlsb"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P, ReadOnly=True, UpdateLinks=0)
    ws=wb.Sheets("Sales Register")
    # find header row and query-ish columns in first 6 rows x 124 cols
    hdr=None; qcols=[]
    for r in range(1,7):
        vals=[(c, str(ws.Cells(r,c).Value or "").strip()) for c in range(1,125)]
        names=[v for _,v in vals if v]
        if len(names)>30:
            hdr=r
            qcols=[(c,v) for c,v in vals if "quer" in v.lower() or "remark" in v.lower() or "comment" in v.lower()]
            break
    print("header row:", hdr, "| query-like columns:", qcols)
    last=ws.Cells(ws.Rows.Count,1).End(-4162).Row
    for c,name in qcols:
        col=ws.Range(ws.Cells(hdr+1,c), ws.Cells(min(last,24301),c)).Value
        cnt=collections.Counter()
        for row in col:
            v=str(row[0] or "").strip()
            if v: cnt[v]+=1
        print("\n=== column %r (%d filled):" % (name, sum(cnt.values())))
        for k,v in cnt.most_common(20): print("  x%-5d %s" % (v,k[:140]))
finally:
    if wb is not None: wb.Close(False)
    xl.Quit()
