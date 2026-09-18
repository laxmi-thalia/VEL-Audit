import win32com.client as win32, pythoncom
P=r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
pythoncom.CoInitialize()
xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=None
try:
    wb=xl.Workbooks.Open(P)
    pv=wb.Worksheets("Pivot - Month on Month"); pt=pv.PivotTables("PT_MoM")
    msg=[]
    for fld,nm in [("Doc Type","SlDocType"),("State","SlState")]:
        try:
            sc=wb.SlicerCaches.Add2(pt, fld, nm)
            sc.Slicers.Add(SlicerDestination=pv, Name=nm, Caption=fld,
                           Top=10, Left=780 if fld=="Doc Type" else 920, Width=140, Height=170)
            msg.append(f"{fld}: OK")
        except Exception as ex:
            msg.append(f"{fld}: FAILED {ex.args[2][5] if ex.args and len(ex.args)>2 else ex}")
    print(" | ".join(msg))
    print("page field 'Doc Type' present:", any(f.Name=="Doc Type" for f in pt.PageFields))
    print("row field:", [f.Name for f in pt.RowFields], "col field:", [f.Name for f in pt.ColumnFields])
    print("data fields:", [f.Name for f in pt.DataFields])
    gt=pt.TableRange1
    print("pivot range:", gt.Address, "| grand-total cell sample:", pv.Range("A5").Value)
    # read the grand total row values for the three measures
    r=gt.Rows.Count; 
    print("last row label:", pv.Cells(gt.Row+r-1,1).Value)
    wb.Save()
finally:
    if wb is not None: wb.Close(SaveChanges=False)
    xl.Quit()
