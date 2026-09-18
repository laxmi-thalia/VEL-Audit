"""Step 2 (COM): insert 'CC Name' column right after Profit Centre in SR_2025-26,
fill with live INDEX/MATCH into CC Master. Column insert via Excel updates every formula."""
import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f=open(P,'r+b'); f.close()
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P)
    ws=wb.Worksheets("SR_2025-26")
    # find Profit Centre column by header (row 4)
    pc_col=None; hdrs={}
    for c in range(1, 70):
        h=str(ws.Cells(4,c).Value or "").strip()
        hdrs[h]=c
        if h=="Profit Centre": pc_col=c
    assert pc_col, "Profit Centre header not found"
    if "CC Name" in hdrs:
        print("CC Name already present at col", hdrs["CC Name"]); wb.Close(False); raise SystemExit
    ncc=wb.Worksheets("CC Master").Cells(wb.Worksheets("CC Master").Rows.Count,1).End(-4162).Row
    new=pc_col+1
    ws.Columns(new).Insert(-4161)     # shift right; Excel rewrites all dependent formulas
    h=ws.Cells(4,new); h.Value="CC Name"
    src=ws.Cells(4,pc_col)
    h.Font.Bold=True; h.Font.Color=src.Font.Color; h.Interior.Color=src.Interior.Color
    def colL(n):
        s=""
        while n: n,rem=divmod(n-1,26); s=chr(65+rem)+s
        return s
    from_l=colL(pc_col)
    rng=ws.Range(ws.Cells(5,new), ws.Cells(27006,new))
    rng.Formula = '=IF(%s5="","",IFERROR(INDEX(\'CC Master\'!$B$2:$B$%d,MATCH(VALUE(%s5),\'CC Master\'!$A$2:$A$%d,0)),""))' % (from_l,ncc,from_l,ncc)
    ws.Columns(new).ColumnWidth=22
    xl.CalculateFullRebuild()
    filled=xl.WorksheetFunction.CountIf(rng,"?*")
    print("inserted 'CC Name' at column", new, "| non-empty narrations:", int(filled), "of 27002")
    # spot goldens after the insert
    print("register R2:", format(ws.Range("R2").Value,",.2f"))
    s2=wb.Worksheets("S2 Reco (CA format)")
    print("S2 diff AA74:", format(s2.Range("AA74").Value,",.2f"))
    s7=wb.Worksheets("S7 CN Time-bar"); l7=s7.Cells(s7.Rows.Count,2).End(-4162).Row
    beyond=sum(1 for r in range(4,l7+1) if "BEYOND" in str(s7.Cells(r,8).Value or ""))
    print("S7 beyond:", beyond)
    tot=0
    for w in wb.Worksheets:
        try: e=w.UsedRange.SpecialCells(-4123,16).Count
        except Exception: e=0
        tot+=e
        if e: print("ERRORS in",w.Name,e)
    print("formula ERROR cells:",tot)
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
