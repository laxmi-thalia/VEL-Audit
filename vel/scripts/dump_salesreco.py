"""Dump last year's 'Sales Reco' sheet: values, formulas, and full styling, to salesreco_dump.json."""
import win32com.client as win32, pythoncom, json, os
P = os.path.abspath("VEL_2425_fresh.xlsb")
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wb=xl.Workbooks.Open(P, ReadOnly=True)
    names=[w.Name for w in wb.Worksheets]
    tgt=[n for n in names if "sales reco" in n.lower()]
    print("sheets containing 'sales reco':", tgt)
    ws=wb.Worksheets(tgt[0])
    ur=ws.UsedRange
    R0,C0=ur.Row,ur.Column; NR,NC=ur.Rows.Count,ur.Columns.Count
    print("used range: rows %d-%d cols %d-%d" % (R0,R0+NR-1,C0,C0+NC-1))
    out={"sheet":ws.Name,"rows":NR,"cols":NC,"r0":R0,"c0":C0,
         "gridlines":bool(xl.ActiveWindow.DisplayGridlines),
         "freeze":None,"colw":{},"rowh":{},"merged":[],"cells":{}}
    try:
        ws.Activate()
        aw=xl.ActiveWindow
        out["freeze"]=[aw.SplitRow,aw.SplitColumn] if aw.FreezePanes else None
        out["gridlines"]=bool(aw.DisplayGridlines)
    except Exception: pass
    for c in range(C0, min(C0+NC, 60)):
        out["colw"][c]=round(ws.Columns(c).ColumnWidth,2)
    for r in range(R0, min(R0+NR, 120)):
        out["rowh"][r]=round(ws.Rows(r).RowHeight,2)
    # merged areas
    seen=set()
    for r in range(R0, min(R0+NR,120)):
        for c in range(C0, min(C0+NC,60)):
            cell=ws.Cells(r,c)
            if cell.MergeCells:
                a=cell.MergeArea.Address
                if a not in seen: seen.add(a); out["merged"].append(a)
    for r in range(R0, min(R0+NR,120)):
        for c in range(C0, min(C0+NC,60)):
            cell=ws.Cells(r,c)
            v=cell.Value; f=cell.Formula
            has = v is not None or (isinstance(f,str) and f.startswith("="))
            fill=int(cell.Interior.Color) if cell.Interior.ColorIndex!=-4142 else None
            bold=bool(cell.Font.Bold); italic=bool(cell.Font.Italic)
            fc=int(cell.Font.Color); sz=float(cell.Font.Size); fn=str(cell.Font.Name)
            bl=[int(cell.Borders(i).LineStyle) for i in (7,8,9,10)]  # L,T,B,R
            if not has and fill is None and not any(x!=-4142 for x in bl): continue
            out["cells"]["%d,%d"%(r,c)]={
                "v":str(v)[:120] if v is not None else None,
                "f":f if isinstance(f,str) and f.startswith("=") else None,
                "nf":str(cell.NumberFormat)[:40],"fill":fill,"b":bold,"i":italic,
                "fc":fc,"sz":sz,"fn":fn,"bd":bl,
                "ha":int(cell.HorizontalAlignment),"wrap":bool(cell.WrapText)}
    json.dump(out, open("salesreco_dump.json","w"), indent=0)
    print("dumped", len(out["cells"]), "cells | merged areas:", len(out["merged"]))
    wb.Close(SaveChanges=False)
finally:
    xl.Quit()
