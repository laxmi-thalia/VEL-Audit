"""Rebuild INDEX in last year's exact format (Arena | Descriptions | ... | Particulars-link),
Sales arena only. Cell hyperlinks like last year, dark header 333F4F, white canvas."""
import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
ROWS = [
    ("Sales Register for FY 2025-26", "Sales Register", "SR_2025-26"),
    ("Revenue Consolidated Reco-Books vs GST", "Rev Conso", "S9 Sales Reco"),
    ("HSN Summary of Outward Supplies", "HSN Summary", "S5 HSN Summary"),
    ("Sales Rate-wise summary", "Rate-wise Summary", "S5 Rate-wise"),
    ("Sales Register vs GSTR-1", "Reg vs GSTR 1", "S2 Reco (CA format)"),
    ("GSTR 1 vs GSTR 3B", "GSTR 1 vs GSTR 3B", "S3 1 vs 3B"),
    ("GL vs Sales Register", "Outward GL vs  Reg", "S4 GL vs SR"),
    ("Sales Register vs GSTR-1 - Month on Month", "Reg vs GSTR 1 MoM", "S2 Month-on-Month"),
    ("Sales Register vs GSTR-1 - Pivot view", "Reg vs GSTR 1 Pivot", "S2 Pivot Month-on-Month"),
    ("Sales Register vs GSTR-1 - Exceptions", "Reg vs GSTR 1 Exceptions", "S2 Exceptions"),
    ("GSTR 1 vs GSTR 3B - Month on Month", "GSTR 1 vs 3B MoM", "S3 Month-on-Month"),
    ("Sales Register vs GSTR-3B", "Reg vs GSTR 3B", "S3 SR vs 3B"),
    ("Sales Register vs GSTR-3B - Month on Month", "Reg vs GSTR 3B MoM", "S3 SR vs 3B MoM"),
    ("Amendments in FY 2026-27 returns", "Amendment Check", "S3 Amendment Check"),
    ("GL vs Sales Register - Exceptions", "GL vs Reg Exceptions", "S4 Exceptions"),
    ("Advances Control Account", "Advances Control", "S6 Advances Control"),
    ("Advances Control - Month on Month", "Advances MoM", "S6 Month-on-Month"),
    ("Advance Ledger vs Books", "GL Advance Check", "S6 GL Advance Check"),
    ("Credit Note Time-bar (Sec 34(2))", "CN Time-bar", "S7 CN Time-bar"),
    ("Sample Selection with PO/SO references", "Samples", "Sample Selection"),
    ("Data Quality Flags (live gate)", "Flags", "Step 1 Flags"),
    ("Draft Queries to Client", "Queries", "Queries (Draft)"),
    ("SAP Enrichment notes", "SAP Enrichment", "SAP Enrichment"),
    ("Source File Load Log", "Load Log", "Load Log"),
    ("Format Legend (CA's format notes)", "Format Legend", "Format Legend (preserved)"),
    ("GSTR-1 as filed + vice-versa match", "GSTR-1 Data", "GSTR-1 Data"),
    ("GSTR-3B 3.1(a) + vice-versa status", "3B Data", "3B Data"),
    ("SAP Output GL + vice-versa match", "GL Data", "GL Data"),
    ("FS Revenue extract (state level)", "FS Revenue Data", "FS Revenue Data"),
    ("Open Points - all steps", "Open Points (all)", "Open Points (all)"),
    ("Open Points - register", "Open Points S1", "Open Points"),
]
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
DARK = 0x4F3F33   # BGR of RGB 333F4F
try:
    wb = xl.Workbooks.Open(P)
    ws = wb.Worksheets("INDEX")
    ws.Cells.Clear()
    for hl in list(ws.Hyperlinks): hl.Delete()
    ws.Activate(); xl.ActiveWindow.DisplayGridlines = True
    last = 4 + len(ROWS)
    canvas = ws.Range(ws.Cells(2, 1), ws.Cells(last + 2, 10))
    canvas.Interior.Color = 0xFFFFFF
    t = ws.Cells(2, 1); t.Value = "Vikran Engineering Limited- GST Audit for FY 2025-26"
    t.Font.Bold = True; t.Font.Size = 12
    heads = ["Arena", "Descriptions", "State Code", "GSTN", "GSTR-9", "GSTR-9C", "Particulars", "GSTR 9", "GSTR 9C", "Remarks"]
    for c, h in enumerate(heads, 1):
        x = ws.Cells(4, c); x.Value = h
        x.Interior.Color = DARK; x.Font.Color = 0xFFFFFF; x.Font.Bold = True
        x.HorizontalAlignment = -4108
    r = 4
    for desc, part, sheet in ROWS:
        r += 1
        ws.Cells(r, 1).Value = "Sales"
        ws.Cells(r, 2).Value = desc
        ws.Cells(r, 7).Value = part
        ws.Hyperlinks.Add(Anchor=ws.Cells(r, 7), Address="", SubAddress="'%s'!A1" % sheet, TextToDisplay=part)
        for c in range(1, 11):
            ws.Cells(r, c).Borders.LineStyle = 1
    for c in range(1, 11):
        ws.Cells(4, c).Borders.LineStyle = 1
    for c, w in zip(range(1, 11), [11.89, 42.22, 9.33, 17.0, 8.11, 8.11, 25.0, 11.56, 11.56, 22.67]):
        ws.Columns(c).ColumnWidth = w
    # verify every link resolves
    bad = 0
    for hl in ws.Hyperlinks:
        try: xl.Application.Range(hl.SubAddress)
        except Exception: bad += 1; print("BROKEN:", hl.SubAddress)
    print("rows:", len(ROWS), "| broken links:", bad)
    wb.Close(SaveChanges=True)
finally:
    xl.Quit()
