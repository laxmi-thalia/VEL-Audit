import win32com.client as win32, pythoncom
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
fails = []
def ck(name, got, exp, tol=0.05):
    ok = got is not None and abs(got - exp) <= tol
    if not ok: fails.append("%s: %r vs %r" % (name, got, exp))
    print("  %-46s %18s %s" % (name, format(got, ",.2f") if isinstance(got, (int, float)) else got, "OK" if ok else "*** FAIL exp %s" % exp))
try:
    wb = xl.Workbooks.Open(P); xl.CalculateFullRebuild()
    tot = 0
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        tot += e
        if e: print("ERRORS in", w.Name, e)
    print("formula ERROR cells:", tot); fails.extend(["errors %d" % tot] if tot else [])
    ws = wb.Worksheets("SR_2025-26")
    ck("register taxable R2", ws.Range("R2").Value, 8721471166.37)
    g = wb.Worksheets("Step 1 Flags")
    ck("gate blocking", float(g.Range("B4").Value), 3, 0)
    s2 = wb.Worksheets("S2 Reco (CA format)")
    ck("S2 SR total", s2.Range("AA50").Value, 8721471166.37)
    ck("S2 G1 total", s2.Range("AA26").Value, 8704594490.91)
    ck("S2 diff", s2.Range("AA74").Value, 16876675.46)
    mm = wb.Worksheets("S2 Month-on-Month")
    resid = xl.WorksheetFunction.Sum(mm.Range("M4:M1371"))
    ck("S2 MoM residual sum", resid, -9.42, 0.5)
    s3 = wb.Worksheets("S3 1 vs 3B")
    last = s3.Cells(s3.Rows.Count, 2).End(-4162).Row
    ck("S3 G1 total", s3.Cells(last, 3).Value, 8704594490.91)
    ck("S3 3B total", s3.Cells(last, 4).Value, 8669828675.34)
    ck("S3 diff", s3.Cells(last, 5).Value, 34765815.57)
    sv = wb.Worksheets("S3 SR vs 3B")
    lastv = sv.Cells(sv.Rows.Count, 2).End(-4162).Row
    ck("S3 SRvs3B diff", sv.Cells(lastv, 5).Value, 51642491.03)
    s4 = wb.Worksheets("S4 GL vs SR")
    l4 = s4.Cells(s4.Rows.Count, 2).End(-4162).Row
    ck("S4 GL CGST", s4.Cells(l4, 3).Value, 610374814.40)
    ck("S4 diff CGST", s4.Cells(l4, 5).Value, 470390.34)
    ck("S4 residual", s4.Cells(l4, 10).Value, -0.66, 0.2)
    ck("S4 IGST diff", s4.Cells(l4, 8).Value, 0.0)
    h5 = wb.Worksheets("S5 HSN Summary")
    l5 = h5.Cells(h5.Rows.Count, 1).End(-4162).Row
    ck("S5 HSN taxable", h5.Cells(l5, 7).Value, 8486105710.82)
    ck("S5 HSN qty", h5.Cells(l5, 6).Value, 53395416.32)
    r5 = wb.Worksheets("S5 Rate-wise")
    lr5 = r5.Cells(r5.Rows.Count, 1).End(-4162).Row
    ck("S5 rate taxable", r5.Cells(lr5, 4).Value, 8486105710.82)
    a6 = wb.Worksheets("S6 Advances Control")
    l6 = a6.Cells(a6.Rows.Count, 1).End(-4162).Row
    ck("S6 opening", a6.Cells(l6, 3).Value, 197901651.69)
    ck("S6 received", a6.Cells(l6, 4).Value, 381672455.59)
    ck("S6 adjusted", a6.Cells(l6, 5).Value, 146306999.99)
    ck("S6 closing", a6.Cells(l6, 6).Value, 433267107.29)
    ck("S6 books vs G1", a6.Cells(l6, 9).Value, -9.49, 0.2)
    s7 = wb.Worksheets("S7 CN Time-bar")
    l7 = s7.Cells(s7.Rows.Count, 2).End(-4162).Row
    beyond = sum(1 for r in range(4, l7 + 1) if "BEYOND" in str(s7.Cells(r, 8).Value or ""))
    untest = sum(1 for r in range(4, l7 + 1) if "not captured" in str(s7.Cells(r, 8).Value or ""))
    ck("S7 beyond-window CNs", float(beyond), 3, 0)
    ck("S7 untestable CNs", float(untest), 280, 0)
    s9 = wb.Worksheets("S9 Sales Reco")
    for r in range(6, 20):
        if str(s9.Cells(r, 1).Value or "") == "Difference (A)":
            ck("S9 Difference (A) total", s9.Cells(r, 21).Value, 3771605598.21, 1); break
    op = wb.Worksheets("Open Points (all)")
    ck("Open Points rows", float(op.Cells(op.Rows.Count, 1).End(-4162).Row - 1), 38, 0)
    wb.Close(SaveChanges=True)
    print("\nVERDICT:", "ALL CHECKS PASS" if not fails else "FAILURES:\n" + "\n".join(fails))
finally:
    xl.Quit()
