"""Final verification: recalc all six deliverables, zero error cells, headline figures vs expected,
and the cross-file ties (register totals reused by every later step)."""
import win32com.client as win32, pythoncom
D = r"C:\Users\pawar\Downloads"
FILES = {
    "S1": D + r"\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx",
    "S2": D + r"\VEL_Step2_Reco_GSTR1_DRAFT.xlsx",
    "S3": D + r"\VEL_Step3_Reco_GSTR1_vs_3B_DRAFT.xlsx",
    "S4": D + r"\VEL_Step4_Reco_GL_vs_SR_DRAFT.xlsx",
    "S5": D + r"\VEL_Step5_HSN_Rate_Summary_DRAFT.xlsx",
    "S6": D + r"\VEL_Step6_Advances_Control_DRAFT.xlsx",
}
pythoncom.CoInitialize()
xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
got = {}; fails = []
def ck(name, actual, expected, tol=0.02):
    ok = actual is not None and abs(actual - expected) <= tol
    if not ok: fails.append("%s: got %r expected %r" % (name, actual, expected))
    print("  %-52s %18s  %s" % (name, format(actual, ",.2f") if isinstance(actual, float) else actual, "OK" if ok else "*** FAIL vs %s" % expected))
    got[name] = actual
try:
    for tag, path in FILES.items():
        wb = xl.Workbooks.Open(path); xl.CalculateFullRebuild()
        err = 0
        for w in wb.Worksheets:
            try: err += w.UsedRange.SpecialCells(-4123, 16).Count
            except Exception: pass
        print("\n== %s  (%s)  formula-error cells: %d %s" % (tag, path.split("\\")[-1], err, "" if err == 0 else "*** FAIL"))
        if err: fails.append("%s error cells: %d" % (tag, err))
        if tag == "S1":
            ws = wb.Worksheets("SR_2025-26")
            ck("register taxable (R2 SUBTOTAL)", ws.Range("R2").Value, 8721471166.37)
            ck("register CGST (T2)", ws.Range("T2").Value, 609904424.06)
            ck("register quantity (AJ2)", ws.Range("AJ2").Value, 53395416.32)
            g = wb.Worksheets("Step 1 Flags")
            ck("gate blocking", float(g.Range("B4").Value), 3, 0)
            ck("gate review", float(g.Range("B5").Value), 152, 0)
            listed = sum(1 for r in range(22, 422) if g.Cells(r, 2).Value not in (None, ""))
            ck("live flag list rows", float(listed), 154, 0)
        if tag == "S2":
            ws = wb.Worksheets("SR vs GSTR-1")
            ck("S2 SR total liability", ws.Range("AA50").Value, 8721471166.37)
            ck("S2 GSTR-1 total liability", ws.Range("AA26").Value, 8704594490.91)
            ck("S2 difference", ws.Range("AA74").Value, 16876675.46)
            ck("S2 explained (auto)", ws.Range("AF74").Value, 16876684.90, 5)
            ck("S2 unexplained residual", ws.Range("AH74").Value, -9.42, 0.5)
        if tag == "S3":
            ws = wb.Worksheets("GSTR-1 vs GSTR-3B")
            ck("S3 GSTR-1 total", ws.Range("C50").Value, 8704594490.91)
            ck("S3 3B 3.1(a) total", ws.Range("G50").Value, 8669828675.34)
            ck("S3 difference", ws.Range("K50").Value, 34765815.57)
            ex = wb.Worksheets("Exceptions")
            last = ex.Cells(ex.Rows.Count, 1).End(-4162).Row
            res = next((ex.Cells(r, 2).Value for r in range(last, 1, -1) if str(ex.Cells(r, 1).Value or "") == "Unexplained residual"), None)
            ck("S3 tie-out residual", res, -0.23, 0.05)
        if tag == "S4":
            ws = wb.Worksheets("GL vs Sales Register")
            ck("S4 GL CGST", ws.Cells(23, 3).Value, 610374814.40)
            ck("S4 register CGST", ws.Cells(23, 7).Value, 609904424.06)
            ck("S4 CGST diff", ws.Cells(23, 11).Value, 470390.34)
            ck("S4 IGST diff", ws.Cells(23, 13).Value, 0.0)
            ck("S4 residual", ws.Cells(23, 16).Value, -0.66, 0.1)
        if tag == "S5":
            hs = wb.Worksheets("HSN Summary")
            last = hs.Cells(hs.Rows.Count, 1).End(-4162).Row
            ck("S5 HSN taxable", hs.Cells(last, 8).Value, 8486105710.82)
            ck("S5 HSN quantity", hs.Cells(last, 7).Value, 53395416.32)
            ck("S5 HSN CGST", hs.Cells(last, 10).Value, 588721533.06)
            rw = wb.Worksheets("Rate-wise Summary")
            lr = rw.Cells(rw.Rows.Count, 1).End(-4162).Row
            ck("S5 rate-wise taxable", rw.Cells(lr, 4).Value, 8486105710.82)
        if tag == "S6":
            ca = wb.Worksheets("Control Account")
            ck("S6 opening", ca.Cells(24, 3).Value, 197901651.69)
            ck("S6 received", ca.Cells(24, 4).Value, 381672455.59)
            ck("S6 adjusted", ca.Cells(24, 5).Value, 146306999.99)
            ck("S6 closing", ca.Cells(24, 6).Value, 433267107.29)
            ck("S6 books-vs-GSTR-1 diff", ca.Cells(24, 9).Value, -9.49, 0.1)
            g = wb.Worksheets("GL Advance Check")
            bad = sum(1 for r in range(4, 12) if abs(g.Cells(r, 5).Value or 0) >= 1)
            ck("S6 ledger-vs-books mismatched states", float(bad), 0, 0)
        wb.Close(SaveChanges=False)
    print("\n---- cross-file ties ----")
    ck("S2 SR total == register taxable", got["S2 SR total liability"], got["register taxable (R2 SUBTOTAL)"])
    ck("S3 GSTR-1 == S2 GSTR-1", got["S3 GSTR-1 total"], got["S2 GSTR-1 total liability"])
    ck("S4 register CGST == register CGST", got["S4 register CGST"], got["register CGST (T2)"])
    ck("S5 taxable == register - advances", got["S5 HSN taxable"], got["register taxable (R2 SUBTOTAL)"] - 235365455.56, 0.1)
    ck("S6 received-adjusted == advances net", got["S6 received"] - got["S6 adjusted"], 235365455.56, 0.1)
    print("\n==================== VERDICT:", "ALL CHECKS PASS" if not fails else "FAILURES:\n" + "\n".join(fails))
finally:
    xl.Quit()
