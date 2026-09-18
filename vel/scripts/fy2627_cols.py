"""ITC Register 2026-27: fill the reporting / 2B columns on the rebuilt rows (single COM session, then xlsb).
Conventions from the previous rows: GSTR 9_Reporting '13', GSTR 9C_Reporting '12C', Reasons text, Consider in 8A reco 'No',
Correct GSTIN = vendor GSTIN, 3B Claim Month as text '01 Apr 2026'; 2B columns LIVE against 'GSTR-2B Apr25-Aug26' (KEY col AW)."""
import os, time, collections, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for p in (P, B): open(p, "r+b").close()
LBL = {4: "01 Apr", 5: "02 May", 6: "03 June", 7: "04 July", 8: "05 Aug", 9: "06 Sep", 10: "07 Oct", 11: "08 Nov", 12: "09 Dec", 1: "10 Jan", 2: "11 Feb", 3: "12 Mar"}
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; sh = wb.Worksheets("ITC Register 2026-27")
    H = {sh.Cells(4, c).Value: c for c in range(1, 80) if sh.Cells(4, c).Value}; RN = sh.Cells(sh.Rows.Count, 4).End(-4162).Row; N = RN - 4
    c = lambda h: L(H[h]); NB = 13869
    rng = lambda h: sh.Range("%s5:%s%d" % (c(h), c(h), RN))
    key = 'UPPER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE($%s5&$%s5," ",""),"-",""),"/",""),".",""),"\'",""),"_",""))' % (c("Vendor GSTIN"), c("Invoice No."))
    m = "MATCH(%s,'GSTR-2B Apr25-Aug26'!$AW$3:$AW$%d,0)" % (key, NB)
    rng("Available in 2B").Formula = '=IF($%s5="Missing","N",IF(ISNUMBER(%s),"Y","N"))' % (c("Vendor GSTIN"), m)
    rng("2B Inv").Formula = '=IF($%s5="Y",INDEX(\'GSTR-2B Apr25-Aug26\'!$E$3:$E$%d,%s),"")' % (c("Available in 2B"), NB, m)
    rng("2B Period").Formula = '=IF($%s5="Y",INDEX(\'GSTR-2B Apr25-Aug26\'!$B$3:$B$%d,%s),"")' % (c("Available in 2B"), NB, m); rng("2B Period").NumberFormat = "mmm-yy"
    rng("2B YEAR").Formula = '=IF($%s5="","",IF(MONTH($%s5)>=4,YEAR($%s5)&"-"&RIGHT(YEAR($%s5)+1,2),YEAR($%s5)-1&"-"&RIGHT(YEAR($%s5),2)))' % ((c("2B Period"),) * 6)
    rng("Final Remarks").Formula = '=IF($%s5="Y",IF($%s5="2026-27","Matched A","Matched - in 2B of "&$%s5),"Not in 2B (Apr-25 to Aug-26)")' % (c("Available in 2B"), c("2B YEAR"), c("2B YEAR"))
    rng("Correct GSTIN").Formula = '=IF($%s5="Missing","",$%s5)' % (c("Vendor GSTIN"), c("Vendor GSTIN"))
    rng("GSTR 9_Reporting").Value = "13"; rng("GSTR 9C_Reporting").Value = "12C"; rng("Reasons").Value = "ITC booked in 2025-26 claimed in 2026-27"; rng("Consider in 8A reco").Value = "No"
    cm = rng("3B Claim  Month"); vals = cm.Value2; txt = []
    import datetime as dt
    for (v,) in vals:
        d = dt.datetime(1899, 12, 30) + dt.timedelta(days=int(v)) if isinstance(v, (int, float)) else None
        txt.append(["%s %d" % (LBL[d.month], d.year) if d else v])
    cm.NumberFormat = "@"; cm.Value = txt
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    av = collections.Counter(v[0] for v in rng("Available in 2B").Value); fr = collections.Counter(str(v[0])[:34] for v in rng("Final Remarks").Value)
    isum = wb.Worksheets("ITC Summary"); print("ITC Summary CH6:", isum.Range("CH6").Formula[:120])
    print("rows %d | error cells %d | Available in 2B: %s | Final Remarks: %s" % (N, e, dict(av), dict(fr.most_common(4))))
    print("ITC Summary total row - Table 13 (CH:CJ) %s | 12C (CV:CX) %s" % ([round(isum.Cells(25, k).Value or 0, 2) for k in (86, 87, 88)], [round(isum.Cells(25, k).Value or 0, 2) for k in (100, 101, 102)]))
    assert e == 0, "errors - not saved"
    wb.Save(); wb.SaveAs(B, FileFormat=50); wb.Close(False); print("saved + xlsb %.1f MB (%.0fs)" % (os.path.getsize(B) / 1e6, time.time() - t0))
finally: xl.Quit()
