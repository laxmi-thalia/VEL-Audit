"""Standard remark vocabulary on 'ITC Register 2025-26' and 'GSTR-2B Apr25-Aug26' (Pawan 21-09): sentence case,
'Matched with <sheet> – <basis>' / 'Not in <sheet> – <scope>', '– review' only on non-exact matches. GSTR-9 vocabulary
('Table 6A1 of GSTR-9 - Unclaimed/Claimed', Table 13 'Claimed/Unclaimed', 'Correction Entries- ...') is unchanged because
ITC Summary / T6A1 test those strings. Single COM session; 6A1 blocks and Net-ITC compared before/after; xlsb exported."""
import os, time, shutil, collections, win32com.client as win32, pythoncom
from openpyxl.utils import get_column_letter as L
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
for p in (P, B):
    try: open(p, "r+b").close()
    except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + p)
shutil.copy(P, "master2_snapshot_before_remarks.xlsx")
MAP = {"Matched (invoice no)": "Matched with 2B – invoice no",
       "Matched (GSTIN + amount within 100) - invoice no differs, review": "Matched with 2B – GSTIN + amount (±100), invoice no differs – review",
       "Matched (invoice similar + amount within 100) - invoice no differs, review": "Matched with 2B – similar invoice no + amount (±100) – review",
       "Matched (date+amount) - invoice no differs, review": "Matched with 2B – date + amount, invoice no differs – review",
       "Matched (amount, single candidate) - invoice no differs, review": "Matched with 2B – amount only, invoice no differs – review",
       "Matched in FY 24-25 2B (working files) - 6A1 component 1": "Matched with 2B of FY 24-25 – Table 6A1",
       "NOT FOUND IN 2B (Apr25-Aug26)": "Not in 2B – Apr-25 to Aug-26"}
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    isum = wb.Worksheets("ITC Summary"); snap = lambda: [round(isum.Cells(25, c).Value or 0, 2) for c in list(range(7, 22)) + [49, 50, 51]]
    before = snap()
    # ---- register
    sh = wb.Worksheets("ITC Register 2025-26"); H = {sh.Cells(5, c).Value: c for c in range(1, 90) if sh.Cells(5, c).Value}; RN = sh.Cells(sh.Rows.Count, 4).End(-4162).Row
    rr = sh.Range(sh.Cells(6, H["Reco Remarks"]), sh.Cells(RN, H["Reco Remarks"])); cat = [v[0] for v in sh.Range(sh.Cells(6, H["Category"]), sh.Cells(RN, H["Category"])).Value]
    old = [v[0] for v in rr.Value]; new = []; unknown = collections.Counter()
    for v, c in zip(old, cat):
        s = "" if v is None else str(v)
        if s in MAP: new.append([MAP[s]])
        elif s.startswith("No vendor GSTIN"): new.append(["Not applicable – RCM self-invoice" if c == "RCM" else ("Not applicable – ISD" if c == "ISD" else "Not applicable – no vendor GSTIN (URD)")])
        elif s.startswith("RCM self-invoice"): new.append(["Not applicable – RCM self-invoice"])
        else: new.append([v]); unknown[s[:40]] += 1
    rr.Value = new
    f9c = sh.Cells(6, H["GSTR 9C_Reporting"]).Formula.replace('"Check - dated "&$M6', '"Check – dated FY "&RIGHT($M6,5)&", time-bar review"')
    sh.Range(sh.Cells(6, H["GSTR 9C_Reporting"]), sh.Cells(RN, H["GSTR 9C_Reporting"])).Formula = f9c
    for col, a, b in (("Reasons", "Prior-year dated - time-bar review", "Prior-year dated – time-bar review"), ("Matching of 12B of FY 25-26 and 12C of 24-25", " - review", " – review")):
        f = sh.Cells(6, H[col]).Formula
        if a in f: sh.Range(sh.Cells(6, H[col]), sh.Cells(RN, H[col])).Formula = f.replace(a, b)
    # ---- 2B sheet
    b2 = wb.Worksheets("GSTR-2B Apr25-Aug26"); BH = {b2.Cells(2, c).Value: c for c in range(1, 70) if b2.Cells(2, c).Value}; NB = b2.Cells(b2.Rows.Count, 1).End(-4162).Row
    assert (BH["Reco Remarks"], BH["6A1 mark"], BH["Table 8A"], BH["GSTR-9/9C"], BH["3B Claim Month"], BH["FY (2B period)"], BH["Doc FY (doc date)"]) == (51, 52, 53, 58, 50, 47, 48), BH
    b2.Range("AY3:AY%d" % NB).Formula = '=IF($AX3="","Not in ITC Register – FY 25-26 claims","Matched with ITC Register – claimed "&TEXT($AX3,"mmm-yy"))'
    b2.Range("AZ3:AZ%d" % NB).Formula = '=IF($AV3="2024-25",IF($AX3="","Table 6A1 – FY 24-25 invoice – unclaimed",IF($AU3="2024-25","Table 6A1 – FY 24-25 invoice in 2B of FY 24-25 – claimed","Table 6A1 – FY 24-25 invoice in 2B of FY 25-26 – claimed")),"")'
    b2.Range("BA3:BA%d" % NB).Formula = '=IF($AU3<>"2025-26","No – 2B period FY "&RIGHT($AU3,5),IF($L3="Yes","No – RCM",IF($AH3="No","No – ITC not available",IF($S3="Yes","No – amendment",IF($AV3<>"2025-26","No – FY "&RIGHT($AV3,5)&" document","Yes")))))'
    b2.Range("BF3:BF%d" % NB).Formula = '=IF(LEFT($AZ3,9)="Table 6A1","Table 6A1 of GSTR-9 - "&IF(ISNUMBER(SEARCH("unclaimed",$AZ3)),"Unclaimed","Claimed"),IF($BC3<>"",$BC3,IF($BA3="Yes","Table 6B of GSTR-9","")))'
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    after = snap(); e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    print("register Reco Remarks:", dict(collections.Counter(v[0] for v in rr.Value).most_common(12)), "| unmapped:", dict(unknown))
    for col in ("Reco Remarks", "6A1 mark", "Table 8A", "GSTR-9/9C", "Final Remarks"):
        vals = [v[0] for v in b2.Range(b2.Cells(3, BH[col]), b2.Cells(NB, BH[col])).Value]; print("2B %s:" % col, dict(collections.Counter(str(v)[:44].replace("claimed Mar-26", "claimed <mon>") if v not in (None, "") else "<blank>" for v in vals).most_common(6)))
    print("ITC Summary 6A1 blocks + Net-ITC before == after:", before == after, "| error cells", e)
    assert e == 0 and before == after and not unknown
    wb.Save(); wb.SaveAs(B, FileFormat=50); wb.Close(False); print("saved + xlsb %.1f MB (%.0fs)" % (os.path.getsize(B) / 1e6, time.time() - t0))
finally: xl.Quit()
