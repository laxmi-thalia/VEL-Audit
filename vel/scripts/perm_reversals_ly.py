"""Stamp last year's 'Permanent Reversals' flags (VEL_GSTR 9_9C FY 24-25.xlsb, sheet 'GSTR-2B Apr 24-Oct 25') into
'GSTR-2B Apr25-Aug26' col 'Permanent Reversals' and into 'GSTR-2B ITC Data' (new col after the KEY helper). Key = my GSTIN +
supplier GSTIN + zero-insensitive doc no (ITC Data has no My GSTIN column -> supplier + doc no). Values only (the flag is a CA
decision from last year); ITC Summary 8C block (CE:CG) delta reported. xlsb export skipped if the xlsb is open in Excel."""
import os, re, time, shutil, collections, win32com.client as win32, pythoncom
from pyxlsb import open_workbook
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"; B = P[:-5] + ".xlsb"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("LOCKED - close in Excel without saving: " + P)
def xlsb_free():
    try: open(B, "r+b").close(); return True
    except PermissionError: return False
shutil.copy(P, "master2_snapshot_before_permrev.xlsx")
def S(v): return "" if v is None else str(v).strip()
def zkey(s): return re.sub(r"[ \-/.'_]", "", re.sub(r"(?<!\d)0+(?=\d)", "", S(s)).upper())
flags = set(); recl = set()
with open_workbook("VEL_2425.xlsb") as w:
    with w.get_sheet("GSTR-2B Apr 24-Oct 25") as sh:
        rows = [[c.v for c in r] for r in sh.rows()]
h = {S(x): j for j, x in enumerate(rows[3]) if S(x)}
for r in rows[4:]:
    if not r or len(r) <= h["Document Number"] or not S(r[h["My GSTIN"]]): continue
    k = (S(r[h["My GSTIN"]]).upper(), S(r[h["Supplier GSTIN"]]).upper(), zkey(r[h["Document Number"]]))
    if S(r[h["Permanent Reversals"]]): flags.add(k)
    if S(r[h["Reclaim - Table 6H"]]): recl.add(k)
print("LY flagged documents: permanent reversals %d | reclaim 6H %d" % (len(flags), len(recl)))
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135
    isum = wb.Worksheets("ITC Summary"); before = [round(isum.Cells(25, c).Value or 0, 2) for c in (83, 84, 85)]
    b2 = wb.Worksheets("GSTR-2B Apr25-Aug26"); BH = {b2.Cells(2, c).Value: c for c in range(1, 80) if b2.Cells(2, c).Value}; NB = b2.Cells(b2.Rows.Count, 1).End(-4162).Row
    col = lambda h_: [v[0] for v in b2.Range(b2.Cells(3, BH[h_]), b2.Cells(NB, BH[h_])).Value]
    my, sup, dn = col("Company GSTIN"), col("Supplier GSTIN"), col("Doc No"); n = 0; out = []
    for a, b, c in zip(my, sup, dn):
        hit = (S(a).upper(), S(b).upper(), zkey(c)) in flags; n += hit; out.append(["Permanent Reversals" if hit else None])
    b2.Range(b2.Cells(3, BH["Permanent Reversals"]), b2.Cells(NB, BH["Permanent Reversals"])).Value = out
    o = wb.Worksheets("GSTR-2B ITC Data"); OH = {o.Cells(5, c).Value: c for c in range(1, 40) if o.Cells(5, c).Value}; ON = o.Cells(o.Rows.Count, 1).End(-4162).Row
    kc = OH.get("Permanent Reversals (LY 9C)")
    if not kc:
        kc = max(OH.values()) + 1; x = o.Cells(5, kc); x.Value = "Permanent Reversals (LY 9C)"; x.Font.Bold = True; x.Font.Color = 0xFFFFFF; x.Interior.Color = 0xB09784; o.Columns(kc).ColumnWidth = 24
    flags2 = {(k[1], k[2]) for k in flags}
    sg = [v[0] for v in o.Range(o.Cells(6, OH["GSTIN of supplier"]), o.Cells(ON, OH["GSTIN of supplier"])).Value]; iv = [v[0] for v in o.Range(o.Cells(6, OH["Invoice number"]), o.Cells(ON, OH["Invoice number"])).Value]
    out2 = [["Permanent Reversals" if (S(a).upper(), zkey(b)) in flags2 else None] for a, b in zip(sg, iv)]; n2 = sum(1 for x in out2 if x[0])
    o.Range(o.Cells(6, kc), o.Cells(ON, kc)).Value = out2
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    after = [round(isum.Cells(25, c).Value or 0, 2) for c in (83, 84, 85)]; e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    reg = wb.Worksheets("ITC Register 2025-26"); H = {reg.Cells(5, c).Value: c for c in range(1, 80) if reg.Cells(5, c).Value}
    print("stamped: 2B Apr25-Aug26 %d rows | GSTR-2B ITC Data %d rows | ITC Summary CE:CG before %s after %s | error cells %d | golden %.2f" % (n, n2, before, after, e, reg.Cells(4, H["Total GST"]).Value))
    assert e == 0
    wb.Save()
    if xlsb_free(): wb.SaveAs(B, FileFormat=50); print("xlsb exported")
    else: print("xlsb OPEN in Excel - export skipped")
    wb.Close(False); print("saved (%.0fs)" % (time.time() - t0))
finally: xl.Quit()
