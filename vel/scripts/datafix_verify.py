"""Differential verification: every dependent sheet must compute IDENTICAL values before (snapshot)
and after (rebuilt master). Plus zero error cells workbook-wide, plus sanity of the 3 data sheets."""
import win32com.client as win32, pythoncom, pickle, os, sys
SP = os.path.dirname(os.path.abspath(__file__))
NEW = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
OLD = os.path.join(SP, "master2_snapshot_before_datafix.xlsx")
DEP = ["S2 Month-on-Month", "S2 SR vs GSTR-1", "S3 Month-on-Month", "S3 1 vs 3B", "S3 SR vs 3B", "S3 SR vs 3B MoM",
       "S3 Amendment Check", "S4 GL vs SR", "S6 Advances Control", "S6 Month-on-Month", "Statewise RCM vs 3B",
       "Month wise RCM vs 3B", "ITCR vs 3B Net ITC", "RCM Paid vs ITC Claimed", "S9 Sales Reco", "S2 Exceptions", "S4 Exceptions"]
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
def dump(path):
    wb = xl.Workbooks.Open(path, ReadOnly=True); xl.CalculateFullRebuild()
    out = {}; errs = {}
    for w in wb.Worksheets:
        try: e = w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: e = 0
        if e: errs[w.Name] = e
    for nm in DEP:
        try: w = wb.Worksheets(nm)
        except Exception: continue
        vals = w.UsedRange.Value
        out[nm] = vals
    wb.Close(False)
    return out, errs
try:
    o, oe = dump(OLD); print("snapshot error cells:", oe)
    n, ne = dump(NEW); print("NEW error cells:", ne)
    tot_diff = 0
    for nm in DEP:
        a, b = o.get(nm), n.get(nm)
        if a is None or b is None: print("  %-26s missing in %s" % (nm, "old" if a is None else "new")); continue
        if not isinstance(a, tuple): a = ((a,),); b = ((b,),)
        ra, rb = len(a), len(b); ca, cb = len(a[0]), len(b[0])
        d = 0
        for i in range(min(ra, rb)):
            for j in range(min(ca, cb)):
                x, y = a[i][j], b[i][j]
                if isinstance(x, float) and isinstance(y, float):
                    if abs(x - y) > 0.005: d += 1
                elif x != y: d += 1
        shape_note = "" if (ra, ca) == (rb, cb) else " SHAPE %sx%s->%sx%s" % (ra, ca, rb, cb)
        tot_diff += d
        print("  %-26s cells %6d  diffs %d%s" % (nm, ra * ca, d, shape_note))
    print("TOTAL DIFFS:", tot_diff, "| NEW ERROR CELLS:", sum(ne.values()))
    wb = xl.Workbooks.Open(NEW, ReadOnly=True); xl.CalculateFullRebuild()
    for nm, spots in (("GSTR-1 Data", ["A2", "V2", "AG2", "AL2", "AM2", "AN2", "AO2", "AP2", "AM3", "AN3", "AO3"]),
                      ("3B Data", ["A2", "D2", "H2", "L2", "P2", "AE2", "AV2", "AW2", "AX2", "AV3", "AW3", "AX3"]),
                      ("GL Data", ["A2", "G2", "AE2", "AI2", "AJ2", "AK2", "AL2", "AM2", "AN2", "AO2", "AJ3", "AK3", "AL3", "AM3", "AN3", "AO3"])):
        w = wb.Worksheets(nm); print(nm, "used", w.UsedRange.Rows.Count, "x", w.UsedRange.Columns.Count, "shapes:", w.Shapes.Count)
        print("   ", [(s, str(w.Range(s).Value)[:30]) for s in spots])
    wb.Close(False)
finally:
    xl.Quit()
