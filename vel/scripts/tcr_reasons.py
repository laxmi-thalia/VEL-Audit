"""Fill 'Tax comp report' Q (Reasons), R (Impact), S-U / V-X / Y-AA / AB-AD bucket amounts and AI (9C note, Matched rows only)
from the deterministic decomposition in tcr_reasons_lib (Computation sheets vs filed 3B). Values only; the sheet's own
AE-AG / AH formulas stay. Snapshot first; COM; recalc; 0-error scan; per-row check that S+V+Y+AB + AE == L per head."""
import time, shutil, collections, win32com.client as win32, pythoncom
from tcr_reasons_lib import load, decompose, nz
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
try: open(P, "r+b").close()
except PermissionError: raise SystemExit("MASTER IS OPEN IN EXCEL - close it (without saving) and rerun")
shutil.copy(P, "master2_snapshot_before_tcr.xlsx")
tcr, filed, comp, missing = load()
COLS = {"S": 19, "V": 22, "Y": 25, "AB": 28}
rows_out = []
for r in tcr:
    d = decompose(r, filed.get((r["gstin"], r["month"])), comp.get((r["gstin"], r["month"]))); r["_d"] = d
    rows_out.append(d)
print("outcome:", dict(collections.Counter(d["reason"] if d["reason"] == "Matched" else ("no computation" if "no computation" in d["flags"] else "explained") for d in rows_out)),
      "| impact:", dict(collections.Counter(d["impact"] for d in rows_out)), "| residual flags:", sum(1 for d in rows_out if "residual" in d["flags"]))
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    t0 = time.time(); wb = xl.Workbooks.Open(P); xl.Calculation = -4135; ws = wb.Worksheets("Tax comp report")
    assert ws.Cells(7, 17).Value == "Reasons" and ws.Cells(7, 18).Value == "Impact" and ws.Cells(6, 19).Value.startswith("Difference to be added") and ws.Cells(7, 35).Value.startswith("Note to be added"), "header layout changed"
    for r, d in zip(tcr, rows_out):
        assert ws.Cells(r["row"], 2).Value == r["gstin"] and ws.Cells(r["row"], 3).Value == r["month"], ("row moved", r["row"])
        ws.Cells(r["row"], 17).Value = d["reason"]; ws.Cells(r["row"], 18).Value = d["impact"]
        ws.Cells(r["row"], 35).Value = d.get("note") or None
        for b, c0 in COLS.items():
            vec = d["buckets"][b] if d.get("buckets") else [0, 0, 0]
            for j in range(3): ws.Cells(r["row"], c0 + j).Value = round(vec[j], 2) if abs(vec[j]) > 0.005 else None
    ws.Range("Q8:Q211").WrapText = True; ws.Columns(17).ColumnWidth = 70; ws.Range("Q8:R211").VerticalAlignment = -4160
    xl.Calculation = -4105; xl.CalculateFullRebuild()
    e = 0
    for w in wb.Worksheets:
        try: e += w.UsedRange.SpecialCells(-4123, 16).Count
        except Exception: pass
    # per-row closure of the decomposition: S + V + Y + AB == L (sheet) within 1.5 for every explained row
    bad = 0; ab_rows = 0
    for r, d in zip(tcr, rows_out):
        if d["reason"] == "Matched" or "no computation" in d["flags"] or "no filed" in d["flags"]: continue
        L_ = [ws.Cells(r["row"], 12 + j).Value or 0 for j in range(3)]
        S_ = [sum((ws.Cells(r["row"], c0 + j).Value or 0) for c0 in (19, 22, 25, 28)) for j in range(3)]
        if any(abs(L_[j] - S_[j]) > 1.5 for j in range(3)): bad += 1; print("   closure gap:", r["state"], r["month"], L_, S_)
        if any(ws.Cells(r["row"], 28 + j).Value for j in range(3)): ab_rows += 1
    print("error cells %d | explained rows where S+V+Y+AB != L: %d | rows with CN bucket (AB): %d | column totals S/V/Y/AB: %s (%.0fs)" % (
        e, bad, ab_rows, [round(ws.Cells(5, c).Value or 0, 2) for c in (19, 22, 25, 28)], time.time() - t0))
    print("AE-AG residual totals (unexplained):", [round(ws.Cells(5, c).Value or 0, 2) for c in (31, 32, 33)])
    wb.Save(); wb.Close(False)
finally: xl.Quit()
print("saved")
