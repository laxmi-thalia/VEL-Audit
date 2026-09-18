"""Dump last year's five ITC sheets from the local xlsb copy: header rows (values+styles), the first data row's
formulas (R1C1 + A1), column widths, freeze panes, merges, autofilter, number formats -> ly_itc_dump.pkl + readable txt."""
import win32com.client as win32, pythoncom, pickle, json
from openpyxl.utils import get_column_letter as L
pythoncom.CoInitialize(); xl = win32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
P = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\ly_9c.xlsb"
SHEETS = {"ITC Summary": (1, 4, 5, 23), "ITC Register 2024-25": (2, 5, 6, 8), "ITC Register 2025-26": (2, 4, 5, 7), "Tax comp report": (1, 7, 8, 10),
          "T6A1 Extract - 23-24": (1, 4, 5, 7), "Table 13 & 6A1 differences": (1, 6, 7, 9)}
out = {}; txt = []
def g(fn, default=None):
    try:
        v = fn(); return default if v is None else v
    except Exception: return default
def cellinfo(c):
    f = str(c.Formula) if c.Formula is not None else ""
    v = c.Value
    import datetime as _dt
    if hasattr(v, "year") and hasattr(v, "hour"): v = _dt.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second)
    return {"v": v, "f": f if f.startswith("=") else None, "r1c1": str(c.FormulaR1C1) if f.startswith("=") else None,
            "nf": g(lambda: c.NumberFormat, "General"), "bold": bool(g(lambda: c.Font.Bold, False)), "fc": int(g(lambda: c.Font.Color, 0)),
            "fill": None if g(lambda: c.Interior.ColorIndex, -4142) == -4142 else int(g(lambda: c.Interior.Color, 0)),
            "wrap": bool(g(lambda: c.WrapText, False)), "ha": int(g(lambda: c.HorizontalAlignment, 1)), "fs": float(g(lambda: c.Font.Size, 11)), "fn": g(lambda: c.Font.Name, "Calibri")}
try:
    wb = xl.Workbooks.Open(P, ReadOnly=True, UpdateLinks=0)
    for nm, (r_first, r_hdr, r_data, r_data_end) in SHEETS.items():
        ws = wb.Worksheets(nm); ur = ws.UsedRange; nc = ur.Columns.Count + ur.Column - 1; nr = ur.Rows.Count + ur.Row - 1
        d = {"ncols": nc, "nrows": nr, "freeze": None, "merges": [], "widths": {}, "rowh": {}, "hdr": {}, "data": {}, "autofilter": None}
        ws.Activate()
        try: d["freeze"] = (xl.ActiveWindow.SplitRow, xl.ActiveWindow.SplitColumn) if xl.ActiveWindow.FreezePanes else None
        except Exception: pass
        try: d["autofilter"] = str(ws.AutoFilter.Range.Address) if ws.AutoFilterMode else None
        except Exception: pass
        for c in range(1, nc + 1): d["widths"][c] = ws.Columns(c).ColumnWidth
        for r in range(1, r_data_end + 1): d["rowh"][r] = ws.Rows(r).RowHeight
        seen = set()
        for r in range(r_first, r_data_end + 1):
            for c in range(1, nc + 1):
                cell = ws.Cells(r, c)
                if cell.MergeCells:
                    a = str(cell.MergeArea.Address).replace("$", "")
                    if a not in seen: seen.add(a); d["merges"].append(a)
                info = cellinfo(cell)
                if r < r_data: d["hdr"][(r, c)] = info
                else: d["data"][(r, c)] = info
        out[nm] = d
        txt.append("\n\n########## %s  (%d cols, %d rows) freeze=%s filter=%s\nmerges=%s" % (nm, nc, nr, d["freeze"], d["autofilter"], d["merges"][:60]))
        for r in range(r_first, r_data + 1):
            txt.append("--- row %d (h=%.1f)" % (r, d["rowh"].get(r, 0)))
            for c in range(1, nc + 1):
                i = d["hdr"].get((r, c)) or d["data"].get((r, c))
                if i and (i["v"] not in (None, "") or i["f"]):
                    txt.append("  %-4s %-70s | f=%s | nf=%s bold=%s fill=%s fc=%s" % (L(c), str(i["v"])[:70].replace("\n", " "), (i["f"] or "")[:260], i["nf"], i["bold"], i["fill"], i["fc"]))
        print(nm, "dumped", nc, "cols", flush=True)
    wb.Close(False)
finally: xl.Quit()
pickle.dump(out, open(P.replace("ly_9c.xlsb", "ly_itc_dump.pkl"), "wb"))
open(P.replace("ly_9c.xlsb", "ly_itc_dump.txt"), "w", encoding="utf-8").write("\n".join(txt))
print("done")
