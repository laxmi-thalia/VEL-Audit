"""Birds Eye View VEL GSTR 3B FY 25-26 (TPEV format, standalone file):
- one sheet per GSTIN: Octa 'GSTR-3B' matrix (Section | Type | Total | 12 months)
- PAN sheet: consolidated matrix (sum across GSTINs)
- 'VEL Liability Summary': per-GSTIN grid, LIVE SUMIFS over the state sheets."""
import os, re, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
B = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-3B/"
OUT = r"C:\Users\pawar\Downloads\Birds Eye View VEL GSTR 3B FY25-26.xlsx"
CODE = {"Jammu and Kashmir":"JK","Punjab":"PB","Rajasthan":"RJ","Uttar Pradesh":"UP","Bihar":"BH",
"Arunachal Pradesh":"AR","Assam":"AS","West Bengal":"WB","Jharkhand":"JH","Chhattisgarh":"CG",
"Madhya Pradesh":"MP","Gujarat":"GJ","Maharashtra":"MH","Karnataka":"KA","Kerala":"KL",
"Tamil Nadu":"TN","Telangana":"TG","Andhra Pradesh":"AP"}
states = []          # (code, state, gstin, df rows [section, type, m1..m12])
months_lbl = None
for f in sorted(os.listdir(B)):
    if not f.lower().endswith((".xlsx", ".xls")): continue
    st = re.sub(r"-2025-26.*$", "", f.split("LIMITED-")[1])
    ov = pd.read_excel(B + f, sheet_name="Overview", header=None)
    gstin = None
    for i in range(len(ov)):
        if str(ov.iloc[i, 1]).strip().lower() == "gstin": gstin = str(ov.iloc[i, 2]).split("(")[0].strip()
    d = pd.read_excel(B + f, sheet_name="GSTR-3B", header=None)
    mlbl = [S(d.iloc[0, j]) for j in range(3, 15)]
    if months_lbl is None: months_lbl = mlbl
    rows = []
    for i in range(1, len(d)):
        sec = S(d.iloc[i, 1]); typ = S(d.iloc[i, 2])
        if not sec: continue
        vals = [pd.to_numeric(d.iloc[i, j], errors="coerce") for j in range(3, 15)]
        vals = [0.0 if pd.isna(v) else float(v) for v in vals]
        rows.append([sec, typ] + vals)
    states.append((CODE.get(st, st[:3].upper()), st, gstin, rows))
    print("  %-22s %s rows %d" % (st, gstin, len(rows)))
# canonical row order = union preserving first file's order
canon = []
seen = set()
for _, _, _, rows in states:
    for r in rows:
        k = (r[0], r[1])
        if k not in seen: seen.add(k); canon.append(k)
print("canonical section/type rows:", len(canon), "| months:", months_lbl)
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); TOT = Font(bold=True); NF = "#,##0.00"
wb = openpyxl.Workbook()
wb.remove(wb.active)
def write_matrix(ws, title_pairs, rows_map, live_total=True):
    r = 1
    for k, v in title_pairs:
        ws.cell(r, 1, k).font = TOT; ws.cell(r, 2, v); r += 1
    r += 1
    hdr = ["Section", "Type", "Total"] + months_lbl
    hr = r
    for c, h in enumerate(hdr, 1):
        x = ws.cell(r, c, h); x.font = HF; x.fill = HB
    for (sec, typ) in canon:
        r += 1
        ws.cell(r, 1, sec); ws.cell(r, 2, typ)
        vals = rows_map.get((sec, typ))
        for j in range(12):
            cell = ws.cell(r, 4 + j, round(vals[j], 2) if vals else 0.0)
            cell.number_format = NF
        ws.cell(r, 3).value = "=SUM(D%d:O%d)" % (r, r)
        ws.cell(r, 3).number_format = NF; ws.cell(r, 3).font = TOT
    ws.column_dimensions["A"].width = 52; ws.column_dimensions["B"].width = 16
    for c in range(3, 16): ws.column_dimensions[L(c)].width = 14
    ws.freeze_panes = "D%d" % (hr + 1)
    return hr
# per-state sheets
sheet_meta = {}
for code, st, gstin, rows in states:
    ws = wb.create_sheet(code)
    rows_map = {(r[0], r[1]): r[2:] for r in rows}
    hr = write_matrix(ws, [("GSTR3B Filed Report", ""), ("GSTIN", gstin), ("LEGAL NAME", "VIKRAN ENGINEERING LIMITED"),
                           ("STATE", st), ("RETURN PERIOD", "042025-032026")], rows_map)
    sheet_meta[code] = (st, gstin, hr)
# PAN consolidated
pan_map = {}
for _, _, _, rows in states:
    for r in rows:
        k = (r[0], r[1])
        cur = pan_map.setdefault(k, [0.0] * 12)
        for j in range(12): cur[j] += r[2 + j]
wsp = wb.create_sheet("PAN_AAECR0503Q", 0)
hrp = write_matrix(wsp, [("GSTR3B Filed Report", ""), ("PAN", "AAECR0503Q"), ("LEGAL NAME", "VIKRAN ENGINEERING LIMITED"),
                         ("RETURN PERIOD", "042025-032026"), ("TOTAL GSTIN", str(len(states)))], pan_map)
# Liability Summary (live SUMIFS over state sheets)
wss = wb.create_sheet("VEL Liability Summary", 1)
wss.cell(1, 1, "VEL - GSTR-3B Birds Eye View - Liability Summary FY 2025-26 (all figures live from the state sheets)").font = TOT
heads = ["Sr No.", "GSTIN", "Location", "Code",
 "GST paid Output (3.1a tax)", "RCM Liability Paid (3.1d tax)", "Net ITC Claimed (4A+4B)", "RCM ITC claimed (4A3)",
 "ISD ITC (4A4)", "Taxable Turnover (3.1a)", "Zero Rated Turnover (3.1b)", "Nil/Exempt Turnover (3.1c)",
 "Non-GST Turnover (3.1e)", "Remarks"]
for c, h in enumerate(heads, 1):
    x = wss.cell(3, c, h); x.font = HF; x.fill = HB; x.alignment = Alignment(wrap_text=True, horizontal="center")
TAX3 = '"Integrated Tax","Central Tax","State/UT Tax"'
def m3(code, hr, secpat, typ):
    rng = lambda col: "'%s'!$%s$%d:$%s$%d" % (code, col, hr + 1, col, hr + 1 + len(canon))
    return "SUMIFS(%s,%s,\"%s\",%s,\"%s\")" % (rng("C"), rng("A"), secpat, rng("B"), typ)
r = 3
for code, st, gstin, _rows in states:
    r += 1
    hr = sheet_meta[code][2]
    wss.cell(r, 1, r - 3); wss.cell(r, 2, gstin); wss.cell(r, 3, st); wss.cell(r, 4, code)
    def tax_sum(secpat):
        return "=" + "+".join(m3(code, hr, secpat, t) for t in ("Integrated Tax", "Central Tax", "State/UT Tax"))
    wss.cell(r, 5).value = tax_sum("3.1.A*")
    wss.cell(r, 6).value = tax_sum("3.1.D*")
    wss.cell(r, 7).value = "=" + "+".join(
        "+".join(m3(code, hr, p, t) for t in ("Integrated Tax", "Central Tax", "State/UT Tax"))
        for p in ("4.A.*", "4.B.*"))
    wss.cell(r, 8).value = tax_sum("4.A.3*")
    wss.cell(r, 9).value = tax_sum("4.A.4*")
    wss.cell(r, 10).value = "=" + m3(code, hr, "3.1.A*", "Supply Value")
    wss.cell(r, 11).value = "=" + m3(code, hr, "3.1.B*", "Supply Value")
    wss.cell(r, 12).value = "=" + m3(code, hr, "3.1.C*", "Supply Value")
    wss.cell(r, 13).value = "=" + m3(code, hr, "3.1.E*", "Supply Value")
    for c in range(5, 14): wss.cell(r, c).number_format = NF
last = r
r += 1
wss.cell(r, 3, "Total").font = TOT
for c in range(5, 14):
    wss.cell(r, c).value = "=SUM(%s4:%s%d)" % (L(c), L(c), last)
    wss.cell(r, c).number_format = NF; wss.cell(r, c).font = TOT
r += 2
wss.cell(r, 1, "NOTE: Haryana (06AAECR0503Q1ZC) has no GSTR-3B portal file in Portal Reports - not included.").font = Font(italic=True, color="808080")
for c, w in zip(range(1, 15), [7, 18, 20, 7, 17, 17, 17, 15, 14, 17, 16, 16, 15, 30]):
    wss.column_dimensions[L(c)].width = w
wss.freeze_panes = "E4"
wb.save(OUT)
print("saved:", OUT)
