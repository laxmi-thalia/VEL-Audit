"""Re-merge the 2B base: FY 25-26 state files (Portal Reports\GSTR-2B, unchanged) + the REVISED FY 26-27 all-states file
(Audit data of FY 2026-27\GSTR-2B, Apr..Aug-26). Rebuilds 'GSTR-2B Apr25-Aug26' + '2B ISD Apr25-Aug26' with the same
column layout (45 raw Octa cols + 16 DPS cols) and re-writes every DPS-column formula; then re-points every formula in the
workbook that bounds a 2B range at the old last row to the new last row."""
import os, re, warnings, datetime as dt, pickle
from copy import copy
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
def S(v): return "" if v is None else str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
B1 = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-2B/"
B2 = "//192.168.1.69/gst folder/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Audit data of FY 2026-27/GSTR-2B/"
SP = os.path.dirname(os.path.abspath(__file__))
open(P, "r+b").close()
CAP = dt.datetime(2026, 8, 31)
def fy(d):
    if not isinstance(d, dt.datetime): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
files = [(B1, f) for f in sorted(os.listdir(B1)) if f.endswith(".xlsx") and not f.startswith("~$") and "Apr 2026" not in f]   # FY 25-26 state files only
files += [(B2, f) for f in sorted(os.listdir(B2)) if f.endswith(".xlsx") and not f.startswith("~$")]                            # revised FY 26-27 file
PN_H = ISD_H = None; pn_rows = []; isd_rows = []; hdr_font = hdr_fill = None
for folder, fn in files:
    wb = openpyxl.load_workbook(folder + fn, data_only=True)
    tag = fn.replace("GSTR2B-Net-VIKRAN ENGINEERING LIMITED-", "").replace("GSTR2B-VIKRAN ENGINEERING LIMITED-", "").replace(".xlsx", "") + (" [revised 18-09]" if folder == B2 else "")
    for sn, store in (("Purchase-Net", pn_rows), ("ISD", isd_rows)):
        if sn not in wb.sheetnames: print("  no", sn, "in", fn); continue
        ws = wb[sn]; H = [S(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
        while H and H[-1] == "": H.pop()
        if sn == "Purchase-Net":
            if PN_H is None: PN_H = H; hdr_font, hdr_fill = copy(ws.cell(1, 1).font), copy(ws.cell(1, 1).fill)
            elif H != PN_H: raise SystemExit("header drift in %s" % fn)
        else:
            if ISD_H is None: ISD_H = H
            elif H != ISD_H: raise SystemExit("ISD header drift in %s" % fn)
        n = 0
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or not r[0]: continue
            if isinstance(r[1], dt.datetime) and r[1] > CAP: continue
            store.append((tag, list(r[:len(H)]))); n += 1
        print("  %-52s %-12s %5d rows" % (tag[:52], sn, n))
    wb.close()
print("Purchase-Net rows:", len(pn_rows), "| ISD rows:", len(isd_rows))
iPer, iDate = PN_H.index("Tax Period"), PN_H.index("Doc Date")
OUR = ["Source file", "FY (2B period)", "Doc FY (doc date)", "KEY", "3B Claim Month", "Reco Remarks", "6A1 mark", "Table 8A", "8A Reco", "Table 8C", "Table 13", "Final Remarks", "GSTR-9/9C", "Permanent Reversals", "Reclaim - Table 6H", "Query"]
wb = openpyxl.load_workbook(P)
old = wb["GSTR-2B Apr25-Aug26"]; OLD_NB = max(r for r in range(3, old.max_row + 1) if old.cell(r, 1).value)
assert [S(old.cell(2, c).value) for c in range(1, len(PN_H) + len(OUR) + 1)] == PN_H + OUR, "layout changed - abort"
pos = wb.sheetnames.index("GSTR-2B Apr25-Aug26"); del wb["GSTR-2B Apr25-Aug26"]
ws = wb.create_sheet("GSTR-2B Apr25-Aug26", pos); ws.row_dimensions[1].height = 21
OF, OB = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="7F6000"); AMB = PatternFill("solid", fgColor="FFF2CC"); NF = "#,##0.00"
for c, h in enumerate(PN_H + OUR, 1):
    x = ws.cell(2, c, h)
    if c <= len(PN_H): x.font = copy(hdr_font); x.fill = copy(hdr_fill)
    else: x.font = OF; x.fill = OB
    x.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[2].height = 32
pn_rows.sort(key=lambda t: (t[1][iPer] if isinstance(t[1][iPer], dt.datetime) else dt.datetime(2099, 1, 1), S(t[1][0])))
r = 2; base = len(PN_H)
for tag, row in pn_rows:
    r += 1
    for c, v in enumerate(row, 1):
        cell = ws.cell(r, c, v)
        if isinstance(v, dt.datetime): cell.number_format = "DD-MM-YYYY"
        elif isinstance(v, float): cell.number_format = NF
    ws.cell(r, base + 1, tag); ws.cell(r, base + 2, fy(row[iPer])); ws.cell(r, base + 3, fy(row[iDate]))
NB = r
BH = {h: L(i + 1) for i, h in enumerate(PN_H + OUR)}
c = lambda h: BH[h]
reg = wb["ITC Register 2025-26"]; RH = {S(reg.cell(5, cc).value): L(cc) for cc in range(1, reg.max_column + 1)}; RN = max(rr for rr in range(6, reg.max_row + 1) if reg.cell(rr, 4).value)
RG = lambda h: "'ITC Register 2025-26'!$%s$6:$%s$%d" % (RH[h], RH[h], RN)
normf = lambda e: 'UPPER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(%s," ",""),"-",""),"/",""),".",""),"\'",""),"_",""))' % e
for r in range(3, NB + 1):
    ws.cell(r, base + 4).value = "=" + normf("%s%d&%s%d" % (c("Supplier GSTIN"), r, c("Doc No"), r))
    ws.cell(r, base + 5).value = '=IFERROR(INDEX(%s,MATCH($%s%d,%s,0)),"")' % (RG("3B Claim  Month"), c("KEY"), r, RG("KEY2 (matched 2B key)")); ws.cell(r, base + 5).number_format = "mmm-yy"
    ws.cell(r, base + 6).value = '=IF($%s%d="","NOT IN ITC REGISTER (FY 25-26 claims)","Matched with ITC Register - claimed "&TEXT($%s%d,"mmm-yy"))' % (c("3B Claim Month"), r, c("3B Claim Month"), r)
    ws.cell(r, base + 7).value = '=IF($%s%d="2024-25",IF($%s%d="","6A1 - Unclaimed (24-25 dated, no 3B claim month)",IF($%s%d="2024-25","6A1 - 24-25 inv in 24-25 2B, claimed 25-26","6A1 - 24-25 inv in 25-26 2B, claimed 25-26")),"")' % (c("Doc FY (doc date)"), r, c("3B Claim Month"), r, c("FY (2B period)"), r)
    ws.cell(r, base + 8).value = ('=IF($%s%d<>"2025-26","No-"&$%s%d,IF($%s%d="Yes","No-25-26 - RCM",IF($%s%d="No","No-25-26 - ITC not available",IF($%s%d="Yes","No-25-26 - Amendment",IF($%s%d<>"2025-26","No-"&$%s%d,"Yes")))))'
                                  % (c("FY (2B period)"), r, c("FY (2B period)"), r, c("Reverse Charge"), r, c("GSTR-9 (8A) ITC Available"), r, c("Is Amendment"), r, c("Doc FY (doc date)"), r, c("Doc FY (doc date)"), r))
    ws.cell(r, base + 9).value = '=IF($%s%d="Yes","To be considered","Ignore")' % (c("Table 8A"), r)
    ws.cell(r, base + 11).value = '=IF(AND($%s%d="2025-26",$%s%d="2026-27"),IF($%s%d="","Unclaimed","Claimed"),"")' % (c("Doc FY (doc date)"), r, c("FY (2B period)"), r, c("3B Claim Month"), r)
    ws.cell(r, base + 10).value = '=IF($%s%d="Claimed","Table 8C of GSTR-9","")' % (c("Table 13"), r)
    ws.cell(r, base + 12).value = '=IF($%s%d<>"",$%s%d,IF($%s%d<>"",$%s%d,$%s%d))' % (c("6A1 mark"), r, c("6A1 mark"), r, c("Table 8C"), r, c("Table 8C"), r, c("Reco Remarks"), r)
    ws.cell(r, base + 13).value = ('=IF(LEFT($%s%d,3)="6A1","Table 6A1 of GSTR-9 - "&IF(ISNUMBER(SEARCH("Unclaimed",$%s%d)),"Unclaimed","Claimed"),IF($%s%d<>"",$%s%d,IF($%s%d="Yes","Table 6B of GSTR-9","")))'
                                   % (c("6A1 mark"), r, c("6A1 mark"), r, c("Table 8C"), r, c("Table 8C"), r, c("Table 8A"), r))
for h in ("Permanent Reversals", "Reclaim - Table 6H", "Query"): ws.cell(2, base + 1 + OUR.index(h)).fill = AMB
NC = len(PN_H) + len(OUR)
for cc in range(1, NC + 1): ws.column_dimensions[L(cc)].width = 14
for cc, w in ((1, 18), (5, 18), (7, 18), (8, 28), (base + 1, 30), (base + 4, 30), (base + 6, 34)): ws.column_dimensions[L(cc)].width = w
ws.freeze_panes = "A3"; ws.auto_filter.ref = "A2:%s%d" % (L(NC), NB)
ws.cell(1, 3, "Octa GSTR-2B exports merged: 14 state files Apr-25..Mar-26 + the REVISED all-states FY 26-27 file (Apr..Aug-26, 18-09-2026). Raw Octa columns A..%s verbatim; DPS columns from %s (live). Missing FY 25-26 state files: Tamil Nadu, Telangana, West Bengal, Haryana; Kerala file empty." % (L(len(PN_H)), L(len(PN_H) + 1))).font = Font(italic=True, size=9, color="808080")
# ISD sheet
old2 = wb["2B ISD Apr25-Aug26"]; pos2 = wb.sheetnames.index("2B ISD Apr25-Aug26"); del wb["2B ISD Apr25-Aug26"]
ws2 = wb.create_sheet("2B ISD Apr25-Aug26", pos2); ws2.row_dimensions[1].height = 21
OUR2 = ["Source file", "FY (2B period)", "3B Claim Month", "Reco Remarks"]
for cc, h in enumerate(ISD_H + OUR2, 1):
    x = ws2.cell(2, cc, h); x.font = OF if cc > len(ISD_H) else Font(bold=True, color="FFFFFF"); x.fill = OB if cc > len(ISD_H) else PatternFill("solid", fgColor="1F4E79"); x.alignment = Alignment(wrap_text=True, vertical="center")
jPer = ISD_H.index("Tax Period"); isd_rows.sort(key=lambda t: (t[1][jPer] if isinstance(t[1][jPer], dt.datetime) else dt.datetime(2099, 1, 1), S(t[1][0])))
r = 2
for tag, row in isd_rows:
    r += 1
    for cc, v in enumerate(row, 1):
        cell = ws2.cell(r, cc, v)
        if isinstance(v, dt.datetime): cell.number_format = "DD-MM-YYYY"
        elif isinstance(v, float): cell.number_format = NF
    ws2.cell(r, len(ISD_H) + 1, tag); ws2.cell(r, len(ISD_H) + 2, fy(row[jPer]))
for cc in range(1, len(ISD_H) + len(OUR2) + 1): ws2.column_dimensions[L(cc)].width = 14
ws2.freeze_panes = "A3"; ws2.auto_filter.ref = "A2:%s%d" % (L(len(ISD_H) + len(OUR2)), r)
# re-point every 2B range bound in the workbook: $OLD_NB -> $NB (only inside references to this sheet)
pat = re.compile(r"(Apr25-Aug26'!\$[A-Z]{1,3}\$3:\$[A-Z]{1,3}\$)%d(?!\d)" % OLD_NB); n = 0
for nm in wb.sheetnames:
    if nm in ("GSTR-2B Apr25-Aug26", "2B ISD Apr25-Aug26"): continue
    for row in wb[nm].iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "Apr25-Aug26" in v:
                nv = pat.sub(lambda m: m.group(1) + str(NB), v)
                if nv != v: cell.value = nv; n += 1
print("2B sheet rows 3..%d (was ..%d) | dependent formulas re-bounded: %d" % (NB, OLD_NB, n))
wb.save(P); pickle.dump({"NB": NB, "OLD_NB": OLD_NB}, open(os.path.join(SP, "b2_remerge_meta.pkl"), "wb")); print("saved")
