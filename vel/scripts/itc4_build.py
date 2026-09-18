"""ITC step 4 + October list:
A. 'ITCR vs 2B GSTN Level' - per GSTIN live: register ITC claimed vs 2B FY25-26 available,
   diff (Less in 2B), unclaimed-in-2B value, DPS Remarks.
B. 'Unclaimed ITC candidates' - 2B rows FY24-25/25-26 with NO claim in this year's register
   and (for 24-25) none in last year's either -> the claim-by-Oct/Nov-2026 review list."""
import pandas as pd, re, datetime
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()
def ninv(v): return re.sub(r"[^A-Z0-9]", "", S(v).upper().lstrip("'"))
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f = open(P, 'r+b'); f.close()
IG = "Integrated Tax(\u20b9)"; CG = "Central Tax(\u20b9)"; SG = "State/UT Tax(\u20b9)"
TXV = "Taxable Value (\u20b9)"
b2 = pd.read_pickle("b2itc.pkl").reset_index(drop=True)
ly = pd.read_pickle("ly_claims.pkl")
b2["vg2"] = b2["GSTIN of supplier"].map(lambda v: S(v).upper())
b2["inv2"] = b2["Invoice number"].map(ninv)
b2["tot2"] = (b2[IG].fillna(0) + b2[CG].fillna(0) + b2[SG].fillna(0)).round(2)
uncl = b2[(b2["claimed_by_reg"] == "") & (b2["FY (derived)"].isin(["2024-25", "2025-26"]))].copy()
uncl["in_ly"] = [(g, i) in ly for g, i in zip(uncl["vg2"], uncl["inv2"])]
cand = uncl[~((uncl["FY (derived)"] == "2024-25") & (uncl["in_ly"]))].copy()
def status(r):
    if r["FY (derived)"] == "2025-26":
        return "CANDIDATE - claimable till Nov 2026 (Sec 16(4))"
    return "FY 24-25 - not found in either register; verify (window closing/closed)"
cand["Status"] = cand.apply(status, axis=1)
cand = cand[cand["tot2"] > 0]
print("unclaimed 2B rows considered:", len(uncl), "| after last-year-claims filter:", len(cand))
for fy in ("2024-25", "2025-26"):
    s = cand[cand["FY (derived)"] == fy]
    print("  %s candidates: %d rows | tax %.2f" % (fy, len(s), s["tot2"].sum()))
FOLDER2G = {"Andhra Pradesh": "37AAECR0503Q1Z7", "Arunachal Pradesh": "12AAECR0503Q1ZJ", "Assam": "18AAECR0503Q1Z7",
"Bihar": "10AAECR0503Q1ZN", "Chattisgarh": "22AAECR0503Q1ZI", "Gujarat": "24AAECR0503Q1ZE",
"Jammu Kashmir": "01AAECR0503Q1ZM", "Jharkhand": "20AAECR0503Q1ZM", "Karnataka": "29AAECR0503Q1Z4",
"Kerala": "32AAECR0503Q1ZH", "Madhya Pradesh": "23AAECR0503Q1ZG", "Maharashtra": "27AAECR0503Q1Z8",
"Punjab": "03AAECR0503Q1ZI", "Rajasthan": "08AAECR0503Q1Z8", "Tamil Nadu": "33AAECR0503Q1ZF",
"Telangana": "36AAECR0503Q1Z9", "Uttar Pradesh": "09AAECR0503Q1Z6", "West Bengal": "19AAECR0503Q1Z5"}
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); TOT = Font(bold=True); NF = "#,##0.00"
wb = openpyxl.load_workbook(P)
rr = wb["ITC Register 2025-26"]
RH = {}
for c in range(1, rr.max_column + 1):
    h = S(rr.cell(5, c).value)
    if h and h not in RH: RH[h] = c
R0, R1 = 6, 6 + 41508 - 1
REG = lambda n: "'ITC Register 2025-26'!$%s$%d:$%s$%d" % (L(RH[n]), R0, L(RH[n]), R1)
n2 = 5 + 48802
B2S = lambda c: "'GSTR-2B ITC Data'!$%s$6:$%s$%d" % (c, c, n2)
NM = "ITCR vs 2B GSTN Level"
if NM in wb.sheetnames: del wb[NM]
pos = wb.sheetnames.index("ITCR vs 3B Net ITC") + 1
ws = wb.create_sheet(NM, pos)
ws.row_dimensions[1].height = 21
ws.cell(2, 1, "VIKRAN ENGINEERING LIMITED").font = TOT
ws.cell(3, 1, "ITC Register vs GSTR-2B at GSTIN level. Positive diff = claimed exceeds 2B ('Less in 2B' - vendor recovery review); '2B unclaimed' = credit available but not claimed. RCM/ISD claims excluded from register side (not in 2B B2B).").font = TOT
hdrs = ["State", "GSTIN",
 "Register ITC claimed IGST", "Register ITC claimed CGST", "Register ITC claimed SGST", "Register ITC claimed Total",
 "2B available FY25-26 IGST", "2B available FY25-26 CGST", "2B available FY25-26 SGST", "2B available FY25-26 Total",
 "Diff (claimed - 2B)", "2B unclaimed-in-register (tax)", "DPS Remarks"]
for c, h in enumerate(hdrs, 1):
    x = ws.cell(5, c, h); x.font = HF; x.fill = HB
r = 5
for folder, g_ in sorted(FOLDER2G.items()):
    r += 1
    st = {"Chattisgarh": "Chhattisgarh", "Jammu Kashmir": "Jammu & Kashmir"}.get(folder, folder)
    ws.cell(r, 1, st); ws.cell(r, 2, g_); ws.cell(r, 15, folder)
    for i, meas in enumerate(("IGST", "CGST", "SGST")):
        ws.cell(r, 3 + i).value = '=SUMIFS(%s,%s,$B%d,%s,"ITC")' % (REG(meas), REG("VEL GSTIN"), r, REG("Category"))
    ws.cell(r, 6).value = "=SUM(C%d:E%d)" % (r, r)
    for i, c2 in enumerate(("Q", "R", "S")):
        ws.cell(r, 7 + i).value = '=SUMIFS(%s,%s,$O%d,%s,"2025-26")' % (B2S(c2), B2S("A"), r, B2S("B"))
    ws.cell(r, 10).value = "=SUM(G%d:I%d)" % (r, r)
    ws.cell(r, 11).value = "=F%d-J%d" % (r, r)
    ws.cell(r, 12).value = ('=SUMIFS(%s,%s,$O%d,%s,"2025-26",%s,"")+SUMIFS(%s,%s,$O%d,%s,"2025-26",%s,"")+SUMIFS(%s,%s,$O%d,%s,"2025-26",%s,"")'
        % (B2S("Q"), B2S("A"), r, B2S("B"), B2S("Z"), B2S("R"), B2S("A"), r, B2S("B"), B2S("Z"), B2S("S"), B2S("A"), r, B2S("B"), B2S("Z")))
    for c in range(3, 13): ws.cell(r, c).number_format = NF
last = r
r += 1
ws.cell(r, 1, "Total").font = TOT
for c in range(3, 13):
    ws.cell(r, c).value = "=SUM(%s6:%s%d)" % (L(c), L(c), last)
    ws.cell(r, c).number_format = NF; ws.cell(r, c).font = TOT
for c_, w in zip(range(1, 14), [20, 18] + [15] * 10 + [36]): ws.column_dimensions[L(c_)].width = w
ws.column_dimensions["O"].hidden = True
ws.freeze_panes = "C6"
# ---- sheet B
NM2 = "Unclaimed ITC candidates"
if NM2 in wb.sheetnames: del wb[NM2]
ws2 = wb.create_sheet(NM2, pos + 1)
ws2.row_dimensions[1].height = 21
ws2.cell(2, 1, "VIKRAN ENGINEERING LIMITED").font = TOT
ws2.cell(3, 1, "GSTR-2B credits with NO claim found in the FY 25-26 register (and, for FY 24-25 rows, none in the FY 24-25 register either). FY 25-26 rows are claimable till Nov 2026 (Sec 16(4)) - the October action list. Cross-check FY 26-27 claims (Apr-Jul) before acting.").font = TOT
CC = ["State folder", "FY (derived)", "2B Return Period", "GSTIN of supplier", "Trade/Legal name",
      "Invoice number", "Invoice Date", TXV, IG, CG, SG, "tot2", "ITC Availability",
      "Supply Attract Reverse Charge", "Status"]
HD2 = [c if c != "tot2" else "Total Tax" for c in CC]
for c, h in enumerate(HD2, 1):
    x = ws2.cell(5, c, h); x.font = HF; x.fill = HB
r = 5
cand = cand.sort_values(["FY (derived)", "State folder", "tot2"], ascending=[False, True, False])
for row in cand[CC].itertuples(index=False):
    r += 1
    for c, v in enumerate(row, 1):
        if v is None or v is pd.NaT: v = None
        elif isinstance(v, pd.Timestamp): v = v.to_pydatetime()
        elif not isinstance(v, (str, datetime.datetime, datetime.date)) and pd.isna(v): v = None
        cell = ws2.cell(r, c, v)
        if c in (8, 9, 10, 11, 12): cell.number_format = NF
        if c == 7: cell.number_format = "DD.MM.YYYY"
last2 = r
for c in (8, 9, 10, 11, 12):
    ws2.cell(4, c).value = "=SUBTOTAL(9,%s6:%s%d)" % (L(c), L(c), last2)
    ws2.cell(4, c).number_format = NF; ws2.cell(4, c).font = TOT
ws2.auto_filter.ref = "A5:O%d" % last2
ws2.freeze_panes = "A6"
for c_, w in zip(range(1, 16), [16, 10, 13, 17, 26, 18, 11, 14, 12, 12, 12, 12, 12, 14, 44]): ws2.column_dimensions[L(c_)].width = w
# INDEX rows
ix = wb["INDEX"]
have = {S(ix.cell(r2, 2).value) for r2 in range(5, ix.max_row + 1)}
thin = Side(style="thin"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
for desc, txt, sheet in (("ITC vs 2B - GSTIN level", "GSTN Level 2B", NM), ("Unclaimed ITC - claim by Oct/Nov 2026", "Unclaimed ITC", NM2)):
    if desc in have: continue
    r2 = ix.max_row + 1
    ix.cell(r2, 1, "ITC"); ix.cell(r2, 2, desc)
    x = ix.cell(r2, 7); x.value = '=HYPERLINK("#\'%s\'!A1","%s")' % (sheet, txt)
    x.font = Font(color="0563C1", underline="single")
    for c in range(1, 11): ix.cell(r2, c).border = BD
wb.save(P)
print("built: %s (19 GSTINs) + %s (%d rows)" % (NM, NM2, len(cand)))
