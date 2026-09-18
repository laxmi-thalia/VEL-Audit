"""Batch-1 step D: 'T6A1 Extract - 24-25' in last year's 19-col layout + 6A1 marks on the register.
Rows (one per contributing document line):
  (1) register rows matched only in the FY 24-25 2B (working files)      -> "ITC dated 24-25 in 2B of 24-25 availed in 25-26"   Source: ITC Register 2025-26 (matched FY 24-25 2B)
  (2) 2B (Apr25-Aug26) rows, doc FY 24-25, with a 3B claim month          -> "ITC dated 24-25 in 2B of 25-26 availed in 25-26"   Source: GSTR-2B Apr25-Aug26
  (3) 2B rows, doc FY 24-25, no claim month                                -> "ITC dated 24-25 in 2B of 25-26 - Unclaimed"       Source: GSTR-2B Apr25-Aug26
  (4) register rows, invoice FY 24-25, not found in any 2B                 -> "Correction Entries - ITC dated 24-25 reversed/booked in 25-26 (not in 2B)"  Source: ITC Register 2025-26
  (5) RCM register rows: FY 24-25 fiscal postings carried into 3B Apr-25   -> "RCM paid in Mar-25 availed in Apr-25"             Source: RCM Register
GSTR-9 Remarks column is LIVE (reads the current mark on the source sheet by key) so CA remark changes flow; amounts are values
(regenerate with this script when data changes). PivotTable State x GSTR-9 Remarks on top (COM)."""
import os, re, warnings, datetime as dt, collections, pickle
warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
def S(v): return "" if v is None else str(v).strip()
def num(v):
    try: return float(v)
    except Exception: return 0.0
def fy(d):
    if not isinstance(d, dt.datetime): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
SP = os.path.dirname(os.path.abspath(__file__))
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER (2).xlsx"
open(P, "r+b").close()
NF = "#,##0.00"; HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); TOT = Font(bold=True); AMB = PatternFill("solid", fgColor="FFF2CC")
G2ST = {"01AAECR0503Q1ZM": "Jammu & Kashmir", "03AAECR0503Q1ZI": "Punjab", "06AAECR0503Q1ZC": "Haryana", "08AAECR0503Q1Z8": "Rajasthan", "10AAECR0503Q1ZN": "Bihar",
        "12AAECR0503Q1ZJ": "Arunachal Pradesh", "18AAECR0503Q1Z7": "Assam", "19AAECR0503Q1Z5": "West Bengal", "20AAECR0503Q1ZM": "Jharkhand", "23AAECR0503Q1ZG": "Madhya Pradesh",
        "24AAECR0503Q1ZE": "Gujarat", "27AAECR0503Q1Z8": "Maharashtra", "32AAECR0503Q1ZH": "Kerala", "33AAECR0503Q1ZF": "Tamil Nadu", "36AAECR0503Q1Z9": "Telangana",
        "37AAECR0503Q1Z7": "Andhra Pradesh", "09AAECR0503Q1Z6": "Uttar Pradesh", "22AAECR0503Q1ZI": "Chhattisgarh", "29AAECR0503Q1Z4": "Karnataka", "27AAECR0503Q2Z7": "Maharashtra (ISD)"}
wb = openpyxl.load_workbook(P)
reg = wb["ITC Register 2025-26"]; RH = {S(reg.cell(5, c).value): c for c in range(1, reg.max_column + 1)}
RN = max(r for r in range(6, reg.max_row + 1) if reg.cell(r, 4).value)
b2 = wb["GSTR-2B Apr25-Aug26"]; BH = {S(b2.cell(2, c).value): c for c in range(1, b2.max_column + 1)}
NB = max(r for r in range(3, b2.max_row + 1) if b2.cell(r, 1).value)
rcm = wb["RCM Register"]; XH = {S(rcm.cell(5, c).value): c for c in range(1, rcm.max_column + 1)}
XN = max(r for r in range(6, rcm.max_row + 1) if rcm.cell(r, XH["MY GSTN"]).value or rcm.cell(r, XH["Business place"]).value)
meta = pickle.load(open(os.path.join(SP, "itc_b1_meta.pkl"), "rb")); verdict = meta["verdict"]; key2 = meta["key2"]
rv = lambda r, h: reg.cell(r, RH[h]).value
bv = lambda r, h: b2.cell(r, BH[h]).value
rows = []   # dict per extract row + (sheet, row) to build the live remark link
# (1) + (4) from the register
mark_reg = {}
for i in range(RN - 5):
    r = 6 + i; v = verdict[i]
    inv_fy = S(rv(r, "Invoice Year"))
    if v.startswith("Matched in FY 24-25 2B"):
        tag = "ITC dated 24-25 in 2B of 24-25 availed in 25-26"; src = "ITC Register 2025-26 (matched FY 24-25 2B)"
    elif v.startswith("NOT FOUND") and inv_fy == "2024-25":
        tag = "Correction Entries - ITC dated 24-25 booked/reversed in 25-26 (not in 2B)"; src = "ITC Register 2025-26"
    else: continue
    mark_reg[r] = tag
    cm = rv(r, "3B Claim  Month")
    rows.append({"Source": src, "State": rv(r, "State Name"), "2B Return Period": rv(r, "2B Period") or "", "3B Return Period": cm, "DPS - 2B Period": "",
                 "DPS Remarks - 3B month": cm.strftime("%b-%y") if isinstance(cm, dt.datetime) else "", "GSTR-9 Remarks": ("reg", r), "Reclaim": "",
                 "My GSTIN": rv(r, "VEL GSTIN"), "GSTIN of supplier": rv(r, "Vendor GSTIN"), "Trade/Legal name": rv(r, "Vendor Name/RCM Category"),
                 "Invoice Date": rv(r, "Invoice Date"), "Invoice number": rv(r, "Invoice No."), "Invoice type": rv(r, "Document Type"),
                 "Taxable Value": num(rv(r, "Taxable Value")), "IGST": num(rv(r, "IGST")), "CGST": num(rv(r, "CGST")), "SGST": num(rv(r, "SGST")), "ITC Availability": ""})
# (2) + (3) from 2B (doc FY 24-25)
for r in range(3, NB + 1):
    if S(bv(r, "Doc FY (doc date)")) != "2024-25": continue
    g = S(bv(r, "Company GSTIN"))
    rows.append({"Source": "GSTR-2B Apr25-Aug26 (Octa)", "State": G2ST.get(g, g), "2B Return Period": bv(r, "Tax Period"), "3B Return Period": ("b2cm", r), "DPS - 2B Period": "",
                 "DPS Remarks - 3B month": "", "GSTR-9 Remarks": ("b2", r), "Reclaim": "", "My GSTIN": g, "GSTIN of supplier": bv(r, "Supplier GSTIN"),
                 "Trade/Legal name": bv(r, "Supplier Name"), "Invoice Date": bv(r, "Doc Date"), "Invoice number": bv(r, "Doc No"), "Invoice type": bv(r, "Doc Type"),
                 "Taxable Value": num(bv(r, "Taxable Value (Net)")), "IGST": num(bv(r, "IGST (Net)")), "CGST": num(bv(r, "CGST (Net)")), "SGST": num(bv(r, "SGST (Net)")),
                 "ITC Availability": bv(r, "ITC Eligible")})
# (5) RCM paid Mar-25 availed Apr-25: fiscal 2024/12 postings carried into 3B Apr-25
n5 = 0
for r in range(6, XN + 1):
    if S(rcm.cell(r, XH["Year/Month"]).value) == "2024/12":
        g = S(rcm.cell(r, XH["MY GSTN"]).value); n5 += 1
        rows.append({"Source": "RCM Register (fiscal 2024/12 posting, 3B Apr-25)", "State": G2ST.get(g, rcm.cell(r, XH["State"]).value), "2B Return Period": "", "3B Return Period": rcm.cell(r, XH["Final 3B Month"]).value if False else "Apr-25",
                     "DPS - 2B Period": "", "DPS Remarks - 3B month": "Apr-25", "GSTR-9 Remarks": "RCM paid in Mar-25 availed in Apr-25", "Reclaim": "", "My GSTIN": g,
                     "GSTIN of supplier": rcm.cell(r, XH["GSTN"]).value, "Trade/Legal name": rcm.cell(r, XH["Vendor Name"]).value, "Invoice Date": rcm.cell(r, XH["Document Date"]).value,
                     "Invoice number": rcm.cell(r, XH["Reference"]).value, "Invoice type": "RCM", "Taxable Value": num(rcm.cell(r, XH["Taxable Value as per SAP"]).value),
                     "IGST": num(rcm.cell(r, XH["IGST AS PER SAP"]).value), "CGST": num(rcm.cell(r, XH["CGST AS PER SAP"]).value), "SGST": num(rcm.cell(r, XH["SGST AS PER SAP"]).value), "ITC Availability": ""})
print("extract rows:", len(rows), "| by source:", dict(collections.Counter(x["Source"] for x in rows)), "| RCM Mar-25 rows:", n5)
# register 6A1 marks (values in the amber CA columns - proposals)
cA, cB = RH["Considered in Table 6A1"], RH["Remarks for accounting entries- For 6A1"]
for r, tag in mark_reg.items():
    reg.cell(r, cA).value = "Yes"
    reg.cell(r, cB).value = tag
print("register rows marked Considered in Table 6A1:", len(mark_reg))
# ---- sheet
NM = "T6A1 Extract - 24-25"
if NM in wb.sheetnames: del wb[NM]
ws = wb.create_sheet(NM, wb.sheetnames.index("ITCR vs 3B Net ITC"))
ws.row_dimensions[1].height = 21
ws.cell(2, 1, "VIKRAN ENGINEERING LIMITED").font = TOT
ws.cell(3, 1, "Table 6A1 Extract of GSTR 9 (FY 2025-26): FY 24-25-dated ITC touching FY 25-26 - five components per CA (17-09). GSTR-9 Remarks is LIVE from the source sheet's mark; amounts regenerated by script. Pivot on the right; refresh after remark changes.").font = Font(italic=True, color="808080")
HDR = ["Source", "State", "2B Return Period", "3B Return Period", "DPS - 2B Period", "DPS Remarks - 3B month", "GSTR-9 Remarks", "Reclaim for 24-25 dated invoices reflecting in GSTR-2B of 25-26",
       "My GSTIN", "GSTIN of supplier", "Trade/Legal name", "Invoice Date", "Invoice number", "Invoice type", "Taxable Value", "IGST", "CGST", "SGST", "ITC Availability", "Source row (helper)"]
for c, h in enumerate(HDR, 1):
    x = ws.cell(4, c, h); x.font = HF; x.fill = HB; x.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[4].height = 34
r = 4
for x in rows:
    r += 1
    for c, h in enumerate(HDR[:-1], 1):
        v = x.get(h if h in x else "Reclaim" if h.startswith("Reclaim") else h)
        if h == "GSTR-9 Remarks" and isinstance(v, tuple):
            sh, sr = v
            if sh == "reg": v = "='ITC Register 2025-26'!$%s$%d" % (L(cB), sr)
            else: v = "='GSTR-2B Apr25-Aug26'!$%s$%d" % (L(BH["GSTR-9/9C"]), sr)
        if h == "3B Return Period" and isinstance(v, tuple): v = "='GSTR-2B Apr25-Aug26'!$%s$%d" % (L(BH["3B Claim Month"]), v[1])
        cell = ws.cell(r, c, v)
        if h in ("Taxable Value", "IGST", "CGST", "SGST"): cell.number_format = NF
        elif isinstance(v, dt.datetime): cell.number_format = "DD-MM-YYYY"
        elif h == "3B Return Period": cell.number_format = "mmm-yy"
    src = x["GSTR-9 Remarks"]
    ws.cell(r, len(HDR), ("%s!%d" % src) if isinstance(src, tuple) else "RCM Register")
RN2 = r
for c, w in zip(range(1, len(HDR) + 1), [34, 18, 12, 12, 12, 14, 46, 22, 17, 17, 30, 12, 20, 12, 14, 12, 12, 12, 12, 14]): ws.column_dimensions[L(c)].width = w
ws.freeze_panes = "C5"; ws.auto_filter.ref = "A4:%s%d" % (L(len(HDR)), RN2)
ws.cell(4, 8).fill = AMB
# INDEX
ix = wb["INDEX"]; thin = Side(style="thin"); BD = Border(left=thin, right=thin, top=thin, bottom=thin)
have = {S(ix.cell(rr, 2).value) for rr in range(5, ix.max_row + 1)}
desc = "Table 6A1 Extract - FY 24-25 dated ITC in FY 25-26 (five components; GSTR-9 remarks live)"
if desc not in have:
    rr = ix.max_row + 1; ix.cell(rr, 1, "ITC"); ix.cell(rr, 2, desc)
    x = ix.cell(rr, 7); x.value = '=HYPERLINK("#\'%s\'!A1","T6A1 Extract")' % NM; x.font = Font(color="0563C1", underline="single")
    for c in range(1, 11): ix.cell(rr, c).border = BD
wb.save(P)
pickle.dump({"RN2": RN2, "HDR": HDR}, open(os.path.join(SP, "t6a1_meta.pkl"), "wb"))
print("saved: %s rows 5..%d" % (NM, RN2))
