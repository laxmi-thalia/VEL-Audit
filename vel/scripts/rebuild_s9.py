"""Rebuild master 'S9 Sales Reco' as an exact replica of last year's 'Sales Reco' (style, layout,
labels, 9C refs, colors, borders, number formats) with THIS year's figures wired live."""
import json, openpyxl
import win32com.client as win32, pythoncom, os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
def S(v): return "" if v is None else str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f=open(P,'r+b'); f.close()
d = json.load(open("salesreco_dump.json")); cells = d["cells"]
def cel(r,c): return cells.get("%d,%d"%(r,c), {})
# full title texts from the xlsb (dump truncated at 120)
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
try:
    wsrc=xl.Workbooks.Open(os.path.abspath("VEL_2425_fresh.xlsb"), ReadOnly=True)
    sr=wsrc.Worksheets("Sales Reco")
    titles={r: str(sr.Cells(r,1).Value or "") for r in (1,2,4)}
    wsrc.Close(SaveChanges=False)
finally:
    xl.Quit()
titles[4] = titles[4].replace("2024-25","2025-26").replace("FY 24-25","FY 25-26")
wb = openpyxl.load_workbook(P)
ws0 = wb["SR_2025-26"]
H = {S(ws0.cell(4,c).value): L(c) for c in range(1, ws0.max_column+1)}
SR = lambda n: "'SR_2025-26'!$%s$5:$%s$27006" % (H[n], H[n])
pos = wb.sheetnames.index("S9 Sales Reco")
del wb["S9 Sales Reco"]
ws = wb.create_sheet("S9 Sales Reco", pos)
BGR2HEX = lambda x: "%02X%02X%02X" % (x & 0xFF, (x>>8)&0xFF, (x>>16)&0xFF)
thin = Side(style="thin")
GSTIN_ROW = {c: S(cel(6,c).get("v")) for c in range(2,21)}
STATE_ROW = {c: S(cel(5,c).get("v")) for c in range(2,21)}
ADV = SR("Adv Bucket (helper)")
def live_formula(r, c):
    g = GSTIN_ROW.get(c); st = STATE_ROW.get(c)
    if not g: return None
    q = lambda dt: '=SUMIFS(%s,%s,"%s",%s,"%s")' % (SR("Taxable Value"), SR("My GSTIN"), g, SR("Document Type Code"), dt)
    if r == 8:  return "=SUMIFS('FS Revenue Data'!$C$2:$C$134,'FS Revenue Data'!$B$2:$B$134,\"%s\")" % st
    if r == 13: return q("INV")
    if r == 14: return q("CRN")
    if r == 15: return '=SUMIFS(%s,%s,"%s",%s,"Received")' % (SR("Taxable Value"), SR("My GSTIN"), g, ADV)
    if r == 16: return q("DBN")
    if r == 18: return '=SUMIFS(%s,%s,"%s",%s,"Adjusted")' % (SR("Taxable Value"), SR("My GSTIN"), g, ADV)
    if r == 34: return "=-IFERROR(INDEX('S6 Advances Control'!$C$4:$C$22,MATCH(%s$6,'S6 Advances Control'!$B$4:$B$22,0)),0)" % L(c)
    if r == 35: return "=IFERROR(INDEX('S6 Advances Control'!$F$4:$F$22,MATCH(%s$6,'S6 Advances Control'!$B$4:$B$22,0)),0)" % L(c)
    return None   # rows 9,17,19,26-29,36,37: CA inputs, blank this year
for key, x in cells.items():
    r, c = map(int, key.split(","))
    if c > 22 or r > 42: continue
    cell = ws.cell(r, c)
    # content
    if x.get("f"):
        fstr = x["f"]
        if "'" not in fstr and "!" not in fstr:   # layout-internal formula: keep verbatim
            cell.value = fstr
        else:
            cell.value = live_formula(r, c)
    elif r in (1,2,4):
        cell.value = titles.get(r, x.get("v"))
    elif c == 1 or r in (5,6) or c == 22:
        cell.value = x.get("v")                   # labels, headers, GSTNs, 9C refs
    else:
        cell.value = live_formula(r, c)           # live or blank input
    # style
    fill = x.get("fill")
    if fill is not None: cell.fill = PatternFill("solid", fgColor=BGR2HEX(fill))
    cell.font = Font(name=x.get("fn","Calibri"), size=x.get("sz",11), bold=x.get("b",False),
                     italic=x.get("i",False), color=BGR2HEX(x.get("fc",0)))
    nf = x.get("nf")
    if nf and nf != "General": cell.number_format = nf if nf.endswith(")") or "#" not in nf else nf + ")"
    ha = x.get("ha")
    align = {}
    if ha == -4108: align["horizontal"]="center"
    if x.get("wrap"): align["wrap_text"]=True
    if align: cell.alignment = Alignment(**align)
    bd = x.get("bd",[-4142]*4)
    if any(b != -4142 for b in bd):
        cell.border = Border(left=thin if bd[0]!=-4142 else None, top=thin if bd[1]!=-4142 else None,
                             bottom=thin if bd[2]!=-4142 else None, right=thin if bd[3]!=-4142 else None)
for c_str, w in d["colw"].items():
    c = int(c_str)
    if c <= 22: ws.column_dimensions[L(c)].width = w
ws.freeze_panes = "B7"
wb.save(P)
print("S9 Sales Reco rebuilt in last year's exact layout | titles:", [titles[r][:40] for r in (1,2,4)])
