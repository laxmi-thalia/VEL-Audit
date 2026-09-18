"""Make the 'Step 1 Flags' detail listing LIVE: each listing row pulls the k-th flagged register row
via AGGREGATE(15,6,...) (no dynamic arrays needed), so fixing a register row removes it from the list."""
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
P = r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
R0, R1, NLIST = 5, 27006, 400
def S(v): return "" if v is None else str(v).strip()
wb = openpyxl.load_workbook(P); ws = wb["SR_2025-26"]; g = wb["Step 1 Flags"]
H = {S(ws.cell(4, c).value): L(c) for c in range(1, ws.max_column + 1)}
F = H["Data Flags"]
SR = lambda col: "'SR_2025-26'!$%s$%d:$%s$%d" % (col, R0, col, R1)
# find the old static DETAIL block and clear everything from its title row down
start = None
for r in range(1, g.max_row + 1):
    if isinstance(g.cell(r, 1).value, str) and g.cell(r, 1).value.startswith("DETAIL"):
        start = r; break
start = start or g.max_row + 2
for r in range(start, g.max_row + 1):
    for c in range(1, 12): g.cell(r, c).value = None; g.cell(r, c).fill = PatternFill(fill_type=None)
g.cell(start, 1, "DETAIL — LIVE list of flagged register rows (blank rows at the bottom are spare capacity, up to %d)" % NLIST).font = Font(bold=True)
hdr = ["#", "Register row", "Data Flags", "State", "My GSTIN", "Document No", "Doc Type", "Taxable", "Document Date", "IRN"]
for j, t in enumerate(hdr, 1):
    x = g.cell(start + 1, j, t); x.font = Font(bold=True, color="FFFFFF"); x.fill = PatternFill("solid", fgColor="1F4E79")
cols = {"Data Flags": F, "State": H["My State"], "My GSTIN": H["My GSTIN"], "Document No": H["Document Number"],
        "Doc Type": H["Document Type Code"], "Taxable": H["Taxable Value"], "Document Date": H["Document Date"], "IRN": H["Invoice Reference No"]}
# helper column in the register: running count of flagged rows (plain formulas, no arrays)
hc = max(openpyxl.utils.column_index_from_string(v) for v in H.values()) + 1
if S(ws.cell(4, hc).value) != "Flag seq (helper)":
    ws.cell(4, hc, "Flag seq (helper)").font = Font(bold=True, color="FFFFFF")
    ws.cell(4, hc).fill = PatternFill("solid", fgColor="1F4E79")
HC = L(hc)
for r in range(R0, R1 + 1):
    ws.cell(r, hc, '=N(%s%d)+IF(%s%d<>"",1,0)' % (HC, r - 1, F, r) if r > R0 else '=IF(%s%d<>"",1,0)' % (F, r))
ws.column_dimensions[HC].width = 10
first = start + 2
for k in range(1, NLIST + 1):
    r = first + k - 1
    g.cell(r, 1, k)
    # k-th register row whose Data Flags is non-empty
    g.cell(r, 2, '=IFERROR(MATCH(A%d,%s,0)+%d,"")' % (r, SR(HC), R0 - 1))
    for j, (name, col) in enumerate(cols.items(), start=3):
        g.cell(r, j, '=IF($B%d="","",INDEX(\'SR_2025-26\'!$%s:$%s,$B%d))' % (r, col, col, r))
    g.cell(r, 8).number_format = "#,##0.00"; g.cell(r, 9).number_format = "dd-mmm-yy"
last = first + NLIST - 1
BLOCK = ["Document Date missing", "Recipient GSTIN missing on B2B document", "IRN missing on B2B document", "Advance / B2B tag contradiction", "Taxable Value missing"]
g.conditional_formatting.add("C%d:C%d" % (first, last), FormulaRule(formula=['OR(%s)' % ",".join('ISNUMBER(SEARCH("%s",$C%d))' % (b, first) for b in BLOCK)], fill=PatternFill("solid", fgColor="FFC7CE"), stopIfTrue=True))
g.conditional_formatting.add("C%d:C%d" % (first, last), FormulaRule(formula=['LEN($C%d)>0' % first], fill=PatternFill("solid", fgColor="FFEB9C")))
for c_, w in zip("ABCDEFGHIJ", [6, 12, 60, 20, 18, 18, 12, 16, 14, 22]): g.column_dimensions[c_].width = max(g.column_dimensions[c_].width or 0, w)
wb.save(P); print("live detail list written: rows %d-%d (%d slots)" % (first, last, NLIST))
