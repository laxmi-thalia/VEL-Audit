"""Step 1 register — Kolshet Road call changes:
 1. remove the Step 2 sheets (Step 2 lives in its own file); SAP Enrichment to the end
 2. Data Flags -> LIVE per-row formula; highlights -> conditional formatting (clear when fixed)
 3. Sample Selection -> 'Source' column naming the SAP file / fields behind each PO ref
 5. gate COUNTIF ranges widened to a buffer so manually appended rows are counted
"""
import os, collections, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
P = r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
R0, R1, BUF = 5, 27006, 60000
def S(v): return "" if v is None else str(v).strip()
wb = openpyxl.load_workbook(P)

# ---- 1. sheets
for n in ["Step 2 Summary", "SR vs GSTR-1", "Month-on-Month 1", "Step 2 Exceptions", "GSTR-1 Data"]:
    if n in wb.sheetnames: del wb[n]
if "SAP Enrichment" in wb.sheetnames:
    wb.move_sheet("SAP Enrichment", offset=len(wb.sheetnames) - 1 - wb.sheetnames.index("SAP Enrichment"))
ws = wb["SR_2025-26"]
H = {S(ws.cell(4, c).value): L(c) for c in range(1, ws.max_column + 1)}
E, Hh, I, J, K, P_, R, Y, AG, AH, AJ, AK = (H["Document Date"], H["Document Type Code"], H["Supply Type Code"], H["Recipient GSTIN"],
    H["Recipient Legal Name"], H["GST rate check"], H["Taxable Value"], H["Invoice Reference No"], H["GOOD (G) or SERVICE(S)"],
    H["HSN or SAC Code"], H["Quantity for uploading"], H["Unit of Measurement for uploading"])
F = H["Data Flags"]

# ---- 2. live flag formula + conditional formatting
def flag_formula(r):
    adv = 'LEFT(%s{r},3)="MOB"' % Hh
    b2c = 'AND(%s{r}="INV",%s{r}="B2C")' % (Hh, I)
    parts = [
        'IF(AND(NOT(%s),%s{r}=""),"Document Date missing; ","")' % (adv, E),
        'IF(AND(NOT(%s),NOT(%s),%s{r}=""),"Recipient GSTIN missing on B2B document; ","")' % (adv, b2c, J),
        'IF(AND(NOT(%s),NOT(%s),%s{r}=""),"IRN missing on B2B document; ","")' % (adv, b2c, Y),
        'IF(OR(AND(%s,OR(%s{r}="B2B",%s{r}="B2C")),AND(OR(%s{r}="INV",%s{r}="CRN",%s{r}="DBN"),LEFT(%s{r},3)="MOB")),"Advance / B2B tag contradiction; ","")' % (adv, I, I, Hh, Hh, Hh, I),
        'IF(%s{r}="","Taxable Value missing; ","")' % R,
        'IF(AND(%s{r}<>"",LEN(%s{r})<>64),"IRN length <> 64; ","")' % (Y, Y),
        'IF(%s{r}="CHECK","GST rate charged <> stated rate; ","")' % P_,
        'IF(AND(NOT(%s),%s{r}=""),"HSN missing; ","")' % (adv, AH),
        'IF(AND(%s{r}="G",OR(N(%s{r})=0,%s{r}="")),"Quantity / UoM missing on goods row; ","")' % (AG, AJ, AK),
        'IF(AND(%s{r}<>"",%s{r}=""),"Recipient name missing; ","")' % (J, K),
    ]
    return ("=" + "&".join(parts)).replace("{r}", str(r))
CHECK_COLS = {"Document Date missing": E, "Recipient GSTIN missing on B2B document": J, "IRN missing on B2B document": Y,
              "Advance / B2B tag contradiction": I, "Taxable Value missing": R, "IRN length <> 64": Y,
              "GST rate charged <> stated rate": P_, "HSN missing": AH, "Quantity / UoM missing on goods row": AJ,
              "Recipient name missing": K}
BLOCK = ["Document Date missing", "Recipient GSTIN missing on B2B document", "IRN missing on B2B document",
         "Advance / B2B tag contradiction", "Taxable Value missing"]
none = PatternFill(fill_type=None)
for r in range(R0, R1 + 1):
    ws["%s%d" % (F, r)] = flag_formula(r)
    ws["%s%d" % (F, r)].fill = none
    for col in set(CHECK_COLS.values()):
        ws["%s%d" % (col, r)].fill = none
RED = PatternFill("solid", fgColor="FFC7CE"); AMB = PatternFill("solid", fgColor="FFEB9C")
rng = lambda col: "%s%d:%s%d" % (col, R0, col, R1)
# flag column: red if any blocking text, amber if anything at all
ws.conditional_formatting.add(rng(F), FormulaRule(formula=['OR(%s)' % ",".join('ISNUMBER(SEARCH("%s",$%s%d))' % (b, F, R0) for b in BLOCK)], fill=RED, stopIfTrue=True))
ws.conditional_formatting.add(rng(F), FormulaRule(formula=['LEN($%s%d)>0' % (F, R0)], fill=AMB))
# each checked column: highlight when its own flag text is present on the row
for name, col in CHECK_COLS.items():
    ws.conditional_formatting.add(rng(col), FormulaRule(formula=['ISNUMBER(SEARCH("%s",$%s%d))' % (name, F, R0)], fill=RED if name in BLOCK else AMB))

# ---- 5. gate ranges -> buffer
g = wb["Step 1 Flags"]
for row in g.iter_rows(min_row=1, max_row=30):
    for c in row:
        if isinstance(c.value, str) and "$%s$%d:$%s$%d" % (F, R0, F, R1) in c.value:
            c.value = c.value.replace("$%s$%d:$%s$%d" % (F, R0, F, R1), "$%s$%d:$%s$%d" % (F, R0, F, BUF))
g["A2"] = ("Counts are LIVE: the register's 'Data Flags' column is a formula on every row, so fixing a cell clears its flag "
           "(and deleting the flag text does nothing - it recalculates). Range covers rows %d-%d; copy the Data Flags formula down for appended rows." % (R0, BUF))

# ---- 3. Sample Selection source column
ss = wb["Sample Selection"]
src_month = {}
cDoc, cMon = H["Document Number"], H["Source File Month"]
for r in range(R0, R1 + 1):
    d = S(ws["%s%d" % (cDoc, r)].value)
    if d and d not in src_month: src_month[d] = S(ws["%s%d" % (cMon, r)].value)
hdr = {S(ss.cell(1, c).value): c for c in range(1, ss.max_column + 1)}
cSI, cPO = hdr["Sample Invoice"], hdr["Sample PO"]
cSrc = ss.max_column + 1
ss.cell(1, cSrc, "Source of Sample PO").font = Font(bold=True, color="FFFFFF"); ss.cell(1, cSrc).fill = PatternFill("solid", fgColor="1F4E79")
MN = {"Apr-25": "April 2025", "May-25": "May 2025", "Jun-25": "June 2025", "Jul-25": "July 2025", "Aug-25": "Aug 2025", "Sep-25": "Sept 2025",
      "Oct-25": "Oct 2025", "Nov-25": "Nov 2025", "Dec-25": "Dec 2025", "Jan-26": "Jan 2026", "Feb-26": "Feb 2026", "Mar-26": "March 2026"}
for r in range(2, ss.max_row + 1):
    d = S(ss.cell(r, cSI).value); po = S(ss.cell(r, cPO).value)
    if not d: continue
    m = src_month.get(d, "")
    ss.cell(r, cSrc, ("Clients Data\\05.08.2026 Main Data\\<%s>\\Sales Register %s.xlsx (SAP) - fields 'Contract No' (= PO) and 'Sales Doc Number', "
                      "matched on ODN = this document number" % (m, MN.get(m, m))) if po else "no Contract No / Sales Doc Number in the SAP register for this document")
ss.column_dimensions[L(cSrc)].width = 95
wb.save(P)
print("saved | sheets:", wb.sheetnames)
print("flag column", F, "now a live formula on rows %d-%d; CF rules: %d; gate range to row %d" % (R0, R1, len(ws.conditional_formatting), BUF))
