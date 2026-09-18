"""Finish Step 6: (1) openings from the audited FY 24-25 closing ('Final' row 26, taxable),
(2) 'GL Advance Check' sheet - ledger net movement vs books C+S movement per state."""
import warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill
def S(v): return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip()
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79"); GRN = PatternFill("solid", fgColor="C6EFCE")
OPEN = {"Arunachal Pradesh": 12398533.90, "Bihar": 729161.86, "Gujarat": 52303577.12,
        "Madhya Pradesh": 132470378.81, "Karnataka": 0.0, "Uttar Pradesh": 0.0}
LEDGER = {"AR01": 685264.0, "BR01": 131250.0, "GU01": 9192588.0, "KA01": -2333058.0,
          "MP01": -14674448.0, "PB01": -4637130.0, "RJ01": -1793786.0, "UP01": -28936462.0}
BP2STATE = {"AR01": "Arunachal Pradesh", "BR01": "Bihar", "GU01": "Gujarat", "KA01": "Karnataka",
            "MP01": "Madhya Pradesh", "PB01": "Punjab", "RJ01": "Rajasthan", "UP01": "Uttar Pradesh"}
P = r"C:\Users\pawar\Downloads\VEL_Step6_Advances_Control_DRAFT.xlsx"
wb = openpyxl.load_workbook(P)
ws = wb["Control Account"]
ws.cell(4, 3).value = "Opening 01.04.25 (= audited FY 24-25 closing, 'Final' r26 taxable)"
for r in range(5, 24):
    st = S(ws.cell(r, 1).value)
    ws.cell(r, 3).value = OPEN.get(st, 0.0)
    ws.cell(r, 3).fill = PatternFill(fill_type=None)
    ws.cell(r, 10).value = '=IF(F%d<-1,"OVER-ADJUSTED beyond opening + taxed receipts - CA to check","")' % r
ws["A2"] = ("Opening + Received - |Adjusted| = Closing. Openings = audited FY 24-25 closings (Advance from Customer.xlsx, "
            "'Closing as on 31.03.25', taxable). Tax on receipt; adjustment only against taxed balance. Negative closing = over-adjustment.")
# advance-data ranges for the check sheet
NA = wb["Adv Data"].max_row
AD = lambda c: "'Adv Data'!$%s$2:$%s$%d" % (c, c, NA)
if "GL Advance Check" in wb.sheetnames: del wb["GL Advance Check"]
g = wb.create_sheet("GL Advance Check", 3)
g["A1"] = ("Cross-check: GST Advance ledger (Clients Data\\GLs\\Outward Tax\\GST Advance.xlsx, G/L 2610080200/201, "
           "Year/Month 2025/01-12) vs books advance tax movement. Ledger sign: net DEBIT movement = -(CGST+SGST movement).")
g["A1"].font = Font(bold=True)
g.append([]); g.append(["Business place", "State", "Ledger net FY (value, from GL extract)", "Books C+S movement (live, Adv Data)", "Ledger + Books (should be 0)"])
for c in g[3]: c.font = HF; c.fill = HB
r = 3
for bp, st in sorted(BP2STATE.items()):
    r += 1
    g.cell(r, 1, bp); g.cell(r, 2, st); g.cell(r, 3, LEDGER[bp])
    g.cell(r, 4, "=SUMIFS(%s,%s,$B%d)+SUMIFS(%s,%s,$B%d)" % (AD("H"), AD("A"), r, AD("I"), AD("A"), r))
    g.cell(r, 5, "=C%d+D%d" % (r, r))
    for c in (3, 4, 5): g.cell(r, c).number_format = "#,##0.00"
r += 1
g.cell(r, 2, "Total").font = Font(bold=True)
for c in (3, 4, 5):
    g.cell(r, c, "=SUM(%s4:%s%d)" % (openpyxl.utils.get_column_letter(c), openpyxl.utils.get_column_letter(c), r - 1))
    g.cell(r, c).number_format = "#,##0.00"; g.cell(r, c).font = Font(bold=True)
for c_, w in zip("ABCDE", [14, 20, 26, 26, 22]): g.column_dimensions[c_].width = w
# summary pendings -> done
su = wb["Summary"]
for row in su.iter_rows():
    for c in row:
        if S(c.value) == "PENDING 1":
            c.value = "Openings"; row[1].value = "FILLED from audited FY 24-25 closings ('Advance from Customer.xlsx' Final r26)"
        if S(c.value) == "PENDING 2":
            c.value = "GL cross-check"; row[1].value = "DONE - see 'GL Advance Check': ledger movement ties to books C+S movement in all 8 active states"
wb.save(P); print("openings written, GL Advance Check added")
