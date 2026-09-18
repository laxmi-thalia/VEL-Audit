"""Long-format month-on-month for the standalone Step 2 workbook:
State | Month | GSTR-1 Type | SR Value | GSTR-1 Value | Difference  (taxable value, live SUMIFS).
One row per State x Month x Type (19 x 12 x 6). Replaces the wide 'Month-on-Month' sheet."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
P = r"C:\Users\pawar\Downloads\VEL_Step2_Reco_GSTR1_DRAFT.xlsx"
GSTINS = [("01AAECR0503Q1ZM", "Jammu & Kashmir"), ("03AAECR0503Q1ZI", "Punjab"), ("06AAECR0503Q1ZC", "Haryana"),
 ("08AAECR0503Q1Z8", "Rajasthan"), ("10AAECR0503Q1ZN", "Bihar"), ("12AAECR0503Q1ZJ", "Arunachal Pradesh"),
 ("18AAECR0503Q1Z7", "Assam"), ("19AAECR0503Q1Z5", "West Bengal"), ("20AAECR0503Q1ZM", "Jharkhand"),
 ("23AAECR0503Q1ZG", "Madhya Pradesh"), ("24AAECR0503Q1ZE", "Gujarat"), ("27AAECR0503Q1Z8", "Maharashtra"),
 ("32AAECR0503Q1ZH", "Kerala"), ("33AAECR0503Q1ZF", "TamilNadu"), ("36AAECR0503Q1Z9", "Telangana"),
 ("37AAECR0503Q1Z7", "Andhra Pradesh"), ("09AAECR0503Q1Z6", "Uttar Pradesh"), ("22AAECR0503Q1ZI", "Chhattisgarh"),
 ("29AAECR0503Q1Z4", "Karnataka")]
MONTHS = ["Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26"]
TYPES = ["B2B", "B2C", "Credit note", "Debit note", "Advance received", "Advance adjusted"]
HF = Font(bold=True, color="FFFFFF"); HB = PatternFill("solid", fgColor="1F4E79")
wb = openpyxl.load_workbook(P)
NSR = wb["SR Data"].max_row; NG = wb["GSTR-1 Data"].max_row
for n in ("Month-on-Month", "Pivot - Month on Month"):
    if n in wb.sheetnames: del wb[n]
ws = wb.create_sheet("Month-on-Month", 2)
ws["A1"] = "Sales Register vs GSTR-1 — state / month / type (taxable value, live SUMIFS). Difference = SR minus GSTR-1."
ws["A1"].font = Font(bold=True)
ws.append([]); ws.append(["State", "GSTIN", "Month", "GSTR-1 Type",
           "Books Taxable", "Books IGST", "Books CGST", "Books SGST",
           "GSTR-1 Taxable", "GSTR-1 IGST", "GSTR-1 CGST", "GSTR-1 SGST",
           "Diff Taxable", "Diff IGST", "Diff CGST", "Diff SGST"])
for c in ws[3]: c.font = HF; c.fill = HB; c.alignment = Alignment(horizontal="center", wrap_text=True)
SRm = {0: "F", 1: "G", 2: "H", 3: "I"}; G1m = {0: "F", 1: "G", 2: "H", 3: "I"}
r = 3
for g, st in GSTINS:
    for mo in MONTHS:
        for t in TYPES:
            r += 1
            ws.cell(r, 1, st); ws.cell(r, 2, g); ws.cell(r, 3, mo); ws.cell(r, 4, t)
            for mi in range(4):
                ws.cell(r, 5 + mi, "=SUMIFS('SR Data'!$%s$2:$%s$%d,'SR Data'!$A$2:$A$%d,$B%d,'SR Data'!$B$2:$B$%d,$C%d,'SR Data'!$J$2:$J$%d,$D%d)" % (SRm[mi], SRm[mi], NSR, NSR, r, NSR, r, NSR, r))
                ws.cell(r, 9 + mi, "=SUMIFS('GSTR-1 Data'!$%s$2:$%s$%d,'GSTR-1 Data'!$A$2:$A$%d,$B%d,'GSTR-1 Data'!$K$2:$K$%d,$C%d,'GSTR-1 Data'!$J$2:$J$%d,$D%d)" % (G1m[mi], G1m[mi], NG, NG, r, NG, r, NG, r))
                ws.cell(r, 13 + mi, "=%s%d-%s%d" % (openpyxl.utils.get_column_letter(5 + mi), r, openpyxl.utils.get_column_letter(9 + mi), r))
            for c in range(5, 17): ws.cell(r, c).number_format = "#,##0.00"
last = r
ws.conditional_formatting.add("M4:M%d" % last, CellIsRule(operator="notBetween", formula=["-1", "1"], fill=PatternFill("solid", fgColor="FFC7CE")))
ws.auto_filter.ref = "A3:P%d" % last
for c_, w in zip("ABCDEFGHIJKLMNOP", [20, 18, 9, 18] + [15] * 12): ws.column_dimensions[c_].width = w
ws.freeze_panes = "E4"
wb.save(P)
print("long-format sheet written: rows", last - 3, "x 16 cols")
