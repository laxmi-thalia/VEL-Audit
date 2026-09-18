"""3B vice-versa: on '3B Data', per row (GSTIN x month) add live books figure, diff, and status -
so anything reported in 3B with nothing in books (and vice versa) is flagged on the 3B side itself."""
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as L
def S(v): return "" if v is None else str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
try:
    f=open(P,'r+b'); f.close()
except Exception as e:
    print("LOCKED - close the master first:", e); raise SystemExit(1)
HF=Font(bold=True,color="FFFFFF"); HB=PatternFill("solid",fgColor="1F4E79")
wb = openpyxl.load_workbook(P)
ws0 = wb["SR_2025-26"]
H = {S(ws0.cell(4,c).value): L(c) for c in range(1, ws0.max_column+1)}
SR = lambda n: "'SR_2025-26'!$%s$5:$%s$27006" % (H[n], H[n])
b3 = wb["3B Data"]
hdr = [S(b3.cell(1,c).value) for c in range(1, b3.max_column+1)]
# idempotent: reuse columns if already there
def col_for(name):
    return hdr.index(name)+1 if name in hdr else None
c0 = b3.max_column
cH = col_for("SR Taxable (books, live)") or c0+1
cI = col_for("Diff (3B - books)") or c0+2
cJ = col_for("Vice-versa status") or c0+3
for c,name in ((cH,"SR Taxable (books, live)"),(cI,"Diff (3B - books)"),(cJ,"Vice-versa status")):
    x=b3.cell(1,c,name); x.font=HF; x.fill=HB
for r in range(2, b3.max_row+1):
    b3.cell(r,cH).value = '=SUMIFS(%s,%s,$B%d,%s,$C%d)' % (SR("Taxable Value"),SR("My GSTIN"),r,SR("Month"),r)
    b3.cell(r,cI).value = "=D%d-%s%d" % (r, L(cH), r)
    b3.cell(r,cJ).value = ('=IF(ABS(%s%d)<1,"Matched with books",'
        'IF(%s%d=0,"In books, NOT reported in 3B",'
        'IF(D%d=0,"Reported in 3B, NOT in books",'
        '"3B differs from books - see S3 SR vs 3B MoM")))' % (L(cI),r,L(cH),r,r))
    b3.cell(r,cH).number_format = "#,##0.00"; b3.cell(r,cI).number_format = "#,##0.00"
for c,w in ((cH,18),(cI,16),(cJ,34)): b3.column_dimensions[L(c)].width = w
b3.auto_filter.ref = "A1:%s%d" % (L(cJ), b3.max_row)
wb.save(P)
print("3B Data: added cols %s/%s/%s on %d rows" % (L(cH),L(cI),L(cJ),b3.max_row-1))
