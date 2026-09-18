"""Bring master 'S2 Reco (CA format)' to the step file's full 35-col layout:
AF = books docs not in GSTR-1 (live, register), AG = GSTR-1 docs not in books (live, GSTR-1 Data),
AH = AA - AF - AG (unexplained residual), remarks move AE -> AI."""
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as L
def S(v): return "" if v is None else str(v).strip()
P = r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
f=open(P,'r+b'); f.close()
wb = openpyxl.load_workbook(P)
ws0 = wb["SR_2025-26"]
H = {S(ws0.cell(4,c).value): L(c) for c in range(1, ws0.max_column+1)}
SR = lambda n: "'SR_2025-26'!$%s$5:$%s$27006" % (H[n], H[n])
ng = wb["GSTR-1 Data"].max_row
G1 = lambda c: "'GSTR-1 Data'!$%s$2:$%s$%d" % (c, c, ng)
ms = wb["S2 Reco (CA format)"]
HF=Font(bold=True,color="FFFFFF"); HB=PatternFill("solid",fgColor="1F4E79")
# move remarks AE -> AI (rows 54-73), then clear AE
for r in range(54, 75):
    v = ms.cell(r, 31).value      # AE
    ms.cell(r, 35).value = v      # AI
    ms.cell(r, 31).value = None
ms.cell(54,35).font=HF; ms.cell(54,35).fill=HB
# new headers
for c, h in ((32,"Auto: books docs not in GSTR-1 (taxable)"),(33,"Auto: GSTR-1 docs not in books (taxable)"),(34,"UNEXPLAINED residual (taxable)")):
    x=ms.cell(54,c,h); x.font=HF; x.fill=HB
for r in range(55, 74):
    ms.cell(r,32).value = '=SUMIFS(%s,%s,$A%d,%s,"Not in GSTR-1")' % (SR("Taxable Value"), SR("My GSTIN"), r, SR("Matched with GSTR-1"))
    ms.cell(r,33).value = '=-SUMIFS(%s,%s,$A%d,%s,"Not in books")' % (G1("F"), G1("A"), r, G1("N"))
    ms.cell(r,34).value = "=AA%d-AF%d-AG%d" % (r,r,r)
    for c in (32,33,34): ms.cell(r,c).number_format="#,##0.00"
for c in (32,33,34):
    ms.cell(74,c).value = "=SUM(%s55:%s73)" % (L(c),L(c))
    ms.cell(74,c).number_format="#,##0.00"; ms.cell(74,c).font=Font(bold=True)
for c,w in ((31,4),(32,22),(33,22),(34,20),(35,60)): ms.column_dimensions[L(c)].width=w
wb.save(P)
print("S2 Reco (CA format): AF/AG/AH added, remarks moved to AI | GSTR-1 Data rows:", ng)
