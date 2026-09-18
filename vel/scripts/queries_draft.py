import os, re, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
SP=os.path.dirname(os.path.abspath(__file__))
S=lambda v: "" if v is None or (isinstance(v,float) and pd.isna(v)) else str(v).strip()
REG=r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
# compute this-year counts from cached values
wv=openpyxl.load_workbook(REG, read_only=True, data_only=True)["SR_2025-26"]
hdr={}; rows=[]
for i,r in enumerate(wv.iter_rows(min_row=4,max_row=27006,values_only=True),start=4):
    if i==4: hdr={S(v):k for k,v in enumerate(r)}; continue
    rows.append(r)
gA,gH,gF,gAT = hdr["Month"],hdr["Document Type Code"],hdr["Document Number"],hdr["GSTR-1 Month"]
late=collections.Counter(); adv_invno=0
docs_late=set()
invpat=re.compile(r"^[A-Z]{2}\d{8,}$")
for r in rows:
    mA,mT,dt,dn=S(r[gA]),S(r[gAT]),S(r[gH]),S(r[gF])
    if mT and not mT.startswith("No ") and "Summary" not in mT and mA and mT!=mA and not dt.startswith("MOB"):
        docs_late.add(dn); late[(mA,mT)]+=1
    if dt.startswith("MOB") and invpat.match(dn): adv_invno+=1
print("late-reported documents (doc-date month != GSTR-1 month):", len(docs_late), "| top month pairs:", late.most_common(3))
print("advance rows with invoice-like document numbers:", adv_invno)
try:
    f=open(REG,"r+b"); f.close()
except Exception:
    print("register LOCKED"); raise SystemExit
wb=openpyxl.load_workbook(REG)
if "Queries (Draft)" in wb.sheetnames: del wb["Queries (Draft)"]
ws=wb.create_sheet("Queries (Draft)", 2)
HF=Font(bold=True,color="FFFFFF"); HB=PatternFill("solid",fgColor="1F4E79"); W=Alignment(wrap_text=True,vertical="top")
ws["A1"]="DRAFT QUERIES FY 2025-26 — raised on the same lines as last year's Query column (CA query 13). All counts computed from this year's data."
ws["A1"].font=Font(bold=True,size=12)
ws.append([]); ws.append(["#","Last year's query (verbatim, with row count)","This year's position","Proposed query for FY 25-26","Where to verify"])
for c in ws[3]: c.font=HF; c.fill=HB
DATA=[
 ("In GSTR-1, the said transactions are reflecting in Dec-24 but in sales register the same is reflecting in Jan-25 (x1,584)",
  "%d documents reported in a different GSTR-1 month than their document-date month" % len(docs_late),
  "Transactions where the GSTR-1 reporting month differs from the invoice month - please confirm the document dates are correct and whether any belong to a different tax period.",
  "Register: filter column A <> column AT (GSTR-1 Month)"),
 ("Time barred Credit note (x109)",
  "3 credit notes beyond the Sec 34(2) window (MP2500068031/32/33 vs invoice dt 28.11.2023); 280 CNs untestable - no original-invoice reference",
  "Three Mar-26 credit notes cite a FY 23-24 invoice - the GST reduction is beyond the Sec 34(2) time limit; please justify or reverse. Also provide original invoice numbers/dates for the remaining credit notes so the balance can be tested.",
  "Step 7 file 'CN Time-bar Check'"),
 ("Why there are Invoice no. in advances (x21)",
  "%d advance rows carry invoice-series document numbers" % adv_invno,
  "Several mobilization-advance entries carry invoice-series numbers - please confirm these are advances and not unbilled invoices.",
  "Register: filter Document Type MOB*, Document Number pattern XX25......"),
 ("Advance adjusted in Nov-24 but Document date is of Jan-25 (x3)",
  "Advances reconcile month-on-month to Rs 9.49 for the year; one MP Dec-25 rounding row",
  "No adjustment-before-receipt timing found this year (advances reconcile month-wise); please confirm the Dec-25 MP rounding of Rs 9.",
  "Step 6 'Month-on-Month'"),
 ("Mismatch in Taxable value with GSTR-1 & wrong GST Rate (x2)",
  "152 rows where the effective rate differs from the stated rate; 0 taxable-value mismatches among matched documents",
  "152 line items charge a rate different from the stated rate (17%%, 19%%, 27%%, 33%% etc.) - please confirm whether these are data errors or genuine mixed-rate supplies.",
  "Register: filter GST rate check = CHECK; Step 5 Rate-wise Summary"),
 ("Advances of Sep-24 reported in GSTR-1 of Aug-24 (x1)",
  "All 425 summary-level rows found their GSTR-1 line in the matching month",
  "No advance month-mismatch found this year - no query needed.",
  "Register column AT on advance rows"),
 ("(new this year)",
  "34 credit notes in GSTR-1 only (Bihar Jan-26, no IRN, Rs 1.14 crore) + 47 nil duplicates",
  "34 credit notes appear only in GSTR-1 (no IRN, not in books, not taken in 3B) and 47 credit notes are re-reported at nil value - please explain both sets.",
  "Step 2 'Exceptions'"),
 ("(new this year)",
  "9 e-invoiced Telangana documents missing from GSTR-1; tax paid via 3B",
  "Nine e-invoiced documents (TS2500085465-71, 86465-67) were never uploaded to GSTR-1 though tax was paid - please confirm for GSTR-9 disclosure/correction.",
  "Step 2/3 'Exceptions'"),
 ("(new this year)",
  "Bihar Aug-25 3B taxable mis-keyed (5,73,80,517 as 57,38,017), tax correct - repeat of last year",
  "The Aug-25 GSTR-3B taxable value for Bihar appears mis-keyed while tax was paid correctly - same error type as FY 24-25; please confirm for GSTR-9 correction.",
  "Step 3 'Exceptions' section A"),
]
r=3
for k,(ly,ty,q,wh) in enumerate(DATA,1):
    r+=1
    ws.cell(r,1,k); ws.cell(r,2,ly); ws.cell(r,3,ty); ws.cell(r,4,q); ws.cell(r,5,wh)
    for c in range(1,6): ws.cell(r,c).alignment=W
for c_,w in zip("ABCDE",[4,52,52,72,36]): ws.column_dimensions[c_].width=w
ws.freeze_panes="A4"
wb.save(REG); print("Queries (Draft) sheet written: %d query lines" % len(DATA))
