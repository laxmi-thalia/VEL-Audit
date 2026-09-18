import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
SP=os.path.dirname(os.path.abspath(__file__))
BASE="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
JULY=os.path.join(SP,"July2025_converted.xlsx")
M=[("Apr-25","01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),("May-25","02 May 2025","Cleartax sales may 2025.xlsx","working",2),
("Jun-25","03 June 2025","Cleartax Sales Register June 2025.xlsx","Working",1),("Jul-25",None,JULY,"Working",2),
("Aug-25","05 Aug 2025","Cleartax Sales Report August 2025.xlsx","WORKING",2),("Sep-25","06 Sep 2025","Cleartax sales register sep 2025.xlsx","Working",2),
("Oct-25","07 Oct 2025","Cleartax sales Register Oct 2025.xlsx","Working",2),("Nov-25","08 Nov 2025","Cleartax Sales Reg Nov 2025.xlsx","Working",2),
("Dec-25","09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx","Working",2),("Jan-26","10 Jan 2026","Cleartax sales register Jan 2026.xlsx","Working",2),
("Feb-26","11 Feb 2026","Cleartax Sales Register Feb 2026.xlsx","Working",2),("Mar-26","12 Mar 2026","Cleartax sales register Mar 2026.xlsx","Working",1)]
def n(c): return str(c).strip().lower()
# every numeric column we map, source side
COLS={"Item Unit Price":13,"Item Total Amount":14,"Item GST Rate":15,"Item Assessable Amount":17,
      "Item IGST Amount":18,"Item CGST Amount":19,"Item SGST Amount":20,"Item CESS Amount":21,"Item Quantity":35}
src=collections.defaultdict(float)
for lbl,m,f,s,h in M:
    p=f if m is None else os.path.join(BASE,m,f)
    d=pd.read_excel(p,sheet_name=s,header=h); cm={n(c):c for c in d.columns}
    for name in COLS:
        src[name]+=pd.to_numeric(d[cm[n(name)]],errors="coerce").fillna(0).sum()
wb=openpyxl.load_workbook(r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx",read_only=True,data_only=True)
ws=wb["SR_2025-26"]; reg=collections.defaultdict(float); seen=collections.Counter(); rows=0
for r in ws.iter_rows(min_row=5,max_row=27006,min_col=1,max_col=50,values_only=True):
    rows+=1
    for name,idx in COLS.items():
        v=r[idx-1]
        if isinstance(v,(int,float)): reg[name]+=v
    seen[(r[2],r[5],r[7],r[32],r[16],r[34],r[30])]+=1   # gstin,docno,type,hsn,taxable,qty,description
wb.close()
print(f"rows checked: {rows}\n{'column':26s} {'source':>20s} {'register':>20s} {'diff':>12s}")
bad=0
for name,idx in COLS.items():
    d=round(src[name]-reg[name],2)
    if abs(d)>0.05: bad+=1
    print(f"{name:26s} {src[name]:>20,.2f} {reg[name]:>20,.2f} {d:>12,.2f}")
dup=[(k,v) for k,v in seen.items() if v>1]
print(f"\nEXACT duplicate item lines (same GSTIN+DocNo+Type+HSN+Taxable+Qty+Description): {len(dup)}")
for k,v in dup[:5]: print(f"   x{v}  doc={k[1]} type={k[2]} taxable={k[4]}")
print(f"\nVERDICT: {'ALL COLUMN TOTALS TIE' if bad==0 else str(bad)+' COLUMN(S) DIFFER'}")
