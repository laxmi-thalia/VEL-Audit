import os, warnings, random
warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
SP=os.path.dirname(os.path.abspath(__file__))
BASE="//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
JULY=os.path.join(SP,"July2025_converted.xlsx")
MONTHS=[("Apr-25","01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),
 ("May-25","02 May 2025","Cleartax sales may 2025.xlsx","working",2),
 ("Jun-25","03 June 2025","Cleartax Sales Register June 2025.xlsx","Working",1),
 ("Jul-25",None,JULY,"Working",2),
 ("Aug-25","05 Aug 2025","Cleartax Sales Report August 2025.xlsx","WORKING",2),
 ("Sep-25","06 Sep 2025","Cleartax sales register sep 2025.xlsx","Working",2),
 ("Oct-25","07 Oct 2025","Cleartax sales Register Oct 2025.xlsx","Working",2),
 ("Nov-25","08 Nov 2025","Cleartax Sales Reg Nov 2025.xlsx","Working",2),
 ("Dec-25","09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx","Working",2),
 ("Jan-26","10 Jan 2026","Cleartax sales register Jan 2026.xlsx","Working",2),
 ("Feb-26","11 Feb 2026","Cleartax Sales Register Feb 2026.xlsx","Working",2),
 ("Mar-26","12 Mar 2026","Cleartax sales register Mar 2026.xlsx","Working",1)]
def n(c): return str(c).strip().lower()
def num(v):
    x=pd.to_numeric(v, errors="coerce")
    return None if pd.isna(x) else round(float(x),2)
def txt(v):
    if v is None: return ""
    s=str(v).strip()
    return "" if s in ("nan","NaT","None") else s
wb=openpyxl.load_workbook(r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx", read_only=True, data_only=True)
ws=wb["SR_2025-26"]; reg={}
for i,r in enumerate(ws.iter_rows(min_row=5,max_row=27006,min_col=1,max_col=50,values_only=True),start=5):
    reg.setdefault(r[49],[]).append((i,r))
wb.close()
random.seed(11); checked=0; bad=[]
for lbl,m,f,s,h in MONTHS:
    p=f if m is None else os.path.join(BASE,m,f)
    df=pd.read_excel(p, sheet_name=s, header=h); cm={n(c):c for c in df.columns}
    rows=reg[lbl]
    for k in random.sample(range(len(df)), min(25,len(df))):
        i,rr=rows[k]; src=df.iloc[k]
        T=[("DocNo","t",txt(src[cm["document number"]]),txt(rr[5])),
           ("GSTIN","t",txt(src[cm["gstin"]]),txt(rr[2])),
           ("BuyerGSTIN","t",txt(src[cm["buyer gstin"]]),txt(rr[9])),
           ("IRN","t",txt(src[cm["invoice reference no"]]),txt(rr[23])),
           ("Taxable","n",num(src[cm["item assessable amount"]]),num(rr[16])),
           ("IGST","n",num(src[cm["item igst amount"]]),num(rr[17])),
           ("CGST","n",num(src[cm["item cgst amount"]]),num(rr[18])),
           ("SGST","n",num(src[cm["item sgst amount"]]),num(rr[19])),
           ("HSN","n",num(src[cm["item hsn code"]]),num(rr[32])),
           ("Qty","n",num(src[cm["item quantity"]]),num(rr[34])),
           ("Rate","n",num(src[cm["item gst rate"]]),num(rr[14]))]
        for fld,kind,a,b in T:
            checked+=1
            same = (a==b) if kind=="n" else (a==b)
            if not same: bad.append((lbl,i,fld,repr(a),repr(b)))
print(f"rows sampled: {sum(min(25,len(reg[l])) for l,*_ in MONTHS)}   field comparisons: {checked}")
print(f"mismatches: {len(bad)}")
for b in bad[:12]: print("   ", b)
print("VERDICT:", "ROWS ALIGNED - no shuffling, values faithful" if not bad else "*** MISALIGNMENT ***")
