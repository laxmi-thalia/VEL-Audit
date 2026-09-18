import os, warnings, collections
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
SRCTOT={}; SRCDOC=collections.Counter()
for lbl,m,f,s,h in MONTHS:
    p = f if m is None else os.path.join(BASE,m,f)
    df=pd.read_excel(p, sheet_name=s, header=h); cm={n(c):c for c in df.columns}
    SRCTOT[lbl]=dict(rows=len(df),
        tax=round(pd.to_numeric(df[cm["item assessable amount"]],errors="coerce").sum(),2),
        igst=round(pd.to_numeric(df[cm["item igst amount"]],errors="coerce").sum(),2),
        cgst=round(pd.to_numeric(df[cm["item cgst amount"]],errors="coerce").sum(),2),
        sgst=round(pd.to_numeric(df[cm["item sgst amount"]],errors="coerce").sum(),2))
    for k,v in df[cm["document type"]].astype(str).str.strip().value_counts().items(): SRCDOC[k]+=v

wb=openpyxl.load_workbook(r"C:\Users\pawar\Downloads\VEL_Sales_Register_FY2025-26_DRAFT.xlsx", read_only=True, data_only=True)
ws=wb["SR_2025-26"]
REG=collections.defaultdict(lambda: dict(rows=0,tax=0.0,igst=0.0,cgst=0.0,sgst=0.0))
for r in ws.iter_rows(min_row=5,max_row=27006,min_col=1,max_col=50,values_only=True):
    k=r[49]
    if k is None: continue
    d=REG[k]; d["rows"]+=1
    for fld,idx in (("tax",16),("igst",17),("cgst",18),("sgst",19)):
        v=r[idx]; d[fld]+= v if isinstance(v,(int,float)) else 0.0
wb.close()

print(f"{'month':7s} {'rows src/reg':16s} {'taxable diff':>14s} {'igst':>10s} {'cgst':>12s} {'sgst':>12s}")
ok=True
for lbl,*_ in MONTHS:
    s_,r_=SRCTOT[lbl],REG[lbl]
    dr=s_["rows"]-r_["rows"]
    dt=round(s_["tax"]-round(r_["tax"],2),2); di=round(s_["igst"]-round(r_["igst"],2),2)
    dc=round(s_["cgst"]-round(r_["cgst"],2),2); ds=round(s_["sgst"]-round(r_["sgst"],2),2)
    if dr or abs(dt)>0.01 or abs(di)>0.01 or abs(dc)>0.01 or abs(ds)>0.01: ok=False
    print(f"{lbl:7s} {s_['rows']:6d}/{r_['rows']:<6d}   {dt:14.2f} {di:10.2f} {dc:12.2f} {ds:12.2f}")
print("\nGRAND TOTALS")
for fld in ("rows","tax","igst","cgst","sgst"):
    s=sum(SRCTOT[l][fld] for l,*_ in MONTHS); r=sum(REG[l][fld] for l,*_ in MONTHS)
    print(f"  {fld:5s} source={round(s,2):>18} register={round(r,2):>18} diff={round(s-r,2)}")
print("\nVERDICT:", "MERGE IS FAITHFUL - every month ties exactly" if ok else "*** DIFFERENCES FOUND ***")
