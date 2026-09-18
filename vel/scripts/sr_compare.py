import os, sys, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
BASE="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
JULY=os.path.join(SP,"July2025_converted.xlsx")
# (label, folder, SAP file, SAP sheet, SAP header idx, CT file, CT sheet, CT header idx)
P=[("Apr-25","01 April 2025","Sales Register April  2025.XLSX","WORKING",1,"Cleartax Sales April 2025.xlsx","Working",1),
   ("May-25","02 May 2025","Sales Register May 2025.XLSX","Working",0,"Cleartax sales may 2025.xlsx","working",2),
   ("Jun-25","03 June 2025","Sales Register June 2025.XLSX","DUMB",0,"Cleartax Sales Register June 2025.xlsx","Working",1),
   ("Jul-25","04 July 2025","Sales Register  July 2025.XLSX","Sheet1",0,None,"Working",2),
   ("Aug-25","05 Aug 2025","Sales Register Aug 2025.xlsx","Working",2,"Cleartax Sales Report August 2025.xlsx","WORKING",2)]
def n(c): return str(c).strip().lower()
def col(df,name):
    for c in df.columns:
        if n(c)==name.lower(): return c
a,b=int(sys.argv[1]),int(sys.argv[2])
for lbl,d,sf,ss,sh,cf,cs,ch in P[a:b]:
    sap=pd.read_excel(os.path.join(BASE,d,sf),sheet_name=ss,header=sh)
    ctp = JULY if cf is None else os.path.join(BASE,d,cf)
    ct=pd.read_excel(ctp,sheet_name=cs,header=ch)
    odn=col(sap,"ODN"); irn=col(sap,"Irn No"); gstn=col(sap,"Bill to Customer GSTN")
    base=col(sap,"Base Value"); btype=col(sap,"Billing Type"); dtyp=col(sap,"Invoice Doc Type")
    cdoc=col(ct,"Document Number"); cirn=col(ct,"Invoice Reference No"); ctax=col(ct,"Item Assessable Amount")
    S=lambda v: "" if pd.isna(v) else str(v).strip()
    sap_docs={S(x) for x in sap[odn] if S(x)}
    ct_docs={S(x) for x in ct[cdoc] if S(x)}
    only_sap=sap_docs-ct_docs; only_ct=ct_docs-sap_docs
    noG=sap[sap[gstn].map(lambda v: S(v)=="")]
    noI=sap[sap[irn].map(lambda v: S(v)=="")] if irn else sap.iloc[0:0]
    print(f"\n===== {lbl}")
    print(f"  SAP rows {len(sap):5d}  base value {pd.to_numeric(sap[base],errors='coerce').sum():>18,.2f}   docs {len(sap_docs)}")
    print(f"  CT  rows {len(ct):5d}  taxable    {pd.to_numeric(ct[ctax],errors='coerce').sum():>18,.2f}   docs {len(ct_docs)}")
    print(f"  documents ONLY in SAP: {len(only_sap)}   |  ONLY in ClearTax: {len(only_ct)}")
    print(f"  SAP rows with NO customer GSTIN (B2C candidates): {len(noG)}  base {pd.to_numeric(noG[base],errors='coerce').sum():,.2f}")
    print(f"  SAP rows with NO IRN: {len(noI)}")
    print(f"  Billing Type: {sap[btype].astype(str).str.strip().value_counts().head(8).to_dict()}")
    print(f"  Invoice Doc Type: {sap[dtyp].astype(str).str.strip().value_counts().to_dict()}")
