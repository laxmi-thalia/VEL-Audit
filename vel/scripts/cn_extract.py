import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
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
S=lambda v: "" if pd.isna(v) else str(v).strip()
out=[]
for lbl,m,f,s,h in M:
    p=f if m is None else os.path.join(BASE,m,f)
    d=pd.read_excel(p,sheet_name=s,header=h); cm={n(c):c for c in d.columns}
    cn=d[d[cm["document type"]].astype(str).str.strip().isin(["CRN","CRN( Manual)"])]
    for _,r in cn.iterrows():
        out.append(dict(month=lbl,state=S(r[cm["state"]]),gstin=S(r[cm["gstin"]]),docno=S(r[cm["document number"]]),
                        cndate=r[cm["document date"]],orig=S(r[cm["original invoice number"]]),
                        pdate=r[cm["preceding invoice date"]],tax=pd.to_numeric(r[cm["item assessable amount"]],errors="coerce") or 0))
D=pd.DataFrame(out)
D.to_pickle(os.path.join(SP,"cn_rows.pkl"))
docs=D.groupby(["gstin","docno"]).agg(state=("state","first"),month=("month","first"),cndate=("cndate","first"),
    orig=("orig","first"),pdate=("pdate","first"),tax=("tax","sum")).reset_index()
print("CN item rows:", len(D), "| distinct credit notes:", len(docs))
has_o=docs["orig"].map(lambda v: v not in ("","nan")).sum(); has_p=docs["pdate"].notna().sum()
print("with Original Invoice Number:", int(has_o), "| with Preceding Invoice Date:", int(has_p))
pd_=pd.to_datetime(docs["pdate"],errors="coerce")
prior=docs[pd_ < "2025-04-01"]
print("credit notes whose ORIGINAL invoice is BEFORE 01.04.2025 (prior FY):", len(prior), "| taxable:", round(prior['tax'].sum(),2))
if len(prior): print(prior[["state","docno","month","pdate","tax"]].head(8).to_string(index=False))
