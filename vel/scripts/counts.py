import os, json, warnings; warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
SP=os.path.dirname(os.path.abspath(__file__))
JULY=os.path.join(SP,"July2025_converted.xlsx")
MONTHS=[("Apr","01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),
 ("May","02 May 2025","Cleartax sales may 2025.xlsx","working",2),
 ("Jun","03 June 2025","Cleartax Sales Register June 2025.xlsx","Working",1),
 ("Jul",None,JULY,"Working",2),
 ("Aug","05 Aug 2025","Cleartax Sales Report August 2025.xlsx","WORKING",2),
 ("Sep","06 Sep 2025","Cleartax sales register sep 2025.xlsx","Working",2),
 ("Oct","07 Oct 2025","Cleartax sales Register Oct 2025.xlsx","Working",2),
 ("Nov","08 Nov 2025","Cleartax Sales Reg Nov 2025.xlsx","Working",2),
 ("Dec","09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx","Working",2),
 ("Jan","10 Jan 2026","Cleartax sales register Jan 2026.xlsx","Working",2),
 ("Feb","11 Feb 2026","Cleartax Sales Register Feb 2026.xlsx","Working",2),
 ("Mar","12 Mar 2026","Cleartax sales register Mar 2026.xlsx","Working",1)]
def n(c): return str(c).strip().lower()
st={}; dt={}; rows=0
for lbl,m,f,s,h in MONTHS:
    p = f if m is None else os.path.join(BASE,m,f)
    df=pd.read_excel(p, sheet_name=s, header=h); rows+=len(df)
    sc=[c for c in df.columns if n(c)=="state"][0]
    dc=[c for c in df.columns if n(c)=="document type"][0]
    for k,v in df[sc].astype(str).str.strip().value_counts().items(): st[k]=st.get(k,0)+v
    for k,v in df[dc].astype(str).str.strip().value_counts().items(): dt[k]=dt.get(k,0)+v
json.dump({"rows":rows,"states":st,"doctypes":dt}, open(os.path.join(SP,"counts.json"),"w"), indent=1)
print("total rows:", rows)
print("states:", len(st), " doctypes:", len(dt))
print(json.dumps(dt, indent=1))
