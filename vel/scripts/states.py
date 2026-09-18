import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
JULY = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\July2025_converted.xlsx"
M = [("01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),
     ("02 May 2025","Cleartax sales may 2025.xlsx","working",2),
     ("03 June 2025","Cleartax Sales Register June 2025.xlsx","Working",1),
     (None,JULY,"Working",2),
     ("05 Aug 2025","Cleartax Sales Report August 2025.xlsx","WORKING",2),
     ("06 Sep 2025","Cleartax sales register sep 2025.xlsx","Working",2),
     ("07 Oct 2025","Cleartax sales Register Oct 2025.xlsx","Working",2),
     ("08 Nov 2025","Cleartax Sales Reg Nov 2025.xlsx","Working",2),
     ("09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx","Working",2),
     ("10 Jan 2026","Cleartax sales register Jan 2026.xlsx","Working",2),
     ("11 Feb 2026","Cleartax Sales Register Feb 2026.xlsx","Working",2),
     ("12 Mar 2026","Cleartax sales register Mar 2026.xlsx","Working",1)]
def norm(c): return str(c).strip().lower()
frames=[]
for m,f,s,h in M:
    p = f if m is None else os.path.join(BASE,m,f)
    df = pd.read_excel(p, sheet_name=s, header=h)
    stc=[c for c in df.columns if norm(c)=="state"][0]
    gs =[c for c in df.columns if norm(c)=="gstin"][0]
    frames.append(df[[stc,gs]].rename(columns={stc:"State",gs:"GSTIN"}))
all_=pd.concat(frames)
all_["State"]=all_["State"].astype(str).str.strip()
g=all_.groupby("State").size().sort_values(ascending=False)
print("ClearTax book sales — rows per state (12 months):")
for k,v in g.items(): print(f"   {k:25s} {v:6d}")
print(f"\ndistinct own GSTINs in data: {all_['GSTIN'].astype(str).str.strip().nunique()}")
