import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
JULY = r"C:\Users\pawar\AppData\Local\Temp\claude\c--PROJECTS-accountic\ed6b1fc9-75d0-4eb0-9100-e931702d9fa7\scratchpad\July2025_converted.xlsx"
MONTHS = [("Apr","01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),
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
hdrs={}
for lbl,m,f,s,h in MONTHS:
    p = f if m is None else os.path.join(BASE,m,f)
    hdrs[lbl] = {str(c).strip().lower() for c in pd.read_excel(p, sheet_name=s, header=h, nrows=0).columns}
print("headers loaded for:", list(hdrs))
import json
json.dump({k:sorted(v) for k,v in hdrs.items()}, open(os.path.join(os.path.dirname(JULY),"hdrs.json"),"w"))
print("saved hdrs.json")
