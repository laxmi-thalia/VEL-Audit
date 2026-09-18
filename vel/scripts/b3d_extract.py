import os, re, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
B="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-3B/"
MEAS={"Supply Value":"taxable","Integrated Tax":"igst","Central Tax":"cgst","State/UT Tax":"sgst"}
rows=[]
for f in sorted(os.listdir(B)):
    if not f.lower().endswith((".xlsx",".xls")): continue
    st=f.split("LIMITED-")[1]
    st=re.sub(r"-2025-26.*$","",st)
    ov=pd.read_excel(B+f, sheet_name="Overview", header=None)
    gstin=None
    for i in range(len(ov)):
        if str(ov.iloc[i,1]).strip().lower()=="gstin": gstin=str(ov.iloc[i,2]).split("(")[0].strip()
    d=pd.read_excel(B+f, sheet_name="GSTR-3B", header=None)
    months=[str(d.iloc[0,j]).strip() for j in range(3,15)]
    found=0
    rec={"state":st,"gstin":gstin}
    for m in months: 
        for v in MEAS.values(): rec[f"{m}|{v}"]=0.0     # zeros filled by default
    for i in range(1,len(d)):
        sec=str(d.iloc[i,1]).strip(); typ=str(d.iloc[i,2]).strip()
        if not sec.startswith("3.1.D"): continue
        if typ not in MEAS: continue
        found+=1
        for j,m in enumerate(months, start=3):
            v=d.iloc[i,j]
            rec[f"{m}|{MEAS[typ]}"]= float(v) if pd.notna(v) else 0.0
    rec["_rows_found"]=found
    rows.append(rec)
    print(f"  {st:22s} {gstin} | 3.1.D rows present: {found}/4")
T=pd.DataFrame(rows); T.to_pickle(os.path.join(SP,"b3d.pkl"))
print(f"\nfiles read: {len(T)}")
tot={v:sum(T[f'{m}|{v}'].sum() for m in [c.split('|')[0] for c in T.columns if '|taxable' in c]) for v in ['taxable','igst','cgst','sgst']}
print("FY totals from 3B table 3.1(d):")
for k,v in tot.items(): print(f"   {k:8s} {v:>20,.2f}")
