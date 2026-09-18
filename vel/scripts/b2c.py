import os, re, warnings
warnings.filterwarnings("ignore")
import pandas as pd
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
M = {"01 April 2025":("Cleartax Sales April 2025.xlsx","Working",1),
     "02 May 2025":("Cleartax sales may 2025.xlsx","working",2),
     "03 June 2025":("Cleartax Sales Register June 2025.xlsx","Working",1),
     "05 Aug 2025":("Cleartax Sales Report August 2025.xlsx","WORKING",2),
     "06 Sep 2025":("Cleartax sales register sep 2025.xlsx","Working",2),
     "07 Oct 2025":("Cleartax sales Register Oct 2025.xlsx","Working",2),
     "08 Nov 2025":("Cleartax Sales Reg Nov 2025.xlsx","Working",2),
     "09 Dec 2025":("Cleartax Sales Register Dec 2025.xlsx","Working",2),
     "10 Jan 2026":("Cleartax sales register Jan 2026.xlsx","Working",2),
     "11 Feb 2026":("Cleartax Sales Register Feb 2026.xlsx","Working",2),
     "12 Mar 2026":("Cleartax sales register Mar 2026.xlsx","Working",1)}
def norm(c): return str(c).strip().lower()
first=True
tot_b2c=0; tot_noGSTIN=0; tot=0
for m,(f,s,h) in M.items():
    df = pd.read_excel(os.path.join(BASE,m,f), sheet_name=s, header=h)
    supcols=[c for c in df.columns if "supply" in norm(c)]
    if first:
        print("COLUMNS containing 'supply':", supcols); first=False
    st=[c for c in df.columns if norm(c)=="supply type"][0]
    bg=[c for c in df.columns if norm(c)=="buyer gstin"][0]
    vc = df[st].astype(str).str.strip().str.upper().value_counts().to_dict()
    b2c = sum(v for k,v in vc.items() if "B2C" in k)
    nog = int(df[bg].isna().sum() | 0) if df[bg].isna().any() else 0
    nog = int((df[bg].isna() | (df[bg].astype(str).str.strip().isin(["","nan","URP"]))).sum())
    tot_b2c+=b2c; tot_noGSTIN+=nog; tot+=len(df)
    print(f"{m}: rows={len(df)} SupplyType={vc}  B2C={b2c}  blank/URP BuyerGSTIN={nog}")
print(f"\nTOTAL (11 months, ex-July): rows={tot} B2C-by-SupplyType={tot_b2c} blank/URP-BuyerGSTIN={tot_noGSTIN}")
