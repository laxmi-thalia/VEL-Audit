import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
B="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-1/"
fr=[]
for f in sorted(os.listdir(B)):
    if not f.endswith(".xlsx"): continue
    st=f.split("LIMITED-")[1].rsplit("-Apr",1)[0]
    d=pd.read_excel(B+f, sheet_name="SalesSummary-Net")
    if len(d)==0: continue
    d["__state"]=st; fr.append(d)
S=pd.concat(fr, ignore_index=True)
print("SalesSummary-Net columns:")
for i,c in enumerate(S.columns,1): print(f"  {i:2d}. {c!r}")
print("\nrows:", len(S))
print("Summary Type:", S["Summary Type"].value_counts().to_dict())
S.to_pickle(os.path.join(SP,"g1_summary.pkl"))
