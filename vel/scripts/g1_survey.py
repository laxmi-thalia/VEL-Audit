import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd
B="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/DPS Workings/Portal Reports/GSTR-1/"
tot=0; dts=collections.Counter(); sts=collections.Counter(); irn_ok=irn_no=0; amd=0
tax=0.0; frames=[]
print(f"{'state':22s} {'rows':>6s} {'IRN blank':>10s} {'amend':>6s}")
for f in sorted(os.listdir(B)):
    if not f.endswith(".xlsx"): continue
    st=f.split("LIMITED-")[1].rsplit("-Apr",1)[0]
    df=pd.read_excel(B+f, sheet_name="Sales-Net", header=0)
    if len(df)==0: print(f"{st:22s} {0:6d}"); continue
    df["__state"]=st; frames.append(df)
    tot+=len(df)
    for k,v in df["Doc Type"].value_counts().items(): dts[k]+=v
    for k,v in df["Sale Type"].value_counts().items(): sts[k]+=v
    b=df["IRN"].isna() | (df["IRN"].astype(str).str.strip().isin(["","nan"]))
    irn_no+=int(b.sum()); irn_ok+=int((~b).sum())
    a=int((df["Is Amendment"].astype(str).str.strip().str.lower()=="yes").sum()); amd+=a
    tax+=pd.to_numeric(df["Taxable Value (Net)"],errors="coerce").sum()
    print(f"{st:22s} {len(df):6d} {int(b.sum()):10d} {a:6d}")
A=pd.concat(frames, ignore_index=True)
A.to_pickle(os.path.join(os.path.dirname(os.path.abspath(__file__)),"gstr1.pkl"))
print(f"\nTOTAL GSTR-1 document rows: {tot}")
print(f"Doc Type: {dict(dts)}")
print(f"Sale Type: {dict(sts)}")
print(f"IRN populated: {irn_ok} | blank: {irn_no}")
print(f"Amendment rows: {amd}")
print(f"Taxable Value (Net) total: {round(tax,2)}")
