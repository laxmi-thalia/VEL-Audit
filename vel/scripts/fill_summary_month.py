import os, pandas as pd, openpyxl, collections
SP=os.path.dirname(os.path.abspath(__file__))
P=r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
S=lambda v: "" if v is None or (isinstance(v,float) and pd.isna(v)) else str(v).strip()
Su=pd.read_pickle(os.path.join(SP,"g1_summary.pkl"))
Su["m"]=pd.to_datetime(Su["Tax Period"]).dt.strftime("%b-%y"); Su["g"]=Su["Company GSTIN"].map(S)
TYPE={"B2CS Sales":"B2C","Advance Received":"Advance received","Advance Adjusted":"Advance adjusted"}
Su["b"]=Su["Summary Type"].map(TYPE)
have={(r.g, r.b, r.m) for r in Su.itertuples()}
wb=openpyxl.load_workbook(P); ws=wb["SR_2025-26"]
H={S(ws.cell(4,c).value):c for c in range(1,ws.max_column+1)}
cG,cA,cH,cI,cR,cAS,cAT=H["My GSTIN"],H["Month"],H["Document Type Code"],H["Supply Type Code"],H["Taxable Value"],H["Matched with GSTR-1"],H["GSTR-1 Month"]
n=collections.Counter()
for i in range(5,27007):
    if not S(ws.cell(i,cAS).value).startswith("Summary level"): continue
    g,m,h,sup=S(ws.cell(i,cG).value),S(ws.cell(i,cA).value),S(ws.cell(i,cH).value),S(ws.cell(i,cI).value)
    tax=ws.cell(i,cR).value or 0
    if h=="INV" and sup=="B2C": b="B2C"
    elif h=="MOB ADV REC" or (h=="MOB ADV REV" and tax>0): b="Advance received"
    else: b="Advance adjusted"
    if (g,b,m) in have: ws.cell(i,cAT).value=m; n["filled "+b]+=1
    else: ws.cell(i,cAT).value="No %s summary line in GSTR-1 for %s" % (b, m); n["missing "+b]+=1
wb.save(P); print(dict(n))
