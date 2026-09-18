import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
SP=os.path.dirname(os.path.abspath(__file__))
S=lambda v: "" if v is None or (isinstance(v,float) and pd.isna(v)) else str(v).strip()
# statuses from the Step 4 workbook's SR Docs sheet (already computed)
src=openpyxl.load_workbook(r"C:\Users\pawar\Downloads\VEL_Step4_Reco_GL_vs_SR_DRAFT.xlsx", read_only=True)
sdws=src["SR Docs"]; smap={}
for r in sdws.iter_rows(min_row=2, values_only=True):
    smap[(S(r[1]).upper(), S(r[2]).upper(), S(r[3]))]=S(r[10])
src.close()
P=r"C:\Users\pawar\Downloads\VEL_Step1_Sales_Register_FY2025-26_DRAFT.xlsx"
wb=openpyxl.load_workbook(P); ws=wb["SR_2025-26"]
H={S(ws.cell(4,c).value):c for c in range(1,ws.max_column+1)}
cG,cF,cH,cAU=H["My GSTIN"],H["Document Number"],H["Document Type Code"],H["Matched with GL"]
n=collections.Counter()
for i in range(5,27007):
    k=(S(ws.cell(i,cG).value).upper(), S(ws.cell(i,cF).value).upper(), S(ws.cell(i,cH).value))
    v=smap.get(k,"")
    ws.cell(i,cAU).value=v or None; n[v or "(blank)"]+=1
wb.save(P); print(dict(n))
