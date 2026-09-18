import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
from openpyxl.utils import get_column_letter
BASE = "//server/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
T=[("June","03 June 2025","Cleartax Sales Register June 2025.xlsx","Working",1),
   ("March","12 Mar 2026","Cleartax sales register Mar 2026.xlsx","Working",1)]
def n(c): return str(c).strip().lower()
for lbl,m,f,s,h in T:
    df=pd.read_excel(os.path.join(BASE,m,f), sheet_name=s, header=h)
    stc=[c for c in df.columns if n(c)=="state"][0]
    ci=list(df.columns).index(stc)+1
    mask=df[stc].astype(str).str.strip().str.lower().str.replace(" ","")=="westbengal"
    sub=df[mask]
    spellings=sorted(set(sub[stc].astype(str)))
    excel_rows=[i+h+2 for i in sub.index]
    print(f"\n{lbl}: sheet={s!r}, header on Excel row {h+1}, data starts Excel row {h+2}")
    print(f"   State column = {get_column_letter(ci)} (header {stc!r})")
    print(f"   exact spelling(s) used: {spellings}")
    print(f"   {len(sub)} rows at Excel rows: {excel_rows}")
