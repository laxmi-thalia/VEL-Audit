import warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl, collections
S=lambda v: "" if pd.isna(v) or v is None else str(v).strip()
m=pd.ExcelFile(r"C:\Users\pawar\Downloads\HSN_SAC.xlsx")
hsn=set(m.parse("HSN_MSTR")["HSN_CD"].map(S)); sac=set(m.parse("SAC_MSTR")["SAC_CD"].map(S))
print("master: HSN codes", len(hsn), "| SAC codes", len(sac))
wb=openpyxl.load_workbook(r"C:\Users\pawar\Downloads\VEL_Step5_HSN_Rate_Summary_DRAFT.xlsx", read_only=True)
ws=wb["SR Items"]
used=collections.Counter()
for r in ws.iter_rows(min_row=2, values_only=True):
    c=S(r[3])
    if c: used[c]+=1
wb.close()
print("distinct codes used by VEL:", len(used))
bad=[]
for c,n in used.items():
    master = sac if c.startswith("99") else hsn
    ok = c in master or (len(c)==8 and c[:6] in master) or (len(c)==6 and any(x.startswith(c) for x in master))
    if not ok: bad.append((c,n))
print("codes NOT found in the master (self or 6-digit parent):", len(bad))
for c,n in sorted(bad,key=lambda x:-x[1])[:15]: print("   %-10s rows %5d" % (c,n))
