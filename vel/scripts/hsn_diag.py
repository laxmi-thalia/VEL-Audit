import win32com.client as win32, pythoncom, os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
P=r"C:\Users\pawar\Downloads\VEL_Step5_HSN_Rate_Summary_DRAFT.xlsx"
pythoncom.CoInitialize(); xl=win32.DispatchEx("Excel.Application"); xl.Visible=False; xl.DisplayAlerts=False
wb=xl.Workbooks.Open(P)
si=wb.Worksheets("SR Items"); hs=wb.Worksheets("HSN Summary")
ns=si.Cells(si.Rows.Count,1).End(-4162).Row
vals=si.Range("A2:L%d"%ns).Value
D=pd.DataFrame(list(vals),columns=["gstin","state","dtc","hsn","gs","uqc","rate","qty","tax","igst","cgst","sgst"])
D=D.fillna({"hsn":"","uqc":"","gs":""})
grp=D.groupby(["gstin","hsn","uqc","rate"],dropna=False)["tax"].sum()
last=hs.Cells(hs.Rows.Count,1).End(-4162).Row
bad=[]
for r in range(4,last):
    key=(hs.Cells(r,2).Value, str(hs.Cells(r,3).Value or ""), str(hs.Cells(r,5).Value or ""), hs.Cells(r,6).Value)
    exp=grp.get((key[0],key[1] if key[1]!="(blank)" else "",key[2],key[3]),None)
    got=hs.Cells(r,8).Value or 0
    if exp is None or abs(got-exp)>1: bad.append((r,key,round(got,2),None if exp is None else round(exp,2)))
print("combo rows where Excel sum != pandas sum:", len(bad))
for b in bad[:10]: print("  ",b)
print("types in rate col:", D["rate"].map(type).value_counts().to_dict())
print("types in hsn col:", D["hsn"].map(type).value_counts().to_dict())
wb.Close(False); xl.Quit()
