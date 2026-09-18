import os, sys, warnings; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
BASE="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
JULY=os.path.join(SP,"July2025_converted.xlsx")
SAPF=[("Apr-25","01 April 2025","Sales Register April  2025.XLSX"),("May-25","02 May 2025","Sales Register May 2025.XLSX"),
("Jun-25","03 June 2025","Sales Register June 2025.XLSX"),("Jul-25","04 July 2025","Sales Register  July 2025.XLSX"),
("Aug-25","05 Aug 2025","Sales Register Aug 2025.xlsx"),("Sep-25","06 Sep 2025","Sales Register Sept 2025.xlsx"),
("Oct-25","07 Oct 2025","Sales Register Oct 2025.xlsx"),("Nov-25","08 Nov 2025","Sales Register Nov 2025.xlsx"),
("Dec-25","09 Dec 2025","sales register dec 2025.xlsx"),("Jan-26","10 Jan 2026","Sales Register Jan 2026.xlsx"),
("Feb-26","11 Feb 2026","Sales Register Feb 2026.xlsx"),("Mar-26","12 Mar 2026","Sales register March 2026.xlsx")]
CTF={"Apr-25":("01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),"May-25":("02 May 2025","Cleartax sales may 2025.xlsx","working",2),
"Jun-25":("03 June 2025","Cleartax Sales Register June 2025.xlsx","Working",1),"Jul-25":(None,JULY,"Working",2),
"Aug-25":("05 Aug 2025","Cleartax Sales Report August 2025.xlsx","WORKING",2),"Sep-25":("06 Sep 2025","Cleartax sales register sep 2025.xlsx","Working",2),
"Oct-25":("07 Oct 2025","Cleartax sales Register Oct 2025.xlsx","Working",2),"Nov-25":("08 Nov 2025","Cleartax Sales Reg Nov 2025.xlsx","Working",2),
"Dec-25":("09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx","Working",2),"Jan-26":("10 Jan 2026","Cleartax sales register Jan 2026.xlsx","Working",2),
"Feb-26":("11 Feb 2026","Cleartax Sales Register Feb 2026.xlsx","Working",2),"Mar-26":("12 Mar 2026","Cleartax sales register Mar 2026.xlsx","Working",1)}
NEED=["fi invoice document","odn","invoice doc type","bill to customer gstn","base value"]
def n(c): return str(c).strip().lower()
def pick(path):
    xl=pd.ExcelFile(path); best=None
    for s in xl.sheet_names:
        pr=xl.parse(s,header=None,nrows=12)
        for hr in range(min(8,len(pr))):
            cols=[n(v) for v in pr.iloc[hr].tolist() if pd.notna(v)]
            if sum(1 for k in NEED if k in cols)>=4:
                df=xl.parse(s,header=hr)
                if best is None or len(df)>len(best[2]): best=(s,hr,df)
    return best
def col(df,name):
    for c in df.columns:
        if n(c)==name.lower(): return c
S=lambda v: "" if pd.isna(v) else str(v).strip()
a,b=int(sys.argv[1]),int(sys.argv[2])
for lbl,d,f in SAPF[a:b]:
    got=pick(os.path.join(BASE,d,f))
    if not got: print(f"\n===== {lbl}: NO SAP DATA SHEET FOUND"); continue
    sh,hr,sap=got
    cd,cf,cs,ch=CTF[lbl]; ct=pd.read_excel(cf if cd is None else os.path.join(BASE,cd,cf),sheet_name=cs,header=ch)
    odn=col(sap,"ODN"); gstn=col(sap,"Bill to Customer GSTN"); base=col(sap,"Base Value")
    dtyp=col(sap,"Invoice Doc Type"); irn=col(sap,"Irn No")
    cdoc=col(ct,"Document Number"); ctax=col(ct,"Item Assessable Amount")
    sd={S(x) for x in sap[odn] if S(x)}; cdset={S(x) for x in ct[cdoc] if S(x)}
    noG=sap[sap[gstn].map(lambda v:S(v)=="")]
    print(f"\n===== {lbl}   [sheet {sh!r} hdr row {hr+1}]")
    print(f"  SAP rows {len(sap):5d} | docs {len(sd):4d} | base value {pd.to_numeric(sap[base],errors='coerce').sum():>18,.2f}")
    print(f"  CT  rows {len(ct):5d} | docs {len(cdset):4d} | taxable    {pd.to_numeric(ct[ctax],errors='coerce').sum():>18,.2f}")
    print(f"  docs ONLY in SAP {len(sd-cdset):4d} | ONLY in ClearTax {len(cdset-sd):4d}")
    print(f"  no-GSTIN rows {len(noG):5d} | no-IRN rows {len(sap[sap[irn].map(lambda v:S(v)=='')]) if irn else 'n/a'}")
    print(f"  Invoice Doc Type: {sap[dtyp].astype(str).str.strip().value_counts().head(7).to_dict()}")
