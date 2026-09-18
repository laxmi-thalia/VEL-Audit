import os, sys, warnings; warnings.filterwarnings("ignore")
import pandas as pd
BASE="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
FILES=[("Apr-25","01 April 2025","Sales Register April  2025.XLSX"),("May-25","02 May 2025","Sales Register May 2025.XLSX"),
("Jun-25","03 June 2025","Sales Register June 2025.XLSX"),("Jul-25","04 July 2025","Sales Register  July 2025.XLSX"),
("Aug-25","05 Aug 2025","Sales Register Aug 2025.xlsx"),("Sep-25","06 Sep 2025","Sales Register Sept 2025.xlsx"),
("Oct-25","07 Oct 2025","Sales Register Oct 2025.xlsx"),("Nov-25","08 Nov 2025","Sales Register Nov 2025.xlsx"),
("Dec-25","09 Dec 2025","sales register dec 2025.xlsx"),("Jan-26","10 Jan 2026","Sales Register Jan 2026.xlsx"),
("Feb-26","11 Feb 2026","Sales Register Feb 2026.xlsx"),("Mar-26","12 Mar 2026","Sales register March 2026.xlsx")]
def looks_header(v):
    nn=[x for x in v if pd.notna(x)]
    return len(nn)>=5 and sum(1 for x in nn if isinstance(x,str))/len(nn)>0.7
a,b=int(sys.argv[1]),int(sys.argv[2])
for lbl,d,f in FILES[a:b]:
    p=os.path.join(BASE,d,f)
    print(f"\n================ {lbl}  {f}")
    try: xl=pd.ExcelFile(p)
    except Exception as e: print("   LOAD FAILED",e); continue
    print("   sheets:", xl.sheet_names)
    for s in xl.sheet_names:
        pr=xl.parse(s,header=None,nrows=15)
        if pr.empty: print(f"   {s!r}: EMPTY"); continue
        hr=next((i for i in range(len(pr)) if looks_header(pr.iloc[i].tolist())),None)
        if hr is None:
            print(f"   {s!r}: shape={pr.shape} no header row in first 15")
            continue
        df=xl.parse(s,header=hr)
        cols=[str(c).strip() for c in df.columns]
        print(f"   {s!r}: hdr_row={hr+1} rows={len(df)} cols={len(cols)}")
        print(f"      first 12 cols: {cols[:12]}")
