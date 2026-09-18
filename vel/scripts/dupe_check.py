import os, warnings, collections; warnings.filterwarnings("ignore")
import pandas as pd
SP=os.path.dirname(os.path.abspath(__file__))
BASE="//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/05.08.2026 Main Data"
JULY=os.path.join(SP,"July2025_converted.xlsx")
M=[("Apr-25","01 April 2025","Cleartax Sales April 2025.xlsx","Working",1),("May-25","02 May 2025","Cleartax sales may 2025.xlsx","working",2),
("Jun-25","03 June 2025","Cleartax Sales Register June 2025.xlsx","Working",1),("Jul-25",None,JULY,"Working",2),
("Aug-25","05 Aug 2025","Cleartax Sales Report August 2025.xlsx","WORKING",2),("Sep-25","06 Sep 2025","Cleartax sales register sep 2025.xlsx","Working",2),
("Oct-25","07 Oct 2025","Cleartax sales Register Oct 2025.xlsx","Working",2),("Nov-25","08 Nov 2025","Cleartax Sales Reg Nov 2025.xlsx","Working",2),
("Dec-25","09 Dec 2025","Cleartax Sales Register Dec 2025.xlsx","Working",2),("Jan-26","10 Jan 2026","Cleartax sales register Jan 2026.xlsx","Working",2),
("Feb-26","11 Feb 2026","Cleartax Sales Register Feb 2026.xlsx","Working",2),("Mar-26","12 Mar 2026","Cleartax sales register Mar 2026.xlsx","Working",1)]
def n(c): return str(c).strip().lower()
src=collections.Counter(); docmonth=collections.defaultdict(set)
for lbl,m,f,s,h in M:
    p=f if m is None else os.path.join(BASE,m,f)
    d=pd.read_excel(p,sheet_name=s,header=h); cm={n(c):c for c in d.columns}
    for _,r in d.iterrows():
        key=(str(r[cm['gstin']]).strip(), str(r[cm['document number']]).strip(), str(r[cm['document type']]).strip(),
             str(r[cm['item hsn code']]).strip(), round(float(pd.to_numeric(r[cm['item assessable amount']],errors='coerce') or 0),2))
        src[key]+=1
        docmonth[(str(r[cm['gstin']]).strip(), str(r[cm['document number']]).strip())].add(lbl)
dup=[(k,v) for k,v in src.items() if v>1]
print(f"exact duplicate item lines IN THE SOURCE FILES: {len(dup)}")
for k,v in dup[:8]: print(f"   x{v}  doc={k[1]} type={k[2]} hsn={k[3]} taxable={k[4]}")
multi=[(k,v) for k,v in docmonth.items() if len(v)>1]
print(f"\ndocuments appearing in MORE THAN ONE monthly file: {len(multi)}")
for k,v in multi[:8]: print(f"   {k[1]} -> {sorted(v)}")
