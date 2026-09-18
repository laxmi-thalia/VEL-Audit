import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as L
SP=os.path.dirname(os.path.abspath(__file__))
S=lambda v: "" if v is None or (isinstance(v,float) and pd.isna(v)) else str(v).strip()
HF=Font(bold=True,color="FFFFFF"); HB=PatternFill("solid",fgColor="1F4E79")
P=r"C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx"
wb=openpyxl.load_workbook(P)
ws0=wb["SR_2025-26"]; H={S(ws0.cell(4,c).value):L(c) for c in range(1,ws0.max_column+1)}
SR=lambda n:"'SR_2025-26'!$%s$5:$%s$27006"%(H[n],H[n])
# ---- 1. embed the GST Advance ledger: SHARE-DEPENDENT, wrapped
try:
    import os as _os
    _os.listdir("//192.168.1.69/GST FOLDER")
    SHARE=True
except Exception:
    SHARE=False
    print("share down - GL Adv Data embed deferred; S6 ledger col stays a stated value for now")
if SHARE:
    gl=pd.read_excel("//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/GLs/Outward Tax/GST Advance.xlsx",
                     sheet_name="Data",header=5)
    gl.columns=[S(c) for c in gl.columns]
    gl=gl[gl["Document Number"].notna()].copy()
    gl["ym"]=gl["Year/Month"].map(S)
    gl=gl[gl["ym"].isin({"2025/%02d"%m for m in range(1,13)})]
    if "GL Adv Data" in wb.sheetnames: del wb["GL Adv Data"]
    ga=wb.create_sheet("GL Adv Data")
    ga.append(["Business place","Year/Month","G/L Account","Document Type","Reference","Amount (+ = liability, sign-flipped)"])
    for c in ga[1]: c.font=HF; c.fill=HB
    BP2ST={"AP01":"Andhra Pradesh","AR01":"Arunachal Pradesh","AS01":"Assam","BR01":"Bihar","CG01":"Chhattisgarh","GU01":"Gujarat",
     "JH01":"Jharkhand","JK01":"Jammu & Kashmir","KA01":"Karnataka","MH01":"Maharashtra","MP01":"Madhya Pradesh","PB01":"Punjab",
     "RJ01":"Rajasthan","TG01":"Telangana","TN01":"Tamil Nadu","UP01":"Uttar Pradesh","WB01":"West Bengal"}
    T=lambda v: ("'"+S(v)) if S(v).startswith("=") else S(v)   # text starting '=' would become a formula
    for _,r in gl.iterrows():
        ga.append([T(r["Business place"]),r["ym"],S(int(r["G/L Account"])) if pd.notna(r["G/L Account"]) else "",T(r["Document Type"]),
                   T(r["Reference"]),round(-float(pd.to_numeric(r["Amount in Local Currency"],errors="coerce") or 0),2)])
    NGA=ga.max_row
    for c_,w in zip("ABCDEF",[13,10,12,10,20,22]): ga.column_dimensions[c_].width=w
    g=wb["S6 GL Advance Check"]
    ST2BP={v:k for k,v in BP2ST.items()}
    for r in range(4,12):
        st=S(g.cell(r,2).value)
        if not st or st=="Total": continue
        bp=ST2BP.get(st, S(g.cell(r,1).value))
        g.cell(r,3).value='=SUMIFS(\'GL Adv Data\'!$F$2:$F$%d,\'GL Adv Data\'!$A$2:$A$%d,"%s")'%(NGA,NGA,bp)
    # note: ledger sign = -(books movement); E=C+D stays the zero check

# ---- 2. S9 advances line -> live
s9=wb["S9 Sales Reco"]
for r in range(6,25):
    if S(s9.cell(r,1).value).startswith("Unadjusted advances"):
        OPEN={"Arunachal Pradesh":12398533.90,"Bihar":729161.86,"Gujarat":52303577.12,"Madhya Pradesh":132470378.81}
        for j in range(2,21):
            st=S(s9.cell(3,j).value); g_=S(s9.cell(4,j).value)
            if not g_: continue
            # net movement = closing - opening = received + adjusted(sign-neg) computed LIVE; opening cancels out
            s9.cell(r,j).value=('=SUMIFS(%s,%s,"%s",%s,"Received")+SUMIFS(%s,%s,"%s",%s,"Adjusted")'
                %(SR("Taxable Value"),SR("My GSTIN"),g_,SR("Adv Bucket (helper)"),
                  SR("Taxable Value"),SR("My GSTIN"),g_,SR("Adv Bucket (helper)")))
            s9.cell(r,j).number_format="#,##0"
        break
# ---- 3. S4 Exceptions contribution -> formula where derivable
ex=wb["S4 Exceptions"]
hdr=[S(ex.cell(1,c).value) for c in range(1,ex.max_column+1)]
cS,cF,cG,cH=hdr.index("Status")+1,hdr.index("Register CGST")+1,hdr.index("GL CGST")+1,hdr.index("Contribution to (GL - SR)")+1
for r in range(2,ex.max_row+1):
    st=S(ex.cell(r,cS).value)
    if st in ("IN REGISTER, NOT IN GL","VALUE DIFFERS","IN GL, NOT IN REGISTER","ADVANCE COMPONENT (state level)"):
        ex.cell(r,cH).value="=%s%d-%s%d"%(L(cG),r,L(cF),r)
wb.save(P); print("fix2 applied | GL Adv Data rows:", (NGA-1) if SHARE else "deferred (share down)")
