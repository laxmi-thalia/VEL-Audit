"""RCM step 0a: transform Conso RCM (322k rows) into last year's 72-col RCM Register layout.
Produces rcm_register.pkl (values grid; formula columns left as None - written as column
formulas at paste time) + prints reconciliation stats vs the conso."""
import pandas as pd, numpy as np, openpyxl, warnings, datetime
warnings.filterwarnings("ignore")
def S(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    return str(v).strip()

print("loading conso...", flush=True)
df = pd.read_excel("conso_rcm.xlsx", sheet_name="RCM OUTPUT WORKING", header=1)
df.columns = [S(c) for c in df.columns]
df = df[df["Document Number"].notna()]
print("conso rows:", len(df), flush=True)

# ---- lookups
GSTINS = {"01AAECR0503Q1ZM":"Jammu & Kashmir","03AAECR0503Q1ZI":"Punjab","06AAECR0503Q1ZC":"Haryana",
"08AAECR0503Q1Z8":"Rajasthan","10AAECR0503Q1ZN":"Bihar","12AAECR0503Q1ZJ":"Arunachal Pradesh",
"18AAECR0503Q1Z7":"Assam","19AAECR0503Q1Z5":"West Bengal","20AAECR0503Q1ZM":"Jharkhand",
"23AAECR0503Q1ZG":"Madhya Pradesh","24AAECR0503Q1ZE":"Gujarat","27AAECR0503Q1Z8":"Maharashtra",
"32AAECR0503Q1ZH":"Kerala","33AAECR0503Q1ZF":"TamilNadu","36AAECR0503Q1Z9":"Telangana",
"37AAECR0503Q1Z7":"Andhra Pradesh","09AAECR0503Q1Z6":"Uttar Pradesh","22AAECR0503Q1ZI":"Chhattisgarh",
"29AAECR0503Q1Z4":"Karnataka"}
BPMAP = {"Jammu & Kashmir":"JK01","Punjab":"PB01","Haryana":"HR01","Rajasthan":"RJ01","Bihar":"BR01",
"Arunachal Pradesh":"AR01","Assam":"AS01","West Bengal":"WB01","Jharkhand":"JH01","Madhya Pradesh":"MP01",
"Gujarat":"GU01","Maharashtra":"MH01","Kerala":"KL01","TamilNadu":"TN01","Telangana":"TG01",
"Andhra Pradesh":"AP01","Uttar Pradesh":"UP01","Chhattisgarh":"CG01","Karnataka":"KA01"}
BP2GSTIN = {bp: g for g, st in GSTINS.items() for s2, bp in [(st, BPMAP[st])]}
# cost-centre master: PC -> (CC narration col F, BP col G)
cc = openpyxl.load_workbook("cost_centres_new.xlsx", read_only=True, data_only=True)["Sheet4"]
PC2 = {}
for i, r in enumerate(cc.iter_rows(values_only=True)):
    if i == 0 or r[0] is None: continue
    try: k = float(r[0])
    except Exception: continue
    if k not in PC2: PC2[k] = (S(r[5]), S(r[6]) if len(r) > 6 else "")
print("PC map:", len(PC2), flush=True)

def fy(d):
    if pd.isna(d): return ""
    y = d.year if d.month >= 4 else d.year - 1
    return "%d-%02d" % (y, (y + 1) % 100)
def pcnum(v):
    try: return float(v)
    except Exception: return None

n = len(df)
out = {}
out["3B Month"] = df["3B Month"].map(S)
out["Final 3B Month"] = pd.to_datetime(df["3B Month"].map(S), format="%d %b %Y", errors="coerce")
out["Business place"] = df["Business place"].map(S)
out["MY GSTN"] = df["Business place"].map(lambda v: BP2GSTIN.get(S(v), ""))
out["State"] = df["State"].map(S)
out["Vikran State code"] = None   # formula =LEFT(D,2)
out["Vendor state code"] = None   # formula =LEFT(V,2)
out["Fiscal Year"] = df["Fiscal Year"]
out["Year/Month"] = df["Year/Month"].map(S)
out["G/L Account"] = df["G/L Account"]
out["GL Name"] = df["GL Name"].map(S)
out["Document type"] = df["Document type"].map(S)
out["Document Number"] = df["Document Number"]
out["Posting Date"] = pd.to_datetime(df["Posting Date"], errors="coerce")
out["Posting Year"] = out["Posting Date"].map(fy)
out["Assignment"] = df["Assignment"].map(S)
out["Reference"] = df["Reference"].map(S)
out["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")
out["Doc year"] = out["Document Date"].map(fy)
out["Vendor Code"] = df["Vendor Code"]
out["Key"] = None                 # formula =D&V
out["GSTN"] = df["GSTN"].map(S)
out["Vendor Name"] = df["Vendor Name"].map(S)
out["Posting Key"] = df["Posting Key"]
out["Amount in Local Currency"] = pd.to_numeric(df["Amount in Local Currency"], errors="coerce")
out["Local Currency"] = df["Local Currency"].map(S)
out["Tax Code"] = df["Tax Code"].map(S)
out["Clearing Document"] = df["Clearing Document"]
out["Profit Center"] = df["Profit Center"]
out["BP as per Profit Centre"] = df["Profit Center"].map(lambda v: PC2.get(pcnum(v), ("", ""))[1])
out["CC as per Profit Centre"] = df["Profit Center"].map(lambda v: PC2.get(pcnum(v), ("", ""))[0])
out["Cost Center"] = df["Cost Center"]
out["Text"] = df["Text"].map(S)
out["Offsetting Account"] = df["Offsetting Account"]
out["Expenditure GL"] = ""        # PENDING: Trial Balance not yet received
out["GL Correct"] = ""            # PENDING: judgment vs Nature of Services (post-ITC)
out["Inv. No."] = df["Inv. No."].map(S)
out["Inv. Date"] = pd.to_datetime(df["Inv. Date"], errors="coerce")
out["GST Rate"] = pd.to_numeric(df["GST Rate"], errors="coerce")
out["Taxable Value As per filed return"] = pd.to_numeric(df["Taxable Value As per filed return"], errors="coerce")
out["IGST AS PER GST PORTAL"] = pd.to_numeric(df["IGST AS PER GST PORTAL"], errors="coerce")
out["CGST AS PER GST PORTAL"] = pd.to_numeric(df["CGST AS PER GST PORTAL"], errors="coerce")
out["SGST AS PER GST PORTAL"] = pd.to_numeric(df["SGST AS PER GST PORTAL"], errors="coerce")
out["Taxable Value as per SAP"] = pd.to_numeric(df["Taxable Value as per SAP"], errors="coerce")
out["IGST AS PER SAP"] = pd.to_numeric(df["IGST AS PER SAP"], errors="coerce")
out["CGST AS PER SAP"] = pd.to_numeric(df["CGST AS PER SAP"], errors="coerce")
out["SGST AS PER SAP"] = pd.to_numeric(df["SGST AS PER SAP"], errors="coerce")
out["Total GST"] = None           # formula =SUM(SAP tax cols)
out["Inv. Value"] = pd.to_numeric(df["Inv. Value"], errors="coerce")
out["Nature of Services"] = df["Service"].map(S)
out["Payment Status"] = df["Payment Status"].map(S)
out["Rate Check"] = None          # formula =TotalGST/SAPtaxable*100
out["Query"] = ""
out["Query Description"] = ""
out["Vendor"] = None              # formula =LEFT(GSTN,2)
out["My GSTN"] = None             # formula =LEFT(D,2)
out["As per State"] = None        # formula =IF(Vendor=My,"Intra State","Inter State")
out["As per Amounts"] = None      # formula =IF(IGSTSAP=0,"Intra State","Inter State")
out["POS Check"] = None           # formula =(As per Amounts = As per State)
out["Query2"] = ""
out["Query Description2"] = ""
out["ITC Eligibility"] = ""       # PENDING: post-ITC phase
out["GSTR 3B Claim month"] = ""   # PENDING: from ITC register
out["Found in GL"] = ""           # step 4
out["GL Remarks"] = ""            # step 4
res = pd.DataFrame(out)
res.to_pickle("rcm_register.pkl")

print("\n==== stats ====")
print("rows:", len(res), "| columns:", len(res.columns))
print("MY GSTN unmapped BPs:", sorted(set(res.loc[res['MY GSTN']=='', 'Business place'].unique()))[:10],
      "(%d rows)" % (res['MY GSTN']=='').sum())
print("PC->CC unmapped rows:", (res['CC as per Profit Centre']=='').sum())
print("SAP taxable sum:      %18.2f" % res["Taxable Value as per SAP"].sum())
print("SAP IGST/CGST/SGST:   %15.2f %15.2f %15.2f" % (res["IGST AS PER SAP"].sum(), res["CGST AS PER SAP"].sum(), res["SGST AS PER SAP"].sum()))
print("portal taxable sum:   %18.2f" % res["Taxable Value As per filed return"].sum())
print("portal I/C/S:         %15.2f %15.2f %15.2f" % (res["IGST AS PER GST PORTAL"].sum(), res["CGST AS PER GST PORTAL"].sum(), res["SGST AS PER GST PORTAL"].sum()))
print("Amount in LC sum:     %18.2f" % res["Amount in Local Currency"].sum())
print("3B months:", sorted(res["3B Month"].unique())[:14])
print("nature of services filled:", (res["Nature of Services"]!="").sum(), "| blank:", (res["Nature of Services"]=="").sum())
print("distinct GL accounts:", res["G/L Account"].nunique(), "| doc types:", dict(res["Document type"].value_counts().head(6)))
