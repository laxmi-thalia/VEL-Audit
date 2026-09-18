import sys, warnings; warnings.filterwarnings("ignore")
import pandas as pd
B = "//192.168.1.69/GST FOLDER/GST Returns/GST Audit & Annual Return/FY 2025-26/1. Corporate Clients/VEL/Clients Data/GLs/Outward Tax/"
f = sys.argv[1]
xl = pd.ExcelFile(B + f)
print("FILE:", f, "| sheets:", xl.sheet_names)
def looks_header(vals):
    nn = [v for v in vals if pd.notna(v)]
    return len(nn) >= 5 and sum(1 for v in nn if isinstance(v, str)) / len(nn) > 0.7
for s in xl.sheet_names:
    pr = xl.parse(s, header=None, nrows=12)
    if pr.empty:
        print("  %r: empty" % s); continue
    hr = next((i for i in range(len(pr)) if looks_header(pr.iloc[i].tolist())), None)
    if hr is None:
        print("  %r: no header row in first 12; first cells:" % s, [str(v)[:30] for v in pr.iloc[0].tolist() if pd.notna(v)][:6]); continue
    df = xl.parse(s, header=hr)
    print("  %r: header row %d | %d rows x %d cols" % (s, hr + 1, len(df), df.shape[1]))
    for i, c in enumerate(df.columns, 1):
        col = df[c]; nn = col.notna() & (col.astype(str).str.strip() != "")
        samp = col[nn].astype(str).head(2).tolist()
        print("     %2d. %-34s filled %6d  e.g. %s" % (i, str(c).strip()[:34], int(nn.sum()), [x[:28] for x in samp]))
