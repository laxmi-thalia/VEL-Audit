import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
P=r"C:\Users\pawar\Downloads\VEL_Step3_Reco_GSTR1_vs_3B_DRAFT.xlsx"
wb=openpyxl.load_workbook(P)
if "Exceptions" in wb.sheetnames: del wb["Exceptions"]
sh=wb.create_sheet("Exceptions",1)
H=Font(bold=True,color="FFFFFF"); F=PatternFill("solid",fgColor="C00000")
SUB=Font(bold=True,size=12); AMB=PatternFill("solid",fgColor="FFF2CC"); BAD=PatternFill("solid",fgColor="FFC7CE")
W=Alignment(wrap_text=True,vertical="top")
sh.append(["GSTR-1 vs GSTR-3B — why each difference exists (FY 2025-26)"])
sh["A1"].font=Font(bold=True,size=13)
sh.append(["Only 2 states differ materially. Every rupee of the annual difference (Rs 3,47,65,815.57) is explained below."])
sh.append([])
sh.append(["State","Month","GSTR-1 taxable","GSTR-3B taxable","Difference (G1 - 3B)","CGST diff","Cause","Evidence","Effect on tax paid"])
for c in sh[4]: c.font=H; c.fill=F; c.alignment=Alignment(horizontal="center",wrap_text=True)
rows=[
 ("Bihar","Aug 2025",57380517.22,5738017.00,51642500.22,0.39,
  "3B taxable value MIS-KEYED: 5,73,80,517 was entered as 57,38,017 (digits dropped). The tax columns were keyed correctly.",
  "CGST/SGST agree to Rs 0.39 while taxable differs by Rs 5.16 crore — impossible unless only the taxable cell is wrong. IDENTICAL failure mode to FY 24-25, where the office noted 'Reporting Error in GSTR-3B - Taxes reported and discharged correctly' on a Rs 18.09 crore Bihar difference.",
  "NONE - tax paid correctly; disclosure error only"),
 ("Bihar","Jan 2026",77287336.14,88685438.06,-11398101.92,-1025829.18,
  "81 credit notes (no IRN) were entered in GSTR-1 but NOT taken in 3B - and they do not exist in the books either (Step 2 finding).",
  "Difference equals the Step 2 books-vs-GSTR-1 Bihar gap to the paisa (11,398,101.92); CGST diff 10,25,829.18 = exactly 9% of it. Books and 3B agree with each other; GSTR-1 is the odd one out.",
  "NONE - the CN reduction was never taken in 3B, so no tax was short-paid. GSTR-1 detail is wrong / needs explanation"),
 ("Telangana","May 2025",0.00,1478854.18,-1478854.18,-133096.91,
  "E-invoiced documents never uploaded to GSTR-1, but included in 3B and tax paid.",
  "The 9 books-only documents from Step 2 (TS2500085465-71, 86465-67, all with valid 64-char IRNs). Each month's CGST diff = exactly 9% of the taxable diff.",
  "NONE - tax fully paid via 3B; GSTR-1 reporting omission"),
 ("Telangana","Jun 2025",0.00,4950777.98,-4950777.98,-445570.01,
  "Same as above.","Same batch - contiguous document numbers.","NONE - tax fully paid via 3B"),
 ("Telangana","Aug 2025",0.00,-951049.20,951049.20,85594.40,
  "Same as above: Aug invoice 448,302.52 net of credit notes -1,399,351.72 = -951,049.20.",
  "Nets exactly to the Step 2 Aug Telangana books-only documents.","NONE - tax correctly adjusted via 3B"),
 ("Punjab","May 2025",37941866.67,37941866.67,0.00,1.00,
  "Re 1 rounding on CGST in the advances month.","Taxable identical to the paisa.","Ignore"),
]
r0=5
for r in rows:
    sh.append(list(r))
    i=sh.max_row
    for c in sh[i]: c.alignment=W
    fill = AMB if r[0]!="Punjab" else None
    if fill:
        for c in sh[i]: c.fill=fill
sh.append([])
sh.append(["ANNUAL TIE-OUT:"]); sh[sh.max_row][0].font=SUB
sh.append(["Bihar Aug 2025 keying error",51642500.22])
sh.append(["Bihar Jan 2026 GSTR-1-only credit notes",-11398101.92])
sh.append(["Telangana (3 months, GSTR-1 omission)",-5478582.96])
sh.append(["Rounding (all states)",0.23])
sh.append(["TOTAL = annual GSTR-1 minus GSTR-3B difference",34765815.57])
sh[sh.max_row][0].font=Font(bold=True); sh[sh.max_row][1].font=Font(bold=True)
sh.append([])
sh.append(["BOTTOM LINE: no tax was short-paid anywhere. Both differences are RETURN-PREPARATION errors:"])
sh.append(["  - Bihar Aug: taxable mis-keyed in 3B (tax right)  - repeat of last year's error; fix via GSTR-9"])
sh.append(["  - Bihar Jan: 81 no-IRN credit notes in GSTR-1 only - CA to establish what they are; not in books, not in 3B"])
sh.append(["  - Telangana: 9 e-invoiced documents missing from GSTR-1 - tax paid; disclose/correct via GSTR-9"])
for col,w in zip("ABCDEFGHI",[12,10,18,18,20,14,44,54,34]): sh.column_dimensions[col].width=w
sh.freeze_panes="A5"
wb.save(P); print("Exceptions sheet written")
