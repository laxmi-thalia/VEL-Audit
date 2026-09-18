# VEL master — per-sheet format spec (generated from the verified FY 25-26 master)

Row 1 of every sheet is the reserved button row. Machine-exact copy: formats.json.

## INDEX
- dims 46x10 | header row 4 | freeze None
- columns: A=Arena | B=Descriptions | C=State Code | D=GSTN | E=GSTR-9 | F=GSTR-9C | G=Particulars | H=GSTR 9 | I=GSTR 9C | J=Remarks

## Open Points
- dims 15x5 | header row 5 | freeze A6
- columns: A=# | B=Owner | C=Point | D=Detail / where to look | E=Resolved? (write here)

## SR_2025-26
- dims 27007x61 | header row 5 | freeze None
- columns: A=Month | B=My State | C=My GSTIN | D=Supplier Legal Name | E=Document Date | F=Document Number | G=Invoice Type | H=Document Type Code | I=Supply Type Code | J=Recipient GSTIN | K=Recipient Legal Name | L=Place of Supply | M=Item Price | N=Gross Amount | O=GST Rate | P=Original Invoice Number | Q=Original Invoice Date | R=GST rate check | S=Actual Rate Charged (%) | T=Taxable Value | U=IGST Amount | V=CGST Amount | W=SGST Amount | X=Cess Amount | Y=Total GST | Z=Total Invoice Value | AA=Invoice Reference No | AB=IRN (len) | AC=Status | AD=Cancel Date | AE=Document Status | AF=Ship To GSTIN | AG=Ship To LegalName | AH=Item Description | AI=GOOD (G) or SERVICE(S) | AJ=HSN or SAC Code | AK=Item Description for uploading | AL=Quantity for uploading | AM=Unit of Measurement for uploading | AN=GL Name | AO=Profit Centre | AP=CC Name | AQ=Cost Centre | AR=VEL 2 digit | AS=Customer 2 digit | AT=POS check | AU=GSTR-1 Type | AV=Matched with GSTR-1 | AW=GSTR-1 Month | AX=Matched with GL | AY=Matched with FS - revenue | AZ=GSTR-9 | BA=Queries | BB=Sample Invoices | BC=Sample POs | BD=Data Flags | BE=Source File Month | BF=Business Place | BG=SAP Matched By | BH=Flag seq (helper) | BI=Adv Bucket (helper)

## Queries (Draft)
- dims 13x5 | header row 4 | freeze A5
- columns: A=# | B=Last year's query (verbatim, with ro | C=This year's position | D=Proposed query for FY 25-26 | E=Where to verify

## Step 1 Flags
- dims 422x10 | header row None | freeze A10

## GSTR-1 Data
- dims 2081x15 | header row 2 | freeze A3
- columns: A=My GSTIN | B=Tax Period | C=Source | D=Raw Type | E=Document Number | F=Taxable (Net) | G=IGST (Net) | H=CGST (Net) | I=SGST (Net) | J=Bucket (formula) | K=Month (formula) | L=IRN | M=Is Amendment | N=Match Status | O=Matched SR Doc No

## 3B Data
- dims 230x23 | header row 1 | freeze None
- columns: K=RCM 3.1(d) Taxable | L=RCM 3.1(d) IGST | M=RCM 3.1(d) CGST | N=RCM 3.1(d) SGST | O=Net ITC 4(C) IGST | P=Net ITC 4(C) CGST | Q=Net ITC 4(C) SGST | R=4A(3) RCM ITC IGST | S=4A(3) RCM ITC CGST | T=4A(3) RCM ITC SGST | U=4A(4) ISD ITC IGST | V=4A(4) ISD ITC CGST | W=4A(4) ISD ITC SGST

## GL Data
- dims 6204x14 | header row 2 | freeze A3
- columns: A=State | B=Business place | C=Reference (invoice no) | D=FI Document | E=Doc Type | F=Sales-origin? | G=Year/Month | H=Month | I=Tax head | J=Amount (+ = liability) | K=Posting Date | L=Clearing Document | M=Text | N=Matched with SR

## FS Revenue Data
- dims 135x3 | header row None | freeze A3

## S2 Month-on-Month
- dims 1372x24 | header row 4 | freeze E5
- columns: A=State | B=GSTIN | C=Month | D=GSTR-1 Type | E=SR Taxable | F=SR IGST | G=SR CGST | H=SR SGST | I=SR Total Tax | J=GSTR-1 Taxable | K=GSTR-1 IGST | L=GSTR-1 CGST | M=GSTR-1 SGST | N=GSTR-1 Total Tax | O=Diff Taxable | P=Diff IGST | Q=Diff CGST | R=Diff SGST | S=Diff Total Tax | T=Auto: books docs not in GSTR-1 | U=Auto: GSTR-1 docs not in books | V=UNEXPLAINED residual (taxable) | W=Remark (type here) | X=key

## S2 SR vs GSTR-1
- dims 75x35 | header row 7 | freeze C8
- columns: A=GSTIN | B=State | C=Taxable value | D=IGST | E=CGST | F=SGST | G=Taxable value | H=IGST | I=CGST | J=SGST | K=Taxable value | L=IGST | M=CGST | N=SGST | O=Taxable value | P=IGST | Q=CGST | R=SGST | S=Taxable value | T=IGST | U=CGST | V=SGST | W=Taxable value | X=IGST | Y=CGST | Z=SGST | AA=Taxable value | AB=IGST | AC=CGST | AD=SGST

## S2 Pivot Month-on-Month
- dims 1394x18 | header row 6 | freeze None
- columns: A=State | B=Month | C=GSTR-1 Type | D=SR Taxable | E=SR IGST | F=SR CGST | G=SR SGST | H=SR Total Tax | I=GSTR-1 Taxable | J=GSTR-1 IGST | K=GSTR-1 CGST | L=GSTR-1 SGST | M=GSTR-1 Total Tax | N=Diff Taxable | O=Diff IGST | P=Diff CGST | Q=Diff SGST | R=Diff Total Tax

## S2 Exceptions
- dims 94x10 | header row 2 | freeze A3
- columns: A=Status | B=State | C=My GSTIN | D=Books Doc No | E=Books type | F=Books taxable | G=GSTR-1 Doc No | H=GSTR-1 type | I=GSTR-1 taxable | J=Assessment

## S3 Month-on-Month
- dims 233x20 | header row 4 | freeze D5
- columns: A=State | B=GSTIN | C=Month | D=GSTR-1 Taxable | E=GSTR-1 IGST | F=GSTR-1 CGST | G=GSTR-1 SGST | H=GSTR-1 Total Tax | I=3B Taxable | J=3B IGST | K=3B CGST | L=3B SGST | M=3B Total Tax | N=Diff Taxable | O=Diff IGST | P=Diff CGST | Q=Diff SGST | R=Diff Total Tax | S=Remark (type here) | T=key

## S3 1 vs 3B
- dims 24x18 | header row 4 | freeze C5
- columns: A=GSTIN | B=State | C=GSTR-1 Taxable | D=GSTR-1 IGST | E=GSTR-1 CGST | F=GSTR-1 SGST | G=GSTR-1 Total Tax | H=3B Taxable | I=3B IGST | J=3B CGST | K=3B SGST | L=3B Total Tax | M=Diff Taxable | N=Diff IGST | O=Diff CGST | P=Diff SGST | Q=Diff Total Tax | R=Remarks (auto from S3 Month-on-Month

## S3 SR vs 3B
- dims 24x17 | header row 4 | freeze None
- columns: A=GSTIN | B=State | C=SR Taxable (incl adv) | D=SR IGST | E=SR CGST | F=SR SGST | G=SR Total Tax | H=3B Taxable | I=3B IGST | J=3B CGST | K=3B SGST | L=3B Total Tax | M=Diff Taxable | N=Diff IGST | O=Diff CGST | P=Diff SGST | Q=Diff Total Tax

## S3 SR vs 3B MoM
- dims 233x19 | header row 4 | freeze A5
- columns: A=State | B=GSTIN | C=Month | D=SR Taxable (incl adv) | E=SR IGST | F=SR CGST | G=SR SGST | H=SR Total Tax | I=3B Taxable | J=3B IGST | K=3B CGST | L=3B SGST | M=3B Total Tax | N=Diff Taxable | O=Diff IGST | P=Diff CGST | Q=Diff SGST | R=Diff Total Tax | S=Remark (type here)

## S3 Amendment Check
- dims 8x2 | header row None | freeze None

## S4 GL vs SR
- dims 24x16 | header row 4 | freeze None
- columns: A=Business place | B=State | C=GL IGST | D=GL CGST | E=GL SGST | F=GL Total Tax | G=SR IGST | H=SR CGST | I=SR SGST | J=SR Total Tax | K=Diff IGST | L=Diff CGST | M=Diff SGST | N=Diff Total Tax | O=Explained (S4 Exceptions) | P=Residual (CGST diff - explained)

## S4 Exceptions
- dims 11x8 | header row 2 | freeze None
- columns: A=Status | B=State | C=Reference / Doc No | D=Reg type | E=Month | F=Register CGST | G=GL CGST | H=Contribution to (GL - SR)

## S5 HSN Summary
- dims 499x10 | header row 4 | freeze None
- columns: A=State | B=My GSTIN | C=HSN/SAC | D=UQC (as in register; blank shown '-' | E=Rate % | F=Quantity | G=Taxable | H=IGST | I=CGST | J=SGST

## S5 Rate-wise
- dims 26x8 | header row 4 | freeze None
- columns: A=State | B=My GSTIN | C=Rate % | D=Taxable | E=IGST | F=CGST | G=SGST | H=Actual rate %

## S6 Advances Control
- dims 24x10 | header row 4 | freeze None
- columns: A=State | B=GSTIN | C=Opening 01.04.25 | D=Received (books) | E=Adjusted (books, |sum|) | F=Closing | G=GSTR-1 net advances | H=Books net | I=Books vs GSTR-1 | J=Remark

## S6 Month-on-Month
- dims 233x11 | header row 4 | freeze A5
- columns: A=State | B=GSTIN | C=Month | D=Opening | E=Received (books) | F=Adjusted (books, |sum|) | G=Closing | H=GSTR-1 net advances | I=Books net | J=Books vs GSTR-1 | K=Remark (type here)

## S7 CN Time-bar
- dims 292x8 | header row 4 | freeze A23
- columns: A=State | B=CN Number | C=CN Date | D=Taxable | E=Original Invoice No (live) | F=Original Invoice Date (live) | G=Declare-by (30 Nov) | H=Status (live)

## S9 Sales Reco
- dims 43x22 | header row 5 | freeze B17
- columns: A=Reconciliation of turnover declared  | B=Reconciliation of turnover declared  | C=Reconciliation of turnover declared  | D=Reconciliation of turnover declared  | E=Reconciliation of turnover declared  | F=Reconciliation of turnover declared  | G=Reconciliation of turnover declared  | H=Reconciliation of turnover declared  | I=Reconciliation of turnover declared  | J=Reconciliation of turnover declared  | K=Reconciliation of turnover declared  | L=Reconciliation of turnover declared  | M=Reconciliation of turnover declared  | N=Reconciliation of turnover declared  | O=Reconciliation of turnover declared  | P=Reconciliation of turnover declared  | Q=Reconciliation of turnover declared  | R=Reconciliation of turnover declared  | S=Reconciliation of turnover declared  | T=Reconciliation of turnover declared  | U=Reconciliation of turnover declared  | V=Reconciliation of turnover declared 

## RCM Register
- dims 3186x76 | header row 5 | freeze A6
- columns: A=3B Month | B=Final 3B Month | C=Business place | D=MY GSTN | E=State | F=Vikran State code | G=Vendor state code | H=Fiscal Year | I=Year/Month | J=G/L Account | K=GL Name | L=Document type | M=Document Number | N=Posting Date | O=Posting Year | P=Assignment | Q=Reference | R=Document Date | S=Doc year | T=Vendor Code | U=Key | V=GSTN | W=Vendor Name | X=Posting Key | Y=Amount in Local Currency | Z=Local Currency | AA=Tax Code | AB=Clearing Document | AC=Profit Center | AD=BP as per Profit Centre | AE=CC as per Profit Centre | AF=Cost Center | AG=Text | AH=Offsetting Account | AI=Expenditure GL | AJ=GL Correct | AK=Inv. No. | AL=Inv. Date | AM=GST Rate | AN=Taxable Value As per filed return | AO=IGST AS PER GST PORTAL | AP=CGST AS PER GST PORTAL | AQ=SGST AS PER GST PORTAL | AR=Taxable Value as per SAP | AS=IGST AS PER SAP | AT=CGST AS PER SAP | AU=SGST AS PER SAP | AV=Total GST | AW=Inv. Value | AX=Nature of Services | AY=Payment Status | AZ=Rate Check | BA=Query | BB=Query Description | BC=Vendor | BD=My GSTN | BE=As per State | BF=As per Amounts | BG=POS Check | BH=Query | BI=Query Description | BJ=ITC Eligibility | BK=GSTR 3B Claim month | BL=Found in Output GL | BM=Output GL Remarks | BN=ToS (Doc date+61) | BO=ToS (Posting+61) | BP=ToS (earlier) | BQ=RCM due by (20th after ToS month) | BR=Paid via 3B due date | BS=Delay days | BT=Interest @18% p.a. | BU=ToS Status | BV=Found in Input GL (post-ITC) | BW=Input GL Remarks | BX=Found in 2B

## Statewise RCM vs 3B
- dims 26x19 | header row 5 | freeze B6
- columns: A=State | B=Sum of Taxable Value | C=Sum of IGST AS PER SAP | D=Sum of CGST AS PER SAP | E=Sum of SGST AS PER SAP | G=Row Labels | H=Taxable Value | I=Integrated Tax | J=Central Tax | K=State Tax | M=Row Labels | N=Taxable Value | O=Integrated Tax | P=Central Tax | Q=State Tax | R=DPS Remarks

## Month wise RCM vs 3B
- dims 18x19 | header row 5 | freeze B6
- columns: A=Row Labels | B=Sum of Taxable Value | C=Sum of IGST AS PER SAP | D=Sum of CGST AS PER SAP | E=Sum of SGST AS PER SAP | G=Taxable Value | H=Integrated Tax | I=Central Tax | J=State Tax | K=Total Tax | M=Taxable Value | N=Integrated Tax | O=Central Tax | P=State Tax | Q=Total Tax | R=DPS Remarks

## RCM ToS & Interest
- dims 26x8 | header row 6 | freeze A7
- columns: A=State | B=GSTIN | C=Rows | D=LATE rows | E=Tax on LATE rows | F=Interest @18% | G=No-date rows | H=DPS Remarks

## RCM GL
- dims 36627x21 | header row 5 | freeze A6
- columns: A=Side | B=G/L Account | C=GL Name | D=State | E=Business place | F=Document Number | G=Document Type | H=Document Date | I=Doc FY | J=Posting Key | K=Amount in Local Currency | L=Tax Code | M=Clearing Document | N=Profit Center | O=Text | P=Offsetting Account | Q=Assignment | R=Matched with RCM Register (Output) | S=Match Remarks (Output) | T=Matched with RCM Register (Input) | U=Match Remarks (Input)

## Rate-wise summary-RCM
- dims 39x7 | header row 5 | freeze A6
- columns: A=State | B=MY GSTN | C=GST Rate | D=Sum of Taxable Value as per SAP | E=Sum of IGST AS PER SAP | F=Sum of CGST AS PER SAP | G=Sum of SGST AS PER SAP

## RCM vs 2B
- dims 693x17 | header row 5 | freeze A6
- columns: A=Company GSTIN | B=Tax Period | C=Doc Type | D=Doc No | E=Doc Date | F=Supplier GSTIN | G=Supplier Name | H=Supplier State | I=Place of Supply | J=Taxable Value (Net) | K=IGST (Net) | L=CGST (Net) | M=SGST (Net) | N=ITC Eligible | O=Source | P=Match Status | Q=Match Remarks

## TB Scrutiny-RCM
- dims 24x6 | header row 6 | freeze A7
- columns: A=Account | B=Account Name | C=FY 2025-26 Amount (TB) | D=RCM indicator (rule hit) | E=In RCM register? | F=DPS Remarks

## ITC Register 2025-26
- dims 41513x83 | header row 5 | freeze A6
- columns: A=Company | B=Business place | C=State Name | D=VEL GSTIN | E=3B Claim  Month | F=Category | G=Category as per 3B | H=Document Type | I=Document Number | J=Posting Date | K=Posting Year | L=Invoice No. | M=Correct Invoice no. | N=Invoice Date | O=Correct Invoice date | P=Invoice Year | Q=Vendor Code | R=Vendor GSTIN | S=Correct GSTIN | T=Vendor Name/RCM Category | U=Tax Rate | V=Taxable Value | W=IGST | X=CGST | Y=SGST | Z=Total GST | AA=GSTR 9C_Reporting | AB=Reasons | AC=Matching of 12B of FY 25-26 and 12C  | AD=Invoice Level Match | AE=KEY | AF=Countif | AG=B_IGST | AH=B_CGST | AI=B_SGST | AJ=B_Total GST | AK=KEY2 | AL=2B_IGST | AM=2B_CGST | AN=2B_SGST | AO=2B_Total GST | AP=D_IGST | AQ=D_CGST | AR=D_SGST | AS=D_Total GST | AT=Reco Remarks | AU=Review Remarks | AV=Invoice as per 2B | AW=2B Period | AX=My GSTN as per 2B | AY=Supplier GSTN as per 2B | AZ=Nature of Services | BA=Type for GSTR9 | BB=G/L Account | BC=G/L Account Text | BD=SAP Period | BE=GSTR 2B/6A Period | BF=2B Year | BG=Profit Center | BH=Project Code | BI=RCM Paid Month | BJ=Remarks | BK=Remarks 2 | BL=Comments | BM=Vendor | BN=My GSTN | BO=As per State | BP=As per Amounts | BQ=POS Check | BR=POS Query | BS=POS Query Description | BT=TYPE | BU=Material Description | BV=Expense GL Element | BW=Expense Description | BX=Eligibility | BY=Query | BZ=Consider 8A reco | CA=Considered in Table 6A1 | CB=Remarks for accounting entries- For  | CC=GST CREDIT YES/NO | CD=F.Y/Booking Year | CE=Company Code

## GSTR-2B ITC Data
- dims 48807x27 | header row 5 | freeze A6
- columns: A=State folder | B=FY (derived) | C=F.Y | D=2B Return Period | E=GSTR 3B Month | F=Curent previous | G=GSTIN of supplier | H=Trade/Legal name | I=Invoice number | J=Invoice type | K=Invoice Date | L=Invoice Value(₹) | M=Place of supply | N=Supply Attract Reverse Charge | O=Rate | P=Taxable Value (₹) | Q=Integrated Tax(₹) | R=Central Tax(₹) | S=State/UT Tax(₹) | T=Cess(₹) | U=GSTR-1/5 Period | V=ITC Availability | W=Reason | X=Source | Y=IRN | Z=Claimed in FY 25-26 register? | AA=Claim month (derived)

## ITCR vs 3B Net ITC
- dims 234x18 | header row 5 | freeze D6
- columns: A=State | B=GSTIN | C=Month | D=Register IGST | E=Register CGST | F=Register SGST | G=Register Total | H=3B Net ITC IGST | I=3B Net ITC CGST | J=3B Net ITC SGST | K=3B Net ITC Total | L=Diff IGST | M=Diff CGST | N=Diff SGST | O=Diff Total | P=DPS Remarks

## ITCR vs 2B GSTN Level
- dims 24x15 | header row 5 | freeze C6
- columns: A=State | B=GSTIN | C=Register ITC claimed IGST | D=Register ITC claimed CGST | E=Register ITC claimed SGST | F=Register ITC claimed Total | G=2B available FY25-26 IGST | H=2B available FY25-26 CGST | I=2B available FY25-26 SGST | J=2B available FY25-26 Total | K=Diff (claimed - 2B) | L=2B unclaimed-in-register (tax) | M=DPS Remarks

## Unclaimed ITC candidates
- dims 7010x15 | header row 5 | freeze A6
- columns: A=State folder | B=FY (derived) | C=2B Return Period | D=GSTIN of supplier | E=Trade/Legal name | F=Invoice number | G=Invoice Date | H=Taxable Value (₹) | I=Integrated Tax(₹) | J=Central Tax(₹) | K=State/UT Tax(₹) | L=Total Tax | M=ITC Availability | N=Supply Attract Reverse Charge | O=Status

## RCM Paid vs ITC Claimed
- dims 234x12 | header row 5 | freeze A6
- columns: A=GSTIN | B=State | C=Month (3B) | D=RCM paid - Total tax (RCM Register) | E=4A(3) claimed same month (3B) | F=4A(3) claimed NEXT month (3B) | G=ITC Register RCM claims - month M+1 | H=Diff (paid - next-month 4A(3)) | I=DPS Remarks

## 2B ISD Data
- dims 147x13 | header row 4 | freeze None
- columns: A=State folder | B=Company GSTIN | C=Tax Period | D=Doc Type | E=Doc No | F=Doc Date | G=GSTIN | H=Trade Name | I=IGST | J=CGST | K=SGST | L=Cess | M=ITC Eligible

## Open Points (all)
- dims 40x6 | header row 2 | freeze A3
- columns: A=Step | B=# | C=Owner | D=Point | E=Detail | F=Resolved?

## CC Master  (hidden)
- dims 125x2 | header row None | freeze None

## TB Groupings  (hidden)
- dims 786x2 | header row None | freeze None
