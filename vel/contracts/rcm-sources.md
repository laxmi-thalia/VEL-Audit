# VEL — RCM phase source contracts (FY 2025-26 baseline)

Same doctrine as sources.md: year-roll DIFFS new files against this; drift STOPS the build
for one-line rulings. Engagement root as in sources.md.

## 1. Conso RCM (the register's source)

`DPS Workings\2. RCM Register\Conso RCM FY 25-26.xlsx`, sheet `RCM OUTPUT WORKING`,
header row 2, 42 cols, **3,181 real rows** (used range shows ~322k — PHANTOM ROWS; always
filter `Document Number` notna). GL-posting grain: one row per CGST/SGST/IGST RCM output
posting. Both sides present (as-per-SAP + as-per-portal) — tied to the paisa at totality
in FY 25-26. 3B Month style `01 Apr 2025`; spillover real: Mar-25 postings (fiscal 2024/12)
carry 3B Month Apr-25.
- **State names carry TYPOS** ("Madya Pradesh", "Tamilnadu"/"Tamil Nadu") — key on GSTIN
  or Business place, NEVER on the state label.
- GST Rate carries float dirt (17.999999999999996) — SUMIFS rate criteria need ±0.01 bands.
- Doc types RE/R3/KR/R7/KG + ~28 blank; 2 rows Business place = HOIS (ISD, no GSTIN).
- G/L Account values may carry suffixes ("2610080300-01") — normalize to leading digits.

## 2. Target format

Rashid's draft: `DPS Workings\2. RCM Register\Format_RCM Register_FY2025-26.xlsx` (58 cols).
**The REAL template is last year's final** `RCM Register` sheet in the FY 24-25 xlsb —
72 cols = the draft + portal block (before SAP block) + `Final 3B Month` (date), `Key`
(=MY GSTN & vendor GSTN), Local Currency. Live formulas: Key, Total GST (=SUM SAP taxes),
Rate Check (=TotalGST/SAPtaxable*100), POS block (LEFT-2 codes, Intra/Inter by state
equality and by IGST=0, POS Check = equality). Stamped: GSTIN/state/BP-CC-per-PC/fiscal years.

## 3. GSTR-3B RCM figures (step 2)

Octa files `DPS Workings\Portal Reports\GSTR-3B\` — section **3.1.D** rows (parser:
`b3d_extract.py`, clone of the 3.1.A sales parser). Zero rows omitted by Octa (Kerala 0/4,
Punjab 2/4, TN & WB 3/4) — default zeros. **Haryana file MISSING entirely.**
Figures land as columns K-N (`RCM 3.1(d) *`) on the master's 3B Data sheet.
ITC leg (4A(2)/4A(3)) NOT yet extracted — post-ITC.

## 4. RCM GL dumps (step 4)

`Clients Data\GLs\Inward GL\RCM- Output-All Items.xlsx` + `RCM-Input.xlsx` — SAP FBL3N
"all items" exports: preamble rows 1-5 (`G/L Account: *`), header row 6, 33 cols (most are
junk currency columns), **account markers are trailer rows "Account NNNNNNNNNN" at block
END** (output: 2610080300/301/302; input: 1910060500/501/502). Span 2.5 fiscal years.
- **TRAP: SAP document numbers RECYCLE each fiscal year** — match key MUST be
  (account, doc no, doc-date FY); keying without FY produced 710 false mismatches (real: 4).
- Transfer entries identified via Clearing Document (14,191 rows knocked off).
- No Posting Date column — Assignment (yyyymmdd) is the proxy.

## 5. GSTR-2B (step 6)

`DPS Workings\Portal Reports\GSTR-2B\GSTR2B-Net-...-<State>-Apr 2025-Mar 2026.xlsx` —
sheet `Purchase-Net` (45 cols incl `Reverse Charge` Yes/No, Net value columns, ITC Eligible,
IMS columns) + `ISD` sheet. Kerala file is plain `GSTR2B-` prefix with `Purchase` sheet (nil).
**Files MISSING for Haryana, West Bengal, Tamil Nadu, Telangana.**
Filter Reverse Charge = Yes (688 items FY 25-26). Match cascade: (supplier GSTIN,
normalized invoice no) → (supplier GSTIN, rounded |taxable|). URD/self-invoiced RCM never
appears in 2B — three-way classification, not a defect.

## 6. Trial Balance + Groupings (step 7, Expenditure GL)

**Inside** `Clients Data\Financials\VEL Standalone FS Mar-26 with Notes.xlsx`:
- `Groupings` sheet: header row 4, col A = 10-digit account, col B = FS account name →
  the **Expenditure GL** lookup (CA ruling: by Offsetting Account). Embedded as hidden
  `TB Groupings` sheet in the master; register column is live INDEX/MATCH.
- `Trial Balance` sheet: per (account × profit centre × period) rows; account name col J,
  FY amount col K. Expense series = accounts matching ^4\d{9}.
- CAVEAT: 3,047/3,181 register rows offset to "SR/IR Clearing" (+39 Freight Clearing) —
  the expense-head check is only direct for ~134 rows; the rest need a clearing-doc hop.

## 7. Audit checklist (the step order)

`VEL\Audit Checklist FY 25-26.xlsx` sheet CheckPoints, header row 4: filter Type=RCM,
follow the `Steps` column: 1 rate/POS · 2 vs 3B (liability & ITC) · 3 ToS+interest ·
4 GL O/I vs register · 5 rate-wise · 6 vs 2B · 7 TB scrutiny. NA rows: import of services,
purchase-register/expense-GL scouting (govt contracts), directors' guarantee.

> TRAP UPDATE: state-name typos now include 'Tamilnadu' (lowercase n) in the ITC All
> State Working - THREE spellings of Tamil Nadu exist across client files.
