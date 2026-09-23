---
name: gst-audit-vel-itc
description: VEL (Vikran Engineering) GSTR-9/9C annual-return audit — ITC phase. The complete recipe for the ITC Register, the cumulative 2B merge/matching, the four ordered checklist steps, the unclaimed-ITC October list, and the RCM post-ITC hooks, all inside VEL_GST_Audit_FY2025-26_MASTER.xlsx. Use for any work on the VEL ITC sheets or the FY 26-27 ITC roll-forward. Siblings: gst-audit-vel-sales, gst-audit-vel-rcm.
---

# VEL GST audit — ITC phase

Same engagement/master as the sales and RCM skills — **their Standing rules all apply**
(formulas-only, no LLM without asking, file-lock etiquette, verify-before-reporting,
re-snapshot before edits, user edits win).

Engine home: `vel/` (this repo: C:/PROJECTS/gst-audit-engine)
- `contracts/itc-sources.md` — **read first**: every source with structure + traps.
- `contracts/goldens_fy2526_itc.md` — backtest targets (gitignored; no client figures here).
- `contracts/formats.md` / `reference/formats.json` — layout authority (regenerate on change).
- `scripts/` — itc0_extract/itc0_build (register), b2itc_merge (2B), itc1_match /
  itc1_apply / itc1_apportion / itc1b_fallback (invoice-level reco + ISD/fallback),
  b3itc_extract / itc2_build (vs 3B Net ITC), itc4_prep / itc4_build (GSTIN level +
  October list), rcm_postitc (the RCM hooks).

## The design in one paragraph

The register (41,508 rows, last year's 83-col layout) holds every FY 25-26 claim with its
3B month. The cumulative 2B (48k+ rows, FY 2020-21 onward) is the matching universe: an
invoice-level cascade — (vendor GSTIN, normalized invoice) → (vendor, date, tax) →
(vendor, tax, ≤3 candidates) — derives, for every 2B row, whether and when it was claimed
(the client's own claim marking exists only for Bihar/TN). From that one matching pass
fall out: step 1's data-accuracy columns (correct invoice/date/GSTIN, 2B period), the
B_/2B_/D_ value blocks (2B doc totals apportioned across register lines), the GSTIN-level
"Less in 2B" view, and the unclaimed-candidates October list (filtered against last year's
8,227 claim keys). The vs-3B step computes Net ITC = 4A + 4B-as-reported and ties the
register month-by-month.

## What was built (sheets in the master)

- `ITC Register 2025-26` — 84 cols (83 + GST CREDIT YES/NO etc appended). Live: Total GST,
  POS block (step 3 live from birth), KEY, B_ block, D_ diffs. Stamped: match verdicts,
  2B_ apportioned values, correct-invoice block, Countif (live COUNTIF over 41k rows =
  O(n²) recalc — deliberately stamped, regenerate on data change).
- `GSTR-2B ITC Data` — the merged cumulative 2B + derived "Claimed in FY 25-26 register?" +
  claim month; fallback rows source-tagged. `2B ISD Data` — the ISD section.
- `ITCR vs 3B Net ITC` — 228 GSTIN-months live both sides. FY 25-26 RESULT: register
  claims vs 3B Net ITC differ by ~1.5L for the whole year (near-perfect).
- `ITCR vs 2B GSTN Level` — per GSTIN live: claimed vs 2B available, diff ("Less in 2B"),
  unclaimed-in-2B column, DPS Remarks.
- `Unclaimed ITC candidates` — THE OCTOBER LIST: FY 25-26 2B credits unclaimed in the
  register (claimable till Nov 2026, Sec 16(4)) + FY 24-25 rows unclaimed in either
  register. Standing caveat printed on-sheet: FY 26-27 Apr-Jul claims not yet netted.
- `RCM Paid vs ITC Claimed` — RCM output paid (month M) vs 4A(3) claimed (M and M+1) vs
  ITC-register RCM claims; VEL claims RCM ITC ONE MONTH BEHIND payment; Mar-26 spills to
  FY 26-27. RCM register + RCM GL input columns stamped from the same claims.

## CA/Pawan rulings (changes log — append every new one)

- Register source = ITC All State `Working`; drop "Others"; missing data fetched from
  monthly workings with a flag first.
- 2B source = the `GSTR 2B` sheets INSIDE the July FY 26-27 state working files (user
  note), July = latest cut.
- Checklist order = the `Steps` column; each step's inputs = `File to be used` column.
- Countif stamped, not live (performance); 2B_ apportioned per last year's semantics.
- Claim months for 13 states DERIVED by matching (client marking absent) — Bihar/TN
  markings kept as cross-check only.

## Traps (do not relearn)

- Working sheet stores AMOUNTS AS TEXT — coerce everything; a text-number column sums to
  zero silently in a values-only survey.
- **COM auto-coerces month-like strings ("Apr-25", "02 May 2025") to DATES** on write —
  criterion cells then never match text; set NumberFormat "@" before writing tokens.
- pd.NaT IS a datetime instance — guard `v is pd.NaT` FIRST, before isinstance checks.
- 4B reversals negative as reported (Net = 4A + 4B); 4(C) is a header-only row.
- F.Y column half-blank — derive FY from `2B Return Period`.
- 2B invoice numbers carry leading apostrophes; normalize alphanumerics-only, uppercase.
- Four state-name spellings across files — canonicalize to GSTIN on ingestion, always.
- 3B Data's added columns have headers on ROW 1 (reserved row) vs originals on row 2 —
  resolve headers by scanning BOTH rows (cosmetic cleanup pending).
- A cumulative 2B's unmatched old-FY rows are NOT "unclaimed" — they were claimed in
  prior years' registers; judge only after netting adjacent-year claims.
- Register lines are line-level, 2B is doc-level — many-to-one is normal, not a defect.

## Missing data / open flags (keep current)

- FY 26-27 Apr-Jul claims not netted from the candidates list (the one refinement before
  the client acts on it).
- Arunachal/J&K/Kerala: no 2B sheet in workings (fallback = FY 25-26 window only; their
  FY 24-25 2B history unavailable). Haryana: no folder in FY 26-27 Final at all.
- 617 register rows blank GST-CREDIT flag; 523 blank 2B-periods; `Type for GSTR9` empty.
- 28 ISD rows (~5.8L) still unmatched; ~11.75cr register claims with no 2B counterpart
  (UP 5.25cr, MP 1.92cr, Chh 1.53cr leading) — CA review via Invoice Level Match filter.
- Receivables open items (42,429 rows) pending review (Sr 58/59) — no vendor columns.
- 37 unordered checkpoints (Sr 27-65 minus ordered) — many answerable from built data
  (8A recon, Table 6A1, TDS, high-risk states, Bird's eye view...), several inquiry-type.
- RCM register judgment columns still pending: ITC Eligibility (mine last year's
  nature→Y/N as tagged proposals), GL Correct (semantic pairs).

## Tax Comparison Report (built 2026-08-28)

`Tax comp report` sheet: merged `ITC (Other than IMPG)` from the portal
`Portal Reports\Tax Comparison Reports\2025-26_<GSTIN>_Tax liability and ITC comparison.xlsx`
files - VEL PAN (AAECR0503Q) only, subsidiary files excluded; month rows only; columns till
the Shortfall block (cumulative blocks dropped - user ruling). Two-tier header like last
year's. Source sheets in each file: Tax Liability Summary / Comparison Summary /
Tax liability / Reverse charge / Export and SEZ / ITC (Other than IMPG) / ITC (IMPG) /
RCM_LIABILITY_ITC (data rows start row 7, months as 'Apr-25').
FLAG: **TCR files missing for Haryana AND Andhra Pradesh.**
Bird's eye view (Sr 40): BUILT as standalone file 'Birds Eye View VEL GSTR 3B FY25-26.xlsx' (template: FY 2024-25/1. Corporate Clients/TP EV/Birds Eye View/...New Formula.xlsx - 18 state sheets of the Octa matrix + PAN consolidated + live Liability Summary; script bev_build.py). TRAP: never copy style-dump number formats truncated - invalid codes make Excel refuse the whole workbook; snapshot before any style-replica build.

## Year-roll & maintenance

Template-copy → survey vs contracts → rulings → rebuild → backtest against
`goldens_fy2526_itc.md` to the rupee → gauntlet. Append rulings + flags here after every
session; regenerate formats.md; sync scripts.

## Changes log - 2026-09-17/18: ITC Batch 1 ("till 6A1") per Rashid/Purvi + Priyesh meetings

Transcripts (Downloads): `Meeting with rashid faisal-20260917_103824UTC - Transcript (English).md`,
`GST Audit -Vikran-20260917_121606UTC - Transcript (English).md`; plan `VEL ITC Reporting Layer - Plan v2.docx`.
Built on `VEL_GST_Audit_FY2025-26_MASTER (2).xlsx`; scripts b2_octa_merge / itc_batch1_register / itc_batch1_stepC / itc_batch1_stepD.
- **2B source = Octa exports in `DPS Workings\Portal Reports\GSTR-2B\`** (14 state files Apr-25..Mar-26 +
  all-states FY 26-27 file), merged as `GSTR-2B Apr25-Aug26` (raw 45 cols verbatim, DPS cols LAST) +
  `2B ISD Apr25-Aug26`. Window ruling (Pawan/Priyesh): Apr-25 -> Aug-26; the FY 26-27 export currently
  holds Apr..Jul-26 only. FLAGS: no FY 25-26 file for TN, TG, WB, HR; Kerala file empty (non-Net export).
- **Register (75 cols now)**: removed Category as per 3B, Invoice Level Match (text -> single `Reco
  Remarks` after D_Total), Review Remarks, RCM Paid Month/Remarks/Remarks 2/Comments, POS Query
  Description, Query; added `PO Number`. KEY = UPPER(SUBSTITUTE chain) of vendor GSTIN & invoice (same
  normalisation as the 2B KEY formula). `Countif` = Consider/NA STAMPED (first line per matched 2B doc;
  live running COUNTIF over 41k rows is O(n^2)). **B_ block = SUMIFS over the register by KEY2 on Consider
  lines; 2B_ block = SUMIFS over the 2B sheet by KEY2** (both CAs: one 2B doc across several register
  lines -> total on one line, NA on the rest). KEY2 = the 2B key the cascade matched (stamped). "As per
  2B" cols = INDEX/MATCH into the 2B sheet. POS Check TRUE/blank + POS Query text. Expense GL Element /
  PO Number = INDEX/MATCH into hidden `ZFI06 Data` by Document Number; Expense Description from `TB
  Groupings` by GL code.
- **Cascade re-run** vs the Octa 2B: 30,945 by invoice no, 100 date+amount, 175 amount(<=3), 8,734 not
  found, 472 no vendor GSTIN; **1,082 rows match only the FY 24-25 2B** (old `GSTR-2B ITC Data`, working-
  file basis) = 6A1 component 1; verdict text carries the source.
- **2B sheet vice-versa (live)**: 3B Claim Month + Reco Remarks by KEY2 from the register; 6A1 mark; Table
  8A (Octa's `GSTR-9 (8A) ITC Available` + FY + RCM/amendment); 8A Reco; Table 8C (25-26 dated in 26-27
  2B); Table 13; Final Remarks; GSTR-9/9C. Permanent Reversals / Reclaim 6H / Query amber for CA.
- `Unclaimed ITC candidates` REMOVED (Rashid: it is 6A1-unclaimed). `ITCR vs 3B Net ITC` rebuilt SPLIT
  RCM (4A3) / ISD (4A4) / Other (4A5+4B): ties 2.05cr/11.18cr/93.75cr; net diff 154,184.32 (golden);
  only Arunachal Oct-25 (+62,100) and WB Dec-25 (+92,056) exceed Rs 5,000.
- **`T6A1 Extract - 24-25`** (last year's 19 cols + helper): 3,579 rows = 1,082 (24-25 inv in 24-25 2B
  availed 25-26) + 2,029 correction entries (24-25 dated, in no 2B) + 179 claimed + 275 unclaimed (24-25
  dated in 25-26 2B) + 14 RCM Mar-25->Apr-25 (RCM register fiscal 2024/12). GSTR-9 Remarks LIVE (links to
  the source sheet's mark); amounts values (regenerate by script); PivotTable PT_6A1 (State x Remarks).
  Register `Considered in Table 6A1` = Yes + `Remarks for accounting entries` tag on 3,111 rows (proposals).
- TRAPS: ITC register `3B Claim  Month` is the LABEL coerced to a date - May = 02-May-2025, Jun = 03-Jun..
  (day = month index); month helpers must replicate that or SUMIFS reads zero for 11 months. ZFI06 as
  supplied covers only 1,717 documents (KH type, 85% BR/UP) -> 108 register rows get Expense GL/PO;
  ask client for a full-year all-documents ZFI06 run or approve the Reference+Vendor GSTIN key.
- Deferred by CAs: Eligibility, Material Description, Query, Table 12; batch 2 = ITC Summary (gross-4A5
  auto-populate column blank), Table 13 & 6A1 (Table 13 from FY 24-25 filed GSTR-9), T12B/T12C, 8C vs
  13-12, Tax comp reasons (Computation sheet vs filed 3B, rule-based, no LLM).

## Changes log - 2026-09-18: ITC Batch 2 step E (ITC Summary + Table 13/12C sheets)

Scripts: itc_batch2_stepE.py (+ zip_bisect.py diagnostic). Built on MASTER (2), verified 0 error cells.
- Last year's formula logic replicated exactly (read via COM from a LOCAL copy of the xlsb - COM cannot open
  UNC paths): every 6A1 block = SUMIFS over `T6A1 Extract - 24-25` (My GSTIN col I, GSTR-9 Remarks col G,
  P/Q/R); 6B = 4A5 - 6H - 6A1 + RCM6A1; 6D = 4A3 - RCM6A1; 6G = 4A4; 6H = 4D1 - claimed(PY 2B) - correction;
  6J = 6A - (6A1+6B+6D+6G+6H); 7H = -(4B1+4B2) (Octa negative); 7J = 6B+6D+6G+6H-7H; Difference = 7J - Net3B
  + 6A1 (closes to 0.00 in all 19 states = the consistency check); 8A from 2B `Table 8A`="Yes" until the
  system-generated GSTR-9; 8B = 6B; 8C from 2B "Table 8C of GSTR-9"; 8D = 8A-8B-8C; "As per Tax Comparison
  Report" = SUMIF of Tax comp cols AE:AG by GSTIN; reasons blocks = SUMIF of V:X / Y:AA / AB:AD.
- Table 13 / 12C of FY 24-25 = hidden `LY 24-25 claims` (last year's `ITC Register 2025-26` sheet, 23,649
  rows; 9_Reporting=13 -> 8.03cr, 9C_Reporting=12C -> 3.96cr). Tie to filed PDFs when all 19 arrive.
- Register `GSTR 9C_Reporting` (12B if Invoice Year ends "24-25", NA if "25-26" - the column holds BOTH
  "2025-26" and "25-26" formats), `Reasons`, `Matching of 12B..12C` (COUNTIFS into LY KEY) - live.
- FINDINGS: 12B (CY) 8.20cr vs 12C (PY) 3.96cr - 1,701 of 4,699 rows not in last year's 12C list (CA);
  8D negative (-1.8cr IGST) because TN/TG/WB/HR have no FY 25-26 2B file (8A understated).
- **TRAP (cost an hour): Excel REFUSES TO OPEN (Open method failed, even CorruptLoad) a workbook holding a
  formula with a string literal > 255 chars.** openpyxl writes it happily. Split long note text into <=200-
  char chunks joined with & (last year's CONCATENATE did exactly that). Diagnose with zip-level sheet
  removal (zip_bisect.py: seconds per trial) - NOT openpyxl round-trips (5 min each, and they drop the
  T6A1 PivotTable body -> re-create PT_6A1 via COM after any openpyxl save).
- COM `.Address` is a property (not callable) via dynamic dispatch - compute letters with openpyxl L().
- Step F (Tax comp reasons: state `Computation` sheet vs filed 3B, rule-based) pending - share down.

## Changes log - 2026-09-18: last year's ITC sheets replicated formula-for-formula (Pawan ruling)

Ruling: "refer previous year's 9C (ITC Summary / ITC Register 24-25 / ITC Register 25-26 / Tax comp report /
T6A1 Extract 23-24 / Table 13 & 6A1) - how each formula was made - copy the exact format, rename sheets to the
new year, build it." Method: `ly_dump.py` dumps every header/style/width/merge + row-5 formula of the six sheets
from a LOCAL copy of the xlsb via COM -> `ly_itc_dump.pkl/.txt` (reference/). Then:
- `replica_itc.py` = layouts/headers/fills/number formats/widths/merges (last year's row n -> our row n+1, row 1
  reserved), year strings rolled (23-24->24-25, 24-25->25-26, 25-26->26-27 in labels only).
- `replica_formulas.py` = the formulas, written EXPLICITLY per column (regex-rolling last year's formula text
  mangled references: rolled sheet names became [1] external links, '$C5' stayed on row 5, the 2B column map was
  applied twice). Sheet mapping: Consolidated GSTR-3B Extract -> 3B Data (GSTIN x month, SUMIFS by GSTIN = the
  FY); GSTR-2B Apr 24-Oct 25 -> GSTR-2B Apr25-Aug26; T6A1 Extract - 23-24 -> - 24-25; ITC Register 2024-25 ->
  2025-26; ITC Register 2025-26 (= NEXT-YEAR claims sheet, Table 13/12C detail) -> NEW `ITC Register 2026-27`
  (49 cols, last year's layout; rows = FY 25-26 invoices in the FY 26-27 2B until the client's FY 26-27 working).
- ITC Summary now 129 cols exactly as last year: 6A(blank per Priyesh) | 6A1 x5 | 6B..6J | 7H | 7J | Net3B |
  Diff | 8A (from 2B until system GSTR-9) | 8B | 8C | 8D | TCR | reasons | T12 Unclaimed | T13 | T13-T12 | 12A |
  12B | 12C | 12D | 12E | 12F | Reasons/Net diff/Comment (9C note text, chunked) | Add to 4D1 | amended | adj.
  Last year's hard plug "-11650" in CN dropped (no plugs). 7H = -(4B1+4B2) because Octa reports 4B negative.
- FY 25-26 RESULTS: Total 6A1 9.66cr (V); 6B 99.76cr; 7J 97.31cr; Net-ITC check = 0.00 all states; 8D -1.32cr
  (8A understated: TN/TG/WB/HR have no FY 25-26 2B file); 12F 9.68cr = the 6A1 format variance (DM net -1.17L);
  Table 13 (PY 8.03cr) vs 6A1-less-unclaimed 8.23cr -> +20.6L, 8 states Matched / 11 for CA.
- TRAPS: (1) openpyxl saves DROP the PivotTable body but keep its definition -> Excel refuses to open; strip the
  orphan pivot parts at zip level, then re-create PT_6A1 via COM (both steps scripted). (2) Copying last year's
  cell styles onto the extract sheet produced a locale-tagged date format + row-level style that Excel rejected -
  rebuild such sheets with plain styling. (3) openpyxl column_dimensions can emit OVERLAPPING <col> ranges
  (min=3 max=5 then min=4) - also fatal; normalise per column. Diagnose with zip_bisect.py (1 min/trial).
- T12B_T12C and 8C-vs-13-12 sheets (my earlier additions, not in last year's file) removed.

## Changes log - 2026-09-18 (evening): NA verification + cascade fix (Pawan)

Pawan challenged the NA population ("too harsh, gaps, blunders"). Independent brute-force re-classification of all
34,997 NA rows on ITC lines proved 26,832 correct (25,724 other lines of a matched doc - each verified against its
Consider line's B_ SUMIFS = full document sum, 0 orphans / 0 mismatches / 0 double-Consider; 1,082 FY 24-25-2B;
26 no GSTIN) and found: (a) 32 rows wrongly linked by the amount / date+amount fallbacks to a 2B doc already owned
by an exact match or of a different invoice year (SDIP/24-25/63 -> SDIP/25-26/007 was the reported case);
(b) 279 "not found" rows that exist in the working-file 2B for TN/TG/WB/KL/HR (no Octa FY 25-26 export - the
full Octa report Pawan is sending supersedes); (c) 175 e-invoice rows in the working files that are not 2B.
CASCADE RULES NOW (cascade_fix.py): leading-zero-insensitive keys (07 == 007; KEY2 still stamped as the 2B row's
own KEY so SUMIFS hit); fallbacks may only take 2B docs NOT owned by an exact match, each once; fallback must
respect invoice FY; fallback verdicts end "- invoice no differs, review". Countif column relabelled: Consider /
Not consider - already considered in the Consider line of this document / Not consider - matched in FY 24-25 2B
(Table 6A1) / Not consider - no vendor GSTIN / Not in 2B (Apr-25 to Aug-26) / Not applicable - RCM|ISD line.
The B_/2B_/D_ formulas test only ="Consider", so labels are free text.
RESULT: exact matches 30,945 -> 31,515 (zero-insensitive), fallbacks 275 -> 233 (all review-flagged), not-found
8,734 -> 8,205, Consider docs 5,464 -> 5,516; books vs 2B on matched docs 79.07cr vs 78.83cr, net diff 47.35L ->
24.52L; the 32 suspects: 25 now not-found, 5 legitimate amount matches, 1 exact, 1 FY 24-25 (SDIP/24-25/63).
T6A1 Extract rebuilt on the corrected verdicts (3,595 rows; RCM component = ITC-register RCM lines claimed
Apr-25 = last year's method); ITC Summary extract ranges repointed; 6A1 total 9.58cr; Net-ITC check 0.00.
TRAPS: COM bulk Value= write fails "OSError 22" on naive datetimes -> write Excel serials + NumberFormat;
openpyxl read-only header maps are 1-based - index value tuples with [c-1] (an off-by-one silently emptied the
FY 24-25 lookup once). Pending: full Octa 2B report from Pawan -> re-merge + re-run cascade_fix.py.
- **RULING (Pawan 18-09, later): `Countif` is the de-duplication flag ONLY** - "Consider" on the first register
  line of each matched 2B document, "Not consider" on the other lines of that document, BLANK everywhere else.
  Status text (NOT FOUND IN 2B, No vendor GSTIN, Matched in FY 24-25 2B, review flags) lives in `Reco Remarks`
  only. Never put "Not in 2B" wording in Countif again. Counts: Consider 5,516 / Not consider 26,232 / blank 9,760.
- **RULING (Pawan 18-09, final on Countif): de-duplication runs on the REGISTER'S OWN invoice number, independent
  of 2B.** Key = vendor GSTIN (or vendor name when no valid GSTIN) + normalised Invoice No.; first line = Consider,
  other lines of that invoice = Not consider; RCM/ISD lines blank. So a document NOT found in 2B still gets its
  Consider line: `B_` = SUMIFS of the register by own KEY (+ vendor name) on Consider lines = the books total of
  every ITC document (ties to Category=ITC tax 93,75,28,722.19); `2B_` = SUMIFS of the 2B sheet by KEY2, **0 when
  KEY2 is blank** (not found); `D_` = B_ - 2B_. Counts: Consider 6,830 / Not consider 34,061 / blank 617.
  Result: B_ 93.75cr vs 2B_ 78.81cr, D_ 14.94cr (= not-in-2B documents + matched-doc differences).
  Deliverable also exported as `.xlsb` (COM SaveAs FileFormat=50; ~40s; 52MB -> 24MB) - the .xlsx stays the
  working copy (openpyxl cannot write xlsb); re-export after every change.

## Changes log - 2026-09-18 (night): revised FY 26-27 2B, GL Key, Tax comp Reasons (Pawan)

- **2B re-merge (b2_remerge_2627.py):** revised file `Audit data of FY 2026-27\GSTR-2B` is FY 26-27 only (Apr-Aug 26,
  3,737 rows, 16 GSTINs); FY 25-26 state files unchanged. 2B sheet now rows 3..13869 (was ..12127); 293,933 dependent
  formulas re-bounded. Chain after an openpyxl save: strip_orphan_pivots.py -> readd_buttons2.py -> cascade_fix.py ->
  final_countif_rule.py (the FINAL Countif rule, saved as a script; the 2B bound is read from the sheet). Not-in-2B
  8,205 -> 8,172 only: TN/TG/WB/KL/HR still have no FY 25-26 Octa export. Verified: 0 error cells, golden
  1,06,98,71,702.15, B_ 93,75,28,722.19 = ITC-category tax, 2B_ 78.85cr, Net-ITC check 0.00.
- **Tax comp report Reasons (tcr_comp_extract2.py + tcr_reasons_lib.py + tcr_reasons.py):** FACT: the sheet's 3B column
  equals filed 4A(4)+4A(5)-4D(1) in all 204 rows (the header text says -4B(2); not changed - flagged to Pawan). The CA's
  Computation New (PART A) method: filed 4A5 = '2B current month ITC as per portal' (net of CN) + reclaim + permanent
  reversal; 4D1 = reclaim + permanent reversal; 4B2 = carry-forward; 4B1 = permanent reversal. So every rupee of
  3B-vs-2B difference is one of: 4D1 deviation (S-U, 'Add to 4D1'), portal-2B vs working-2B (V-X, or AB-AD when it is
  the CN netted in 4A5), ISD / other 4A5 deviation (Y-AA), amounts reported in both 4A5 and 4D1 beyond the working or
  in neither (net off, text only). Extractor anchors value columns on the nearest IGST/CGST/SGST header row; the J&K
  sheet is 'Computation New ' (trailing space) - regex the name. Result: 166 Matched, 30 explained (0 residual), 8 MP
  months 'no working file' -> Pending. Impact 'Ignore' only for CN-netting / reconciliation-tied items; unexplained
  4A5/2B deviations are 'Pending' with both figures cited. Last-year vocabulary reused (Matched / Not reported in 4D(1)
  / Not reported in 4D(1)-Negative / Mismatch in 2B amounts / CN netted in 4A(5) / Not reported in 4A(5)-Negative).
  AE-AG (sheet formula, excludes AB) totals 42.8L IGST = MP's unexplained months + the 2 CN rows.

## Changes log - 2026-09-18 (late): MP monthly files, RCM lines at document level, ZFI06 coverage (Pawan)

- **MP working files are named `GSTR 3B Madhyapradesh <Mon> <Year>.xlsx` (space, not hyphen)** - every `gstr-3b` prefix
  match skipped MP for 12 months. Pattern is now `gstr[ -]?3b` in rcm_monthly_survey.py and tcr_comp_extract2.py.
  RCM Register: MP Jul-25 (94 posting rows) inserted, MP Jan-26 Conso (37) replaced by the monthly working (71) - both tie
  to 3B 3.1(d) to the rupee (rcm_add_mp.py, COM insert/delete so ranges follow). Statewise RCM vs 3B: -12.75L -> +5.98L,
  of which HOIS (2 rows, no GSTIN) 5.50L / 99,000 tax; Gujarat +94,649 / TN +7,500 / Telangana -54,372 taxable only.
  Tax comp: MP months now explained (MP Oct-25 -> Add to 4D1: reclaim 34.4L not in 4D1).
- **ITC Register 2025-26 RCM lines rebuilt at DOCUMENT level** (rcm_itc_doclevel.py): the 397 category lines (state x
  claim month x category, Document Number 'RCM') for May-25..Mar-26 claims replaced by 2,866 document lines from the RCM
  Register (Document Number, Posting Date, Invoice No. = Reference, Vendor Name, Vendor GSTIN, category, amounts = legs
  summed, claim month from the register's claim column; Reco Remarks carries the Input-GL same-document check). Apr-25
  claim lines (49, FY 24-25 RCM = 6A1 component) kept. Golden 1,06,98,71,702.15 -> 1,06,99,62,864.15 (+91,162 = the
  register-vs-CA-claim month gaps: UP Aug-25 +44,820, Bihar Dec-25 +1,250, Feb-26 +812, ...). Net-ITC check 0.00.
  TRAP (cost a restore): the register's RCM lines are NOT contiguous per state - delete ALL target rows first (exact rows,
  asserted), THEN insert per state at positions re-scanned after the deletes, bottom-up. Verify per-claim-month totals
  against the register before saving. cascade_fix.py will re-label RCM lines' Reco Remarks if re-run.
- **ZFI06 (BK-BM lookups):** the client's export (all four files = the same 1,717 documents) shares ZERO of the ITC
  register's 6,018 FY 25-26 RE documents (same 29-series range, same months/BPs) - it is a complementary selection, not a
  key problem. Only 29 other-series docs match (108 rows). Needs a fresh ZFI06 run without that selection.

## Changes log - 2026-09-18 (night 2): matching layers per Pawan's reco script, register date fix, Apr-25 RCM docs, 2B Return Period

- **ITC register dates carried the 18:30 tz shift** (41,111 invoice / 40,891 posting dates) -> itc_date_fix.py normalised
  them (same defect as the RCM register). Date+amount fallback works again (0 -> 205 lines).
- **Matching layers (Pawan 18-09, "not too aggressive", from his ITC vs 2B Reco.py):** (1) exact GSTIN + invoice (zero-
  insensitive); (2) GSTIN + amount; (3) invoice-similar + amount - 'similar' = core-normalised numbers (INV/INVOICE and
  trailing FY stripped) contain each other, or same prefix + numeric tail within edit distance 2. Layers 2-3 work at
  DOCUMENT level (register lines summed per vendor+invoice vs 2B document), same recipient GSTIN, tolerance +/-100,
  single candidate only, each 2B document once; verdict text ends '- invoice no differs, review'. Line-level amount
  fallback now single-candidate too. No PAN / cross-state layers. Result: not-found 8,172 -> 5,081 lines; GSTIN+amount
  364 docs (2,229 lines), invoice-similar 180 docs (1,273 lines); Consider 6,830; D_ 14.90cr -> 9.42cr.
  SHIV KIRAN case (register GZ/04/24-25 & 3/GZ/03 vs 2B GZ/04 & GZ/03) now matches; 2B docs read 'Claimed'.
- **Apr-25 RCM lines** (49 category lines) replaced by the 382 Mar-25 documents from last year's RCM Register
  (rcm_apr25_docs.py): supplier GSTIN on 154, vendor names, SAP doc numbers; tax 26,09,545 (Arunachal +6,678 vs the
  CA's claim). Golden 1,06,99,62,864.15 -> **1,06,99,69,542.15**.
- **TRAP (cost 40 min):** inserting rows AT row 6 shifts every range that starts at $6 (own sheet and dependents) to
  $6+K. Never insert at the first data row - insert at row 7 and move, or repair afterwards. Brute-force
  UsedRange.Replace across 71 sheets did not finish in 40 min; the targeted repair (openpyxl scan -> per-column bulk
  Formula writes, repair_range_targets.py) took 79 s. Register B_ columns are rewritten by final_countif_rule.py.
- **T6A1 Extract '2B Return Period'** (t6a1_2b_period.py): the old formulas pointed at a vanished helper column $AF
  (+68 row offset) - 3,141 blanks. Now a self-contained LIVE lookup: FY 24-25 2B ('GSTR-2B ITC Data', helper KEY col AB)
  first, then 'GSTR-2B Apr25-Aug26' (KEY AW -> Tax Period), else 'Not in 2B (Apr-24 to Aug-26)'; RCM rows fixed text.
  SUBTOTAL(9) totals in bold on row 3 for O:R. 1,482 dated / 1,773 not in 2B (register-sourced correction entries).
  cascade_fix.py rebuilds the extract rows -> re-run t6a1_2b_period.py after every cascade.
- ZFI06 unchanged (client export is a complementary document set). FY 26-27 register: Inputs sheets are the cumulative
  FY 25-26 register; Table 13 rows = Invoice Year 25-26 with GSTR 2B PERIOD in FY 26-27 (~3,791 lines) - proposal pending.

## Changes log - 2026-09-18 (night 3): ITC Register 2026-27 rebuilt from the client's Inputs sheets (Pawan)

- Source = Apr-Jun 26 state working files, sheet 'Inputs' (header row holds 'Document Number'; 17 header variants -
  read by name with fallbacks: Business place/Business Place, Vendor GSTIN/GSTN, Reference/Invoice no, Taxable Value/
  Taxable Amt/Taxable Amount, G/L Account/G-L Account, Invoice Year/Invoice year/FY/Document Year, GST CREDIT/GST CREDIT
  YES/NO/YES, GSTR 2B PERIOD/GSTR2B Month). The Inputs sheet is the CUMULATIVE FY 25-26 register carried forward, so
  'GST CREDIT = Yes' is NOT 'claimed in 26-27'. Table 13 / 12C rows = Invoice Year 25-26 AND GSTR 2B PERIOD in FY 26-27,
  de-duplicated across the monthly files on (doc, GL, vendor GSTIN, reference, IGST, CGST, SGST): 2,109 rows, tax
  2,42,80,458 (Apr-26 1.66cr / May 0.24cr / Jun 0.52cr); July files carry no FY 25-26 rows. Replaces the 260 2B-derived rows.
- Live 2B columns against 'GSTR-2B Apr25-Aug26' (KEY AW): Available in 2B Y 1,083 / N 1,026; Final Remarks 'Matched A'
  1,072, 'Matched - in 2B of 2025-26' 11, 'Not in 2B' 1,026. GSTR 9_Reporting '13', 9C '12C', Reasons text, 3B Claim
  Month as text '01 Apr 2026' (Excel auto-dates the string - write with NumberFormat '@'). ITC Summary Table 13 / 12C
  blocks follow via the expanded ranges ($5:$2113): T13 IGST 2,27,25,031.78 / C-SGST 1,74,12,286.58 each.
- Method: replace rows by inserting INSIDE the range (row 6) then deleting the old rows - dependents' ranges follow with
  no repair; single COM session incl. xlsb export (~2 min). fy2627_rebuild.py + fy2627_cols.py.

## Changes log - 2026-09-21: standard remark vocabulary (Pawan), Tax Rate formula, Table 13 tolerance

- **Remarks standardised** on ITC Register 2025-26 'Reco Remarks' and the 2B sheet's Reco Remarks / 6A1 mark / Table 8A
  (remarks_standardize.py; cascade_fix.py and b2_remerge_2627.py emit the same texts): sentence case, 'Matched with 2B – <basis>'
  ('– review' on non-exact), 'Not in 2B – Apr-25 to Aug-26', 'Not applicable – RCM self-invoice / – ISD / – no vendor GSTIN (URD)';
  2B side 'Matched with ITC Register – claimed Mmm-yy' / 'Not in ITC Register – FY 25-26 claims', 'Table 6A1 – FY 24-25 invoice …',
  Table 8A 'No – RCM / No – FY xx-xx document / No – 2B period FY xx-xx / No – ITC not available / No – amendment'.
  UNCHANGED on purpose (ITC Summary / T6A1 test them byte-exact): 'Table 6A1 of GSTR-9 - Unclaimed/Claimed', Table 13
  'Claimed/Unclaimed', 'Correction Entries- ITC dated 24-25 reversed in 25-26', Countif 'Consider/Not consider', T6A1 Source labels.
- ITC Register 'Tax Rate' is a live formula = Total GST / Taxable Value x 100 (7 rows differ >0.5 from the client's stated rate).
- 'Table 13 & 6A1 differences' Remarks: Matched when ABS(Total) < 10 (was < 1).

## Changes log - 2026-09-21 (batch 3): reco_lib, numbered mirrored remarks, Octa FY 24-25 2B base, FY 26-27 reco (Pawan + CA Priyesh recordings)

- **Matching is one tested module**: `vel/scripts/reco_lib.py` (`match_register`, `zkey`, `classify_vendor_gstin`, vocabulary `V`;
  12 tests in `tests/vel/test_reco_lib.py`). cascade_fix.py (FY 25-26) and fy2627_reco.py (FY 26-27) both call it. Layers:
  1 exact GSTIN + zero-insensitive invoice (recipient checked), 1b malformed GSTIN rescued by PAN + exact invoice, 2 GSTIN + document
  amount ±100 (same recipient, single candidate) / similar invoice + amount, 2b line date+amount / amount single candidate, 4 FY 24-25 2B.
  `zkey` strips leading zeros per digit run BEFORE removing separators (`SDIP/25-26/007` == `.../7`).
- **Numbered remark vocabulary (CA Priyesh, recording 2, 21-09)** - the SAME numbered text on the register and on the 2B sheet so
  a filter on "1 –" ties 1-1 on both sides; only `10 – Not in 2B – Apr-25 to Aug-26` (books only) and `11 – Not in books – FY 25-26
  claims` (2B only) differ. Register FY 25-26 (44,310 lines): 1 exact 32,075 | 2 recipient GSTIN differs 1 | 3 vendor GSTIN corrected
  from 2B 17 | 4 amount & date tie 2,714 | 5 similar invoice + amount 311 | 6 GSTIN + amount 457 | 7 date + amount 82 | 8 amount only 1 |
  9 FY 24-25 2B 1,164 | 10 not in 2B 4,948 | 12 RCM 2,527 | 14 URD 7 | 15 vendor GSTIN invalid 2 (MSEDCL, 16 chars).
  Ruling (b): when the document amounts tie to the rupee AND the dates agree, remark 4 carries NO "– review" (only the invoice number is
  written differently) - 2,714 lines left the review pile. Remarks are VALUES on the register (CAs edit them); the 2B sheet's
  `Reco Remarks` is a LIVE lookup of the register remark by KEY (25-26 register first, then 26-27, else remark 11) - remarks_mirror.py:
  2B side 13,867 docs = 11 not in books 7,239 | 1 5,958 | 4 454 | 6 93 | 7 85 | 5 27 | 2 5 | 3 5 | 8 1 (was 7,405 / 6,462 with the old
  claim-month formula; the extra 166 matches come from the FY 26-27 register).
- **Vendor GSTIN "URD" confusion fixed**: `classify_vendor_gstin` - blank / `0` / `NA` / `Missing` = no GSTIN (remark 14 or RCM/ISD 12/13);
  14- or 16-char values = malformed, never URD: rescued by PAN + exact invoice (remark 3, 17 lines) or `15 – Not matched – vendor GSTIN
  invalid (n chars) – review` (2 lines). Countif labels follow (`Not consider - vendor GSTIN invalid` 2, `... no vendor GSTIN` 7).
- **Recipient GSTIN mismatch is precise**: `2 – Matched with 2B – invoice no – recipient GSTIN differs (2B under <GSTIN>) – review`
  names the VEL GSTIN the 2B document sits under (1 register line; 5 2B docs on the mirrored side).
- **`GSTR-2B ITC Data` = the Octa PAN-level FY 24-25 export** (`Audit Data of FY 2024-25\PAN GSTR2B 2024-25.xlsx`, rebuild_2b_itc_data.py):
  10,322 documents + 188 ISD, CDN negative, per-row IGST+CGST+SGST == Total Tax Value gate; header row 5 and the helper KEY column kept
  (KEY = supplier GSTIN + invoice + "|" + FY). FY 24-25-matched register lines (remark 9) are `Consider` and take 2B_ from this sheet by
  that key (ruling (d)) - 2B_ 87.03 cr, D_ 6.72 cr on B_ 93,75,28,722.19 (== ITC-category tax).
- **Permanent reversals from last year's 9C** (perm_reversals_ly.py): 1,532 LY-flagged documents -> 2B Apr25-Aug26 col BG (5 rows) and
  `GSTR-2B ITC Data` new column `Permanent Reversals (LY 9C)` (1,549 rows); ITC Summary CE:CG C/SGST +236.74 each.
- **T6A1 Extract `2B Return Period` keyed on GSTIN + invoice + invoice FY** (Pawan's `VEL 21.9.26.docx`: supplier `09DCEPK6815A2ZS`
  invoice `3` exists in FY 23-24 AND FY 25-26 - a plain VLOOKUP picked Oct-23). Helper `KEY+FY` column on the CY 2B sheet (`=$AW3&"|"&$AV3`).
  Extract 3,395 rows: 1,354 dated (FY 24-25 2B 1,185 / Apr-25..Aug-26 169), 2,041 not in 2B.
- **ITC Register 2026-27 reconciled in the 25-26 format** (fy2627_reco.py; columns KEY … Reco Remarks APPENDED after the last header -
  ITC Summary reads this sheet by letter, never insert in the middle): 2,109 lines = 181 SAP documents (RA bills carry up to 97 GL
  lines) -> Countif Consider 181 / Not consider 1,928; verdicts 1 1,290 | 4 218 | 5 26 | 6 48 | 7 2 | 10 525; B_ 2,42,80,458.10 == ITC tax,
  2B_ 2,36,63,716.09, D_ 6,16,742.01. `2B_` reads `GSTR-2B Apr25-Aug26` V/W/X by KEY AW.
- Chain order now: b2_remerge_2627 → strip_orphan_pivots → readd_buttons2 → cascade_fix → final_countif_rule → t6a1_2b_period →
  fy2627_reco → remarks_mirror → rcm_gl_rebuild (vel/scripts/chain4.sh, ~11 min, one COM session per script). Verified after chain4:
  0 error cells, golden 1,06,99,69,542.15, Net-ITC 0.00, ITC Summary 6A1 blocks unchanged [2,67,48,108.62 / 1,45,33,807.83 /
  3,75,19,734.83 / 78,67,313.94 / 26,09,545.00], ITCR vs 3B net 2,52,024.32.
- Recordings 21-09 (0922 = 9 min, 0931 = 35 min) transcribed locally (faster-whisper medium, translate mode) to
  `Downloads\Meeting 21-09-2026 09xx - Transcript (English).md`; frames reviewed (RCM POS block, 2B Doc No filter, PY 2B_ zeros).
- Read-back after chain4 (document numbers only): remark 2 (recipient GSTIN differs) 5 lines / 4 docs - 2900003093 (2B under 09…),
  2900010470 / 2900011320 / 2900012213 (vendor 23…, 2B under 22…), 2900004027 (2B under 10…); remark 3 (GSTIN corrected from 2B)
  17 lines / 2 vendors - 2900003829 (8 lines, 14-char GSTIN), 2900005509 / 5003 / 5691 / 5698 (14-char); remark 15 (unresolved)
  2 lines - 3500021788 / 3500021787 (MSEDCL, 16 chars). Docx row (Bihar, supplier 09DCEPK6815A2ZS, invoice 3): `Not in 2B (Apr-24 to Aug-26)`.
- Known stray: `T6A1 Extract - 24-25`!V4 holds the constant text `#VALUE!` (pre-existing, not a formula error; SpecialCells reports 0).
- RCM GL Mar-25 restriction (ruling (a)) ran on 22-09 after the `\192.168.1.69` share came back (the 21-09 chain aborted on that
  step before opening the master): 0 error cells, Statewise RCM vs 3B unchanged +5,97,776 - details in the gst-audit-vel-rcm log 21-09.

## Changes log - 2026-09-22: '1 –' filter ties across the register and the 2B sheet (Pawan; CA Priyesh 21-09 31:13)

- Pawan filtered '1 –' on both sheets and the raw IGST subtotals differed (register 58,77,73,381 vs 2B 59,60,59,006). Raw register tax
  is per SAP line (several lines and duplicate bookings per invoice) and carries the genuine books-vs-2B differences (D_); the like-for-like
  tie is the register's 2B_ (Consider lines) against the 2B sheet's IGST/CGST/SGST (Net). Three causes fixed (tie_fix.py, one COM session):
  1. 125 2B rows were matched via ITC Register 2026-27 - the 2B sheet's remark lookup now suffixes those ' (ITCR 26-27)' (remarks_mirror.py).
  2. 616 RCM-category register lines (registered GTA vendors) sat in 2B and carried 1/2/7/9 - they are outside the Countif/B_/2B_ frame,
     so they now read `12 – Not applicable – RCM line (in 2B: <basis>)` on both sheets (reco_lib RCM_BASIS; 489 'invoice no', 127 'FY 24-25 2B').
  3. Two invoices booked twice in SAP (2900009333 / 6200000040 as '0496' and '496'; 2600000588 / 6200000032) pulled the 2B amount twice:
     new register column `2B pull` (BV, values: Yes on the first Consider line per KEY2, No on a later one); 2B_ formulas pull 0 on 'No'
     so D_ = B_ surfaces the duplicate claim (10.4L IGST). final_countif_rule.py stamps the column on every run.
  Result: register 2B_ on '1 –' Consider = 2B sheet (Net) on '1 –' = IGST 58,16,03,159.31 / CGST 9,66,92,174.47 / SGST 9,66,92,174.47;
  B_ on the same lines 58,22,17,579.16 / 9,75,45,864.79 / 9,75,45,864.79 (the D_ to explain). 0 error cells, golden and ITC Summary unchanged.
- **Found, NOT yet applied (needs Pawan/CA):** cascade_fix tagged 127 RCM Mar-25 lines (tax 12,16,544) as Table 6A1 component 1
  ('ITC dated 24-25 in 2B of 24-25 availed in 25-26') because their remark was 9; the same documents are the RCM component
  ('RCM paid in Mar-25 availed in Apr-25', 382 rows, 26,09,545) - a double count inside 6A1. cascade_fix.py now requires category ITC for
  component 1; the next cascade run will drop those 127 lines from component 1 (ITC Summary block G falls by ~12.2L). Not re-run yet.
- The 2B-side '1 –' count is 5,348 rows + 125 '(ITCR 26-27)'; register 12 now 3,143 (2,527 self-invoice + 616 in 2B), 9 now 1,037.

## Changes log - 2026-09-22: ITC Register 2026-27 rebuilt in the exact ITC Register 2025-26 layout (Pawan: "use the SAME FORMAT entirely")

- fy2627_relayout.py rewrote the sheet IN PLACE: rows 2-3 title/description, row 4 SUBTOTALs (R..V), header row 5, data rows 6..2114;
  columns A..BV are the 25-26 headers in the 25-26 order (names, header styling, widths, number formats), formulas re-based to this
  sheet's rows (Tax Rate, Total GST, KEY, B_/2B_/D_ incl. the PY branch and `2B pull`, the four 'as per 2B' lookups, POS block, ZFI06
  lookups, 2B Year). Values mapped by name (STATE NAME -> State Name, VEL GSTN -> VEL GSTIN, Type -> TYPE, Consider in 8A reco ->
  Consider 8A reco, Posting Year -> F.Y/Booking Year); '3B Claim Month' text ('01 Apr 2026' / '03 June 2026') became real dates;
  'Matching of 12B…' = NA; columns with no 26-27 source (Nature of Services, Type for GSTR9, Material Description, Eligibility,
  Considered in Table 6A1, Remarks for accounting entries) blank. The sheet's own columns follow after BV: Correct GSTIN,
  GSTR 9_Reporting, Reasons (GSTR 9), Available in 2B, 2B Inv, 2B Period (client lookup), Final Remarks, Query, POS Remarks, Source.
- External references repointed by header name with rows +1: ITC Summary (114 cells; T/U/V -> S/T/U, X -> BX, Z -> W, D -> D),
  GSTR-2B Apr25-Aug26 remark lookup (BE/BN -> AF/AO), INDEX hyperlink. Verified: ITC Summary values identical before/after, 0 error
  cells, 25-26 golden unchanged, 26-27 Total GST = B_ = 2,42,80,458.10, 2B_ 2,36,63,716.09, D_ 6,16,742.01, Countif 181/1,928,
  remark counts unchanged, 2B sheet remark counts unchanged. xlsb exported.
- fy2627_reco.py / remarks_mirror.py now read header row 5 / data row 6 (and 'VEL GSTIN'). TRAP: fy2627_rebuild.py + fy2627_cols.py
  still produce the OLD layout (header row 4) - after any data rebuild run fy2627_relayout.py again, then fy2627_reco.py.

- 22-09: rows 6-19 of ITC Register 2025-26 (the Apr-25 RCM docs inserted at row 6) carried the dark header fill - reset to the normal row
  format (unfill_rows.py copies row 20 formats; values untouched). TRAP: rows inserted at the first data row inherit the header format.

## Changes log - 2026-09-22/23: remarks on the Consider line only; dd-mm-yy everywhere (Pawan)

- **Reco Remarks / KEY2 only on the Consider line of each ITC document** (Pawan 22-09: "you shouldn't put remarks for not-considered
  items, SUMIFS already brings them under Consider"). cascade_fix.py / fy2627_reco.py `consolidate()`: Countif grouping (vendor GSTIN
  or NM:name + normalised invoice), the Consider line takes the BEST match found on any line of the document (lowest remark number,
  prior-year keys count), every other line -> blank remark + blank KEY2. RCM/ISD lines untouched (each is its own document). The
  per-line verdicts still drive the Table 6A1 tagging and the extract internally. 25-26: 6,830 documents, 34,061 lines blanked,
  2 KG/RE pairs (2600000434/2900006365, 2600000635/2900009426: credit memo first in row order = Consider line, the RE line had
  matched by date+amount) now read `7 –` on the Consider line with KEY2, 2B_ pulled, D_ -422 / -512 instead of the whole B_.
  26-27: 181 documents, 1,928 lines blanked. 2B mirror counts unchanged (it found the Consider line already).
- **Table 6A1 double count NOT applied** (Pawan 23-09: "don't fix T6A1 right now"): cascade_fix.py keeps the RCM Mar-25 lines in
  component 1 (they now read `12 – … (in 2B: FY 24-25 2B)` so the test is `"Matched with 2B of FY 24-25" in v or "(in 2B: FY 24-25 2B)" in v`);
  the ITC-only rule sits behind env `SIX_ITC_ONLY=1`. Block G stays 2,67,48,108.62. TRAP (cost a restore): a vocabulary change can
  silently move Table 6A1 - always compare the 6A1 blocks printed by cascade_fix with the previous run.
- **All full-date columns display dd-mm-yy** (dates_ddmmyy.py, 63 columns on 20 sheets; month-only 'Apr-25' columns untouched).
  Trigger: T6A1 Extract 'Invoice Date' was mm-dd-yy, so 6 Mar 2025 read as 03-06-25. Values/formulas unchanged.
- Runner after a cascade when the master gets locked mid-chain: chain5b.sh = t6a1_2b_period -> fy2627_reco -> remarks_mirror ->
  dates_ddmmyy. TRAP: TaskStop/kill of the chain runner does not stop the running python/COM step - wait for it, then restore.

- 23-09 TRAP (cost a revert): the dd-mm-yy pass was driven by the format survey, which flagged any column holding a few date-formatted
  cells - Invoice No. (register K, 2B base I, extract M), RCM Reference/Inv. No. and three ITC Summary amount columns got a date format;
  numeric invoice numbers then read as #VALUE! in openpyxl and showed as dates/#### in Excel. Restored from the snapshot
  (dates_revert_nondate.py); dates_ddmmyy.py now lists pure date columns only. Values/formulas were never affected (0 error cells).
