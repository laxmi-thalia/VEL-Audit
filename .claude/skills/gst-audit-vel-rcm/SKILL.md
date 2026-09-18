---
name: gst-audit-vel-rcm
description: VEL (Vikran Engineering) GSTR-9/9C annual-return audit — RCM phase. The complete recipe for the RCM Register and its seven checklist steps inside VEL_GST_Audit_FY2025-26_MASTER.xlsx - sources, matching keys, checks, findings, traps, missing-data flags, and post-ITC hooks. Use for any work on the VEL RCM sheets or the FY 26-27 RCM roll-forward. Sales phase: gst-audit-vel-sales. ITC phase: gst-audit-vel-itc (built).
---

# VEL GST audit — RCM phase

Same engagement/master workbook as `gst-audit-vel-sales` — **read that skill's Standing
rules first; every one applies here** (formulas-only, no LLM without asking, file-lock
etiquette, verify-before-reporting, user edits win, re-snapshot before edits).

Engine home: `vel/` (this repo: C:/PROJECTS/gst-audit-engine)
- `contracts/rcm-sources.md` — **read first**: every RCM source file with its structure,
  quirks and traps (phantom rows, doc-number recycling, state-name typos, account markers).
- `contracts/goldens_fy2526_rcm.md` — the backtest targets (gitignored; client figures
  never live in this committable skill).
- `contracts/formats.md` + `reference/formats.json` — per-sheet layouts of the WHOLE
  master (regenerated after RCM; the authority for column letters).
- `scripts/` — rcm0_build/rcm0_paste (register), b3d_extract + rcm2_build (vs 3B),
  rcm3_build (ToS/interest), rcm4_build (GL), rcm5_build (rate-wise), b2rcm_extract +
  rcm6_build (vs 2B), rcm7_build (TB scrutiny), plus the shared readd_buttons.py.
- Meeting that defined this phase: `reference/transcripts/rashid-meeting-transcript.md`
  (Rashid, 27-Aug-2026 — RCM/ITC methodology walkthrough).

## The RCM design in one paragraph

VEL pays RCM **output in the 3B month itself but claims its ITC one month later** — so the
March→April spillover crosses FYs and the input side of everything waits for the ITC
register's claim months. The register is GL-posting grain (one row per CGST/SGST/IGST RCM
output posting), poured from the client's Conso RCM into **last year's 72-column final
layout** (NOT Rashid's 58-col draft — the final adds the portal block, Final 3B Month, Key).
All checks then hang off the register: live rate/POS formulas on the sheet, SUMIFS recos
to 3B 3.1(d), stamped code-side matching to GL and 2B, TB scrutiny by keyword rules.

## Steps (Audit Checklist, Type=RCM, Steps column) and what each built

0. **Register** — `RCM Register` sheet (after S9). Conso 3,181 rows → 72 cols; live Key /
   Total GST / Rate Check / POS block; stamped GSTIN (BP map), BP+CC per Profit Centre
   (Cost Centres Sheet4), fiscal years. Subtotal row 4, headers row 5, data row 6+.
1. **Rate + POS** — live on the register (Rate Check shows effective %, POS = state-equality
   vs IGST-zero test). Exists the moment the register lands.
2. **vs 3B (liability)** — `Statewise RCM vs 3B` + `Month wise RCM vs 3B` in last year's
   three-block layout (RCM reg | GSTR-3B | Difference + DPS Remarks), live SUMIFS; 3B side
   = `RCM 3.1(d)` columns K-N added to the master's 3B Data (b3d_extract). GSTIN-keyed via
   hidden col S. Include a **HOIS/unmapped row WITH diff cells** (its omission broke the
   grand-total tie once). **ITC leg deferred post-ITC.**
3. **ToS + interest** — 8 live columns appended to the register (doc+61 / posting+61 /
   earlier / due 20th-after-ToS-month / paid-via-3B due / delay days / interest 18% / status)
   + `RCM ToS & Interest` per-state summary. Law: Sec 13(3) (61st day), Sec 50 (18%).
   Payment dates don't exist in books — due dates proxy to the 3B cycle 20th; the caveat is
   printed ON the sheet for DPS confirmation.
4. **GL Output & Input vs register** — `RCM GL` sheet (36,622 rows, both dumps side-tagged).
   **SEPARATE Output and Input match columns on BOTH sides** (Rashid's explicit ruling):
   GL sheet has `Matched with RCM Register (Output)` filled + `(Input)` reserved; register's
   old Found-in-GL pair renamed `Found in Output GL`/`Output GL Remarks` + new
   `Found in Input GL (post-ITC)`/`Input GL Remarks` appended. Match key
   **(account, doc no, doc-date FY)** — see trap list. Transfer entries knocked via
   Clearing Document.
5. **Rate-wise summary** — `Rate-wise summary-RCM`, last year's format, live SUMIFS by
   (GSTIN, rate±0.01), incl. HOIS row and a blank-rate row; GT must tie register subtotals.
6. **vs 2B** — `RCM vs 2B` sheet: RCM=Yes line items from the portal 2B files with match
   status; register gains `Found in 2B`. Three-way classification: matched / registered
   vendor not in 2B / **URD-self-invoice (not in 2B BY DESIGN — never call it a miss)**.
   Match cascade: (supplier GSTIN, normalized inv no) → (supplier GSTIN, rounded |taxable|).
7. **TB scrutiny** — `TB Scrutiny-RCM`: expense accounts (^4\d{9}) aggregated from the TB,
   keyword rules with word boundaries (`\brent\b` — "current" trap), flagged heads
   cross-checked against register offsetting accounts; ambers highlighted; every flag
   labelled AI-proposed for CA confirmation.

INDEX: every RCM sheet gets an `Arena = RCM` row in last year's index format.

## CA/Rashid rulings (append new ones here — this is the changes log)

- Expenditure GL = lookup from **Groupings sheet (col A account → col B name)** of the FS
  file, by Offsetting Account; embedded hidden `TB Groupings`, live INDEX/MATCH.
- GL Correct = SEMANTIC match Expenditure GL ↔ Nature of Services (e.g. "Guest House Rent"
  ↔ "Rent on Residential Property") — **deferred post-ITC**, judgment column.
- ITC Eligibility — **deferred post-ITC**; last year used Y/N by nature of service
  (precedent map mineable from last year's register when the time comes).
- Input/Output GL matching = separate columns both sides (never one combined column).
- 3B RCM figures live ON the 3B Data sheet (same master, not a separate workbook).
- Register in the master as a full sheet (size fear was phantom rows; real is 3,181).

## Traps (each cost a round — do not relearn)

- **Phantom rows**: conso + last-year sheets show ~322k used-range rows; real data 3,181.
  Filter Document Number notna before any row count or size decision.
- **SAP doc numbers recycle per fiscal year** — any GL/doc matching without doc-date-FY in
  the key silently merges different years' documents (710 false mismatches → 4 real).
- **Conso state names have typos** (Madya Pradesh, Tamilnadu) — GSTIN/BP keys only.
- **Rate float dirt** (17.999999999999996) — tolerance bands in SUMIFS criteria.
- **FBL3N exports**: preamble rows, header row 6, account markers as TRAILER rows at block
  end ("Account NNNN..."), no posting-date column (Assignment = yyyymmdd proxy), junk
  currency columns.
- **Octa omits nil rows** (Kerala 0/4 3.1.D rows) — default zeros, not errors.
- **TB lives INSIDE the FS-with-Notes file** ("Trial Balance" + "Groupings" sheets), not as
  a standalone file. Groupings header row 4.
- Last year's workbook contains a FOREIGN sheet ("RCM ITC" titled TP EV Charging — another
  client) — never replicate it; classify golden defects before reproducing.
- Register G/L Account values can carry suffixes ("2610080300-01") — normalize to digits.
- A summary row added without its Difference cells breaks grand-total ties (HOIS lesson).

## Missing data / open flags (as of 2026-08-28 — keep current)

- **Haryana GSTR-3B file missing** from Portal Reports → its 3B RCM = zeros.
- **GSTR-2B files missing for Haryana, West Bengal, Tamil Nadu, Telangana**; Kerala 2B nil.
- **2 register rows Business place = HOIS** (ISD) — no GSTIN; ruling needed (they also
  drive the +5.5L HOIS row in the statewise reco).
- **28 register rows with blank Document Type**.
- **4 amount mismatches** GL vs register — register exactly 2x GL each (likely conso
  duplicate rows): HOIS 2900007860, MP 2900009159/2900009490/2700018168. Client query.
- **2 GL output docs NOT in register**: JK doc 2900002128 (CGST+SGST 1,350 each, 01.03.2026).
- **87 2B RCM items NOT in register** (potential unrecorded liability) and **118 registered-
  vendor register rows not in 2B** — CA review lists on the sheets.
- **10 TB amber heads** flagged RCM-indicative but absent from the register (Director
  Remuneration split, Transport Hire, Freight Purchase, guesthouse R&M, licences,
  sponsorship...) — figures in goldens file.
- **304 register rows have no usable dates** — ToS untestable.
- **ZFI purchase report not received** — import-of-services check (checklist NA row) open.
- **3,047 rows offset to SR/IR Clearing** — expense-head verification needs clearing-doc
  hop; only ~134 rows have direct expense GLs.
- Findings pending CA remarks: 3B RCM exceeds register by 1.04cr (state split in goldens);
  64 late-payment rows (interest 20,534.64).

## Post-ITC hooks — FILLED (2026-08-28, see rcm_postitc.py + 'RCM Paid vs ITC Claimed' sheet)

1. Register: `GSTR 3B Claim month` (from ITC register), `ITC Eligibility`, `GL Correct`.
2. `RCM GL`: input-side matching (`Matched with RCM Register (Input)` + remarks) and the
   register's `Found in Input GL (post-ITC)`.
3. Step 2's ITC leg: extract 4A(2)/4A(3) from the Octa 3B files (clone b3d_extract with
   section 4.A) and reconcile claim months (input is one month behind output).

## Year-roll & maintenance

Same procedure as the sales skill (template-copy → survey vs contracts → rulings → rebuild
→ **backtest against goldens_fy2526_rcm.md to the rupee** → gauntlet). After every session:
append rulings to the changes log above, update the flags inventory, regenerate formats.md,
sync scripts. If a change alters numbers, it must reproduce the goldens first.

## Changes log - 2026-09-17 (Pawan's review round)

- `Final 3B Month` (B) is now LIVE `=DATE(YEAR(A),MONTH(A),1)`. Root cause of the 304 blanks:
  source text "03 June 2025"/"04 July 2025" (full month names) failed the `%b` parse; the tz
  shift (18:30 previous day) had put every April row at 31.03.2025 -> `Paid via 3B due date`
  a month early. Interest total moved 20,534.64 -> 22,568.94 (67 late rows). New golden.
- `POS Check` (BG) = `=$BF=$BE` -> full TRUE/FALSE on every row (853 TRUE / 2,328 FALSE; the
  FALSEs are mostly unregistered vendors: blank GSTN -> "Inter State" by state code while
  IGST=0 -> "Intra State" by amount - CA judgement, not a build defect).
- `Month wise RCM vs 3B` = State | GSTIN | Month grid (S3 Month-on-Month style, autofilter,
  standard reco sequence, DPS Remarks amber, hidden helper T = month-start date keyed on
  Final 3B Month) + a 12-row HOIS/unmapped block so the grand total ties the register
  (115,152,683.11) and the Statewise sheet (-10,448,701.69).
- **FINDING (source gap): July-25 RCM is missing from the client's `Conso RCM FY 25-26.xlsx`**
  - only 61 rows for 3B month Jul-25 / SAP 2025/04 (other months ~250-390). 3B 3.1(d) reports
  Rs 1,11,58,479 for July across 15 states (Arunachal 52.04L alone) with nothing in the
  register. This is ~107% of the total Statewise difference; the rest: MP Jan +1.12L, Gujarat
  Apr/Jun/Aug net +0.95L, TG Jan -0.54L, TN Jul/Oct timing 7,500. Ask the client for the July
  RCM postings.
- `TB Scrutiny-RCM` = the FULL trial balance (583 accounts aggregated across PC/period, sums to
  0), keyword indicator, LIVE `In RCM register?` COUNTIF over Offsetting Account, conditional
  shading for hit-but-not-in-register. Script `rcm_fixes.py`.

## Changes log - 2026-09-17 (later): July-25 gap closed from monthly workings + FY 26-27 3B

- **RULING (Pawan): the monthly working files are the fallback source for register rows the
  Conso RCM lacks.** Source = `Clients Data.08.2026 Main Data\<NN Month YYYY>\<State>\GSTR-3B <State>
  <Mon YYYY>.xlsx`, sheet **`RCM`** (header row = the row whose A = "Business place"; layout
  DRIFTS month to month - April 35 cols incl State, May+ drops State, counts 34-38 - map by
  header name; the G/L is written `2610080300-301` = CGST+SGST on ONE row -> split into two
  posting rows, taxable halved, so state-month SUMIFS tie). Survey script
  `rcm_monthly_survey.py` (pickle `rcm_monthly.pkl`); insert script `rcm_add_july.py`.
  Coverage of the monthly files: MP has NO monthly GSTR-3B file any month; Kerala/Haryana/
  Punjab(<=Sep) RCM sheets have a different layout (no Business place header); ISD none.
  TRAP: UP's July rows WERE in Conso (the 61 "July" rows) - inserting July for all 15 states
  duplicated UP; only insert states whose register month is empty.
- Register now 3,470 rows (`Source` column BZ tags Conso vs Monthly). NEW GOLDENS: SAP taxable
  124,326,649.78; Total GST 20,294,387.96; interest 23,369.02 (68 late rows); Statewise diff
  -1,274,735.02 (= MP July -19.85L still missing from every client file + Gujarat +0.95L +
  Telangana Jan -0.54L + TN 7,500 timing + HOIS 5.5L unmapped).
- **TRAP found: `3B Month` (col A) is a DATE in the Conso rows (Excel coerced "04 July 2025"
  to 4-Jul-2025 etc.)** while `RCM Paid vs ITC Claimed` compared it to TEXT -> that sheet's
  "RCM paid" column had read ZERO for every month since it was built. Re-keyed on `Final 3B
  Month` (B) via hidden helper col M (month start). Now paid 20,195,387.96 vs next-month
  4A(3) 20,448,930.00; only 6 state-months differ > Rs 1,000 (UP/Chh Jul, MP Jul & Jan,
  Bihar Nov, WB Dec).
- `3B Data FY26-27` sheet (Octa annual reports Apr-Jul 2026, 18 states, Haryana missing) added
  after `3B Data`; Mar-26 rows on `RCM Paid vs ITC Claimed` read Apr-26 4A(3) live from it -
  all 19 states tie to the rupee. Ruling: 3B only from `Audit data of FY 2026-27\GSTR-3B`, NOT 2B.
- INDEX buttons: every sheet's row 1 forced to 21pt before placing the shape (rebuilt sheets
  had 14.4pt -> button spilled into row 2); hidden CC Master / TB Groupings got a reserved
  row 1 too (30k dependent lookups shifted).

## Changes log - 2026-09-18 (night): RCM GL rebuilt on Posting Date | Document Number (Pawan)

- New GL folder `Clients Data/GLs/RCM`: 4 FBL3N dumps (sheet Data, header row 6, G/L Account per row): Output FY 25-26
  (8,620), Output FY 24-25 (7,460), Output open items Apr-Jun 26 (971), Input 1.4.25-31.7.26 (17,038). 1,239 of 7,231
  document numbers recur across fiscal years -> **GL Key = TEXT(Posting Date,"yyyymmdd")&"|"&Document Number** on both
  RCM GL and RCM Register (rcm_gl_rebuild.py, COM only). Posting FY from Posting Date (102 docs dated 24-25 post in 25-26).
- **Register date defect found & fixed:** the 3,142 Conso-sourced rows carried Posting/Document/Inv. Date as 18:30 of the
  previous day (IST-midnight-in-UTC from the original build); normalised to midnight (2 posting dates changed month).
  Raw key match was 361/3,503; after the fix 3,502/3,503 (Found in Output GL) and 3,501 have the RCM input debit in the
  SAME SAP document (Found in Input GL (GL Key)).
- **Register conventions (verified):** Conso rows carry ONE line per document for CGST/SGST, labelled 2610080300 or
  2610080300-01, amount = the CGST leg (SGST equal, implied); monthly-working rows carry separate 300/301 lines.
  GL-side expected amount: leg 300 -> lines 300 + 300-01; 301 -> line 301 if present else mirror of the CGST leg;
  302 -> line 302. Register-side compares the GL leg with the SUM of its own lines of the same key + label.
- Result: 0 error cells; register amount-differs = the 4 doubled Conso documents (e.g. GL -49,500 vs register -99,000)
  + 3 others; GL side: 634 FY 25-26 output credit lines NOT IN RCM REGISTER (client gap, ties to the Statewise -12.75L /
  MP July), 2,065 debit (payment) lines, 7,440 FY 24-25, 971 FY 26-27 open items. Statewise RCM vs 3B unchanged.
- TRAP: delete/recreate the GL sheet BEFORE writing register formulas that reference it (a delete turns them into #REF!).
  'Found in Input GL (post-ITC)' / 'Input GL Remarks' keep their claim-month semantics - GL-key columns were added.
