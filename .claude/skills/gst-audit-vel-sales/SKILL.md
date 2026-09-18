---
name: gst-audit-vel-sales
description: VEL (Vikran Engineering) GSTR-9/9C annual-return audit — SALES phase. The complete recipe that built VEL_GST_Audit_FY2025-26_MASTER.xlsx - sources, register build, matching, every reco sheet, formats, verification gauntlet, traps, and the year-roll procedure. Use for any work on the VEL sales master, its step files, or the FY 26-27 sales roll-forward. RCM phase: gst-audit-vel-rcm (built). ITC: gst-audit-vel-itc (planned).
---

# VEL GST audit — SALES phase

Engagement: Vikran Engineering Limited (VEL), 19 state GSTINs, books in SAP, returns via
ClearTax, portal extracts via Octa. CA = Priyesh (DPS & Co). Operator = Pawan (developer, not
a CA). FY 2025-26 built and verified; this skill captures exactly how, so FY 26-27 is a
one-line prompt.

Engine home: `c:\PROJECTS\accountic\gst-audit-engine\vel\`
- `scripts/` — every build/fix/verify script that made the master (ported from the session
  scratchpad; load-bearing ones named below).
- `contracts/sources.md` — **read this first**: source paths, header contracts, closed
  vocabularies, the MOB sign rule, the 19 GSTINs. Format drift protocol lives there.
- `contracts/formats.md` — **the per-sheet format spec, generated from the verified
  master**: every sheet's header row, exact column list with letters, freeze panes, filter
  ranges, title rows, header styling. `reference/formats.json` is the machine-exact twin.
  **This is the authority for column letters** — the user edits the file directly (has
  added his own columns), so NEVER trust remembered letters; regenerate the spec
  (`scripts/` has the dump code inline in the session log) or resolve headers BY NAME at
  runtime, then update formats.md.
- `reference/` — style dumps of last year's formats (salesreco_dump.json, index_dump.json)
  + cn_fill_log.json (unparseable/conflicting CN refs).
- `reference/transcripts/` — the RAW instruction sources: the three translated CA-session
  transcripts, the Priyesh CC/PC transcript + briefs. The decisions log below summarizes
  them; when a summary seems off, the transcript is the authority.
- Related repo doc: `docs/gst-audit/VEL_MONTHLY_GST_PROCESS.md` (the monthly GST filing
  process for VEL — adjacent process, same client/sources).

Related general skill: `gst-audit` (Amar engagements; shared doctrine — backtest-first,
no plugs, formula-driven, template-first). The Rashid rules there apply here with Priyesh
as the electing CA.

## Standing rules (Pawan's + CA's — non-negotiable)

1. **NEVER hardcode values.** Every derived cell is a live formula; the only literals allowed
   are (a) raw source data on data sheets, (b) externally-audited inputs (prior-year openings)
   with their source labelled. Proof pattern: scan B7:U42-style figure areas for numeric
   constants = 0. A formula containing a numeric literal (last year's `=65090095-M9`) is
   still a hardcode — sweep formulas for 7+ digit literals.
2. **No LLM/AI call without asking Pawan first.** Everything here is deterministic code.
3. **Client data stays out of the chat** — aggregates, counts, doc numbers ok; never dump
   party names/rows. Files are scanned in batches with reports between (stepwise).
4. **File-lock etiquette:** Pawan's Excel holds the master open most of the day. Check
   `open(path,'r+b')` before every write; on PermissionError, ASK to close — never write a
   twin copy. He must close (not save) before builds; tell him when it's safe to reopen.
5. **Verify before reporting.** Full COM recalc → zero error cells → golden checks. Never
   claim success from reading the diff. The user's stock question is "did you verify?" —
   the answer must already be yes, with output.
6. **He edits the file directly too** (renames/deletes sheets, saves). Re-snapshot before
   every operation; never assume yesterday's sheet list. His edits win; adjust INDEX and
   dependents to match (e.g. he renamed 'S2 Reco (CA format)'→'S2 SR vs GSTR-1' and deleted
   Sample Selection / SAP Enrichment / Load Log / Format Legend / S6 GL Advance Check).

## The deliverable — master workbook anatomy

`C:\Users\pawar\Downloads\VEL_GST_Audit_FY2025-26_MASTER.xlsx` — ONE workbook, ONE source:
every books-side figure everywhere is a SUMIFS over `SR_2025-26`. External sources live as
data sheets (GSTR-1 Data / 3B Data / GL Data / FS Revenue Data), each carrying VICE-VERSA
match columns so each source polices itself.

Sheet inventory (post user-edits, 28 sheets): INDEX; Open Points; SR_2025-26;
Queries (Draft); Step 1 Flags; the 4 data sheets; S2 Month-on-Month / S2 SR vs GSTR-1 /
S2 Pivot Month-on-Month / S2 Exceptions; S3 Month-on-Month / S3 1 vs 3B / S3 SR vs 3B /
S3 SR vs 3B MoM / S3 Amendment Check; S4 GL vs SR / S4 Exceptions; S5 HSN Summary /
S5 Rate-wise; S6 Advances Control / S6 Month-on-Month; S7 CN Time-bar; S9 Sales Reco;
Open Points (all); CC Master (hidden).

**Global format decisions (CA/Pawan-elected):**
- **Row 1 of every sheet is RESERVED** for the floating `<< INDEX` back-button (COM shape,
  hyperlink to INDEX!A1, parked at 2,2). Headers start row 2+. Defined name INDEX_HOME
  allows F5-jump. Register headers are therefore ROW 5, data rows 6..27007, SUBTOTAL row 3.
- **INDEX sheet in LAST YEAR's format** (from the FY 24-25 xlsb 'Index'): dark 333F4F header
  `Arena|Descriptions|State Code|GSTN|GSTR-9|GSTR-9C|Particulars|GSTR 9|GSTR 9C|Remarks`,
  Arena="Sales" rows only (ITC/RCM rows join in their phases), cell-hyperlink in Particulars,
  bordered rows, widths 11.89/42.22/9.33/17/8.11/8.11/25/11.56/11.56/22.67. Only sheets that
  EXIST get rows.
- **Standard reco column sequence** (Pawan-elected, applies to every reco):
  `[side] Taxable | IGST | CGST | SGST | Total Tax` in order books/SR → comparator → Diff,
  then auto-explain / residual / "Remark (type here)" / hidden key. S4 adapts to tax-heads
  only (GL has no taxable). S2 SR vs GSTR-1 keeps the CA's vertical-block grid (exempt).
- **Remarks live IN the reco sheets** (CA query 2): typed per state-month in MoM sheets;
  totality sheets auto-concatenate via 12-term `TRIM(IFERROR(IF(LEN(INDEX(...MATCH(
  gstin&"|"&month, keycol)))...)))`.
- **Pivot** `PT_Long` on S2 Month-on-Month A?:X?: rows State/Month/GSTR-1 Type (tabular,
  subtotals off, RepeatAllLabels ON), 15 sum fields in the standard sequence,
  PivotStyleLight16. Real PivotTable = CA's double-click drill-down. Needs manual Refresh
  after data edits — it snapshots.
- **S9 Sales Reco is an exact style replica of last year's 'Sales Reco'** (42r x 22c; dark
  333F4F header, D6DCE4 section bands, 8497B0 total rows, accounting `_(* #,##0_)...` format,
  9C refs col V: 5A/5Q/5O/5H/5I/5C/5D; two reason sections "In Financials not in GST" (CA
  inputs, negative) and "In GST not in Financials" (advance opening −, closing +, live from
  S6). Rebuild method: `scripts/rebuild_s9.py` + `reference/salesreco_dump.json` — copy
  styles/labels/internal formulas, swap data cells for live formulas, BLANK judgment cells.

## Build order (what actually happened, as the replayable pipeline)

0. **Survey** every source vs `contracts/sources.md`; STOP on drift (protocol in contract).
1. **Register** (`build_register.py`, `sr_compare2.py`): 12 ClearTax monthlies → SR sheet in
   the CA's frozen format (his column layout preserved; Format Legend recorded it). Header-
   name mapping; Working sheets only; my-GSTIN keying (NOT state names — J&K trap).
2. **SAP enrichment** (`enrich.py`): IRN→ODN join adds GL Name (AL), Profit Centre (AM),
   Cost Centre code (AO), Business Place, SAP Matched By. Later: **CC Name (AN)** inserted
   BESIDE Profit Centre via COM column-insert (Priyesh: Sheet4 A→F narration = place-of-
   supply check; hidden CC Master sheet; `cc_embed.py` + `cc_insert.py`).
   (Column LETTERS below were true at build time — the user has since added his own columns;
   resolve by HEADER NAME per contracts/formats.md, never by remembered letter.)
3. **Register add-on columns**: Sample Invoices/POs (document REFERENCES, not flags;
   PO = SAP Contract No; selection basis: materiality/sampling/unusual/risk —
   `step1_samples_v3.py`); Data Flags (BA, live formula); live flags gate sheet (Step 1
   Flags: blocking/review counts + running-count-helper+MATCH detail list — AGGREGATE-in-
   IFERROR does NOT work); Original Invoice Number/Date (BF/BG) from the **Credit Note
   Statement** (`fill_cn_originals.py` — parse col Y `<doc> Dated dd.mm.yyyy` keyed on Bill
   No; 239/288 CN docs filled); Adv Bucket helper (BH: MOB sign rule).
4. **Matching** (code-side, stamped): SR↔GSTR-1 cascade IRN → GSTIN+DocNo → Customer+Amount,
   each register doc consumable ONCE; nil-CN duplicates labelled, 34 docs "Not in books";
   SR↔GL on (Business Place state, Reference=invoice no), sales-origin doc types only.
   Stamps land on the register ("Matched with GSTR-1"/"Matched with GL") AND the data
   sheets (vice-versa columns).
5. **Recos** (all SUMIFS over the register; scripts master_build_1/2/3 + later sheets):
   S2 (CA grid + MoM + pivot + exceptions), S3 (1vs3B, SRvs3B, both MoMs, amendment check),
   S4 (GL grid + exceptions; Sales-origin="Y"!), S5 (HSN/rate, masters check), S6 (control
   account + MoM with rolling opening), S7 (CN time-bar LIVE off BF/BG), S9 (replica).
6. **Navigation**: INDEX (last-year format, `index_lastyear.py`) + reserved row-1 buttons
   (`button_row.py` once; `readd_buttons.py` after every openpyxl save — openpyxl DROPS
   shapes; pivots survive).
7. **Gauntlet** (`verify_shift.py` = current golden set): full COM recalc, zero error cells,
   ~18 goldens (see below), buttons present, pivot alive.

## Golden values

The backtest target (every FY 25-26 figure the gauntlet must reproduce to the rupee) lives
with the engine, NOT here: `gst-audit-engine/vel/contracts/goldens_fy2526.md` (prose) and
`scripts/verify_shift.py` (executable, with cell addresses). Client figures stay out of
this committable skill file by design. On a year-roll, the new year's verified figures
become `goldens_fy2627.md` — goldens are per-FY test expectations, never workbook content.

## Decisions log (the CA's 13 queries → what was built)

1 PO source → SAP Contract No via ODN (provenance column). 2 remarks in-place. 3 GSTR-1
vice-versa cols. 4 GL vice-versa col. 5 SR-vs-3B reco. 6 GL/SR key = state+invoice-no.
7 HSN master check (all 153 codes valid). 8 Bucket Bridge (sign rule). 9 THE MASTER (one
workbook one source). 10 advance ledger tie (all 8 states zero). 11 originals into register
BF/BG. 12 S9 per last year + 9C Table 5 refs + AV state-level pointer (no doc-level FS
exists). 13 Queries (Draft) from last year's query bank (6 re-tested + 3 new).

## Traps (each cost a debugging round — do not relearn)

- openpyxl saves DROP shapes (buttons) — always re-run `readd_buttons.py`; pivots survive.
- COM AddShape/AddPicture fail 0x800A03EC unless the sheet is `.Activate()`d first.
- COM parameterized `.Address` property isn't callable via dynamic dispatch — compute column
  letters in Python.
- Row/column INSERTS must go through COM (Excel rewrites all formulas); openpyxl inserts
  corrupt every absolute reference silently.
- `iter_rows(values_only=True)` on a non-data_only load returns FORMULA STRINGS — copied
  sheets then carry dead refs.
- INDEX() on a blank cell returns 0 — guard `N(...)=0` (falsely flagged 279 CNs BEYOND).
- SUMIFS empty-criterion-cell means "=0"; blank criteria need `"="&IF(x="-","",x)`.
- Excel refuses files where a TEXT cell starts with "=" — guard every text write.
- Excel AutoFilter dropdown: values hidden by OTHER active filters don't appear (user
  chased a "missing" negative value — it was a leftover filter).
- SUBTOTAL rows respond to filters; goldens must be checked unfiltered.
- Filter on the HELPER column (Adv Bucket = Received → 38,16,72,455.56 exactly), never on
  raw doc types (MOB ADV REV splits by sign: 30 Received / 1 Adjusted −1,06,066.67).
- Label vs GSTIN keying: "TamilNadu"≠"Tamil Nadu", "Jammu and"≠"Jammu &" — key on GSTIN.
- Last year's workbook hides judgment values INSIDE formulas (`=65090095-M9`) — sweep
  ported formulas for numeric literals.
- 3B has no document level — GSTIN×month is the finest possible vice-versa.
- The step-file DRAFTs in Downloads are superseded working papers; `...DRAFT (1).xlsx` is a
  stale duplicate. The MASTER is the deliverable. Verify the newest file by mtime.

## Year-roll procedure (FY 26-27)

1. Copy THIS master as the template (skill `gst-audit` rule 7 — never invent a layout).
2. Survey new sources vs `contracts/sources.md`; get rulings on every drift line.
3. Rebuild register data in place (same columns), re-run enrichment + matching + CN fill.
4. Recos recompute live (SUMIFS); re-stamp match columns; regenerate exceptions/S7 listings;
   refresh pivot; re-add buttons; INDEX prune to existing sheets.
5. **Backtest**: re-run the pipeline on FY 25-26 inputs → must reproduce every golden above
   to the rupee before the FY 26-27 output is trusted.
6. Openings roll: S6 opening = this year's closing (433,267,107.25 becomes 26-27's 5C-side
   opening); S3 Amendment Check consumes FY 27-28 returns when filed.

## Open / pending (as of 2026-08-27)

- Buffer-range conversion (ranges → rows 6..50,000) + prefilled helper columns + one-click
  `VEL_REFRESH.bat` pipeline: DESIGNED, not yet built (user paused for CA review round).
- 144-PO re-verification against SAP files: pending (share was down; script ready).
- User deleted Sample Selection & S6 GL Advance Check from the master — that evidence now
  lives only in the standalone Step 1/Step 6 files; may need re-adding if CA asks.
- Awaiting: CA's answers to 38 Open Points, client fixes (3 blocking rows, 45 CN originals,
  14 CN ref conflicts), FY 26-27 GSTR-1 exports for the amendment check.
- RCM phase: BUILT — see `gst-audit-vel-rcm` (register + 7 checklist steps in the same master). ITC phase: `gst-audit-vel-itc` (planned).

## After every session on this engagement

Append new CA/Pawan rulings, traps, and format decisions HERE and to
`contracts/sources.md`. If a change alters numbers, it must reproduce the goldens first.

## Changes log - 2026-09-17 (Pawan's review round, file now `..._MASTER (2).xlsx`)

- **Data sheets carry the SOURCE format, our columns LAST** (ruling): `GSTR-1 Data` = Octa
  Sales-Net's 34 columns verbatim (+3 summary-only fields prefixed "Summary:") then
  Source / Bucket / Month / Match Status / Matched SR Doc No (AL..AP). Summary rows sit in the
  same sheet, mapped into the matching Sales-Net columns. `GL Data` = FBL3N's 35 columns A..AI
  verbatim, then LIVE derived cols State / Month / Tax head / Sales-origin? / Amount(+liab)
  (AJ..AN; BP->State lookup lives in hidden `CC Master` D:E), then Matched with SR (AO).
  `3B Data` = State|GSTIN|Month + every Octa 3B table in form order (Overview, 3.1(a), 3.1(d),
  4A(3..5), 4B(1..2), 4(C), 4D(1..2), Payment), headers on row 2, then SR Taxable / Diff /
  Vice-versa status LAST. Script `datafix_build.py` remaps every dependent formula by column
  map (11,848 cells) - differential check: all 17 dependent sheets identical before/after.
- **S6 is ONE sheet in the CA's `Advance Format VEL.xlsx` layout** (states across, 3 cols each
  Taxable/CGST/SGST; Opening / Add received (12 month rows) / Less adjusted (12) / Closing /
  Net GSTR-1 / Diff; Remarks col). GSTIN key in the format's blank row 4 (S9 rows 35/36 look up
  by it). Opening = last year's closing row of the format file (cached), audited input with a
  cell comment. `S6 Month-on-Month` deleted; INDEX row dropped. Script `s6_format_build.py`.
- **S9 Other Income** row 10 live from `FS Revenue Data` rows tagged Account="Other income"
  (source: Statewise PnL FY25-26.xlsx Sheet2 r9 - Maharashtra only). Row 9 now excludes that
  tag. FLAG: PnL audited total (B9) is Rs 1,00,000 higher than the state split.
- TRAP: openpyxl read_only `ws.cell(r,c)` random access is catastrophically slow on a 3k x 76
  sheet - always `iter_rows`.

## Changes log - 2026-09-17 (evening): '3B Birds Eye View' sheet (Priyesh's format)

- Meeting transcript (Downloads\GST Audit -Vikran-20260917 - Transcript (English).md, line 74):
  Priyesh asked for the 3B in the master to be a FULL birds-eye view (every 3B table in form
  sequence, reversals/reclaims/ISD included), vice-versa column at the end. He supplied his
  layout: `Downloads\Birds_Eye_View.xlsx` (another client's, PAN AAGCR6808Q).
- RULING (Pawan): build it as a SEPARATE sheet `3B Birds Eye View`; do NOT overwrite `3B Data`.
- Built by `bev3b_master.py`: 3-row header (r5 group band merged / r6 exact Section text /
  r7 exact Type text = SUMIFS criteria; template styles copied), 67 data columns incl. the
  nil tables Octa omits (3.1.B/C/E, 4.A.1/4.A.2, Cess everywhere - amber, read zero), tie-out
  block (row sum vs source column sum vs diff - all 228 rows = 0), then OUR three columns LAST
  (SR Taxable live / Diff vs 3.1.A / Vice-versa status). Per state 12 months + TOTAL; ALL
  STATES block; notes. 18 raw Octa matrices embedded as HIDDEN `GSTR-3B XX` sheets (row 1
  reserved, header r3, data r4+) so every cell is a live SUMIFS. Haryana absent (no file).
  Ties to `3B Data` to the paisa. INDEX row under Sales.
- TRAP: verify scripts must resolve tie-out/our columns by HEADER NAME (row 7) - I misread
  them by offset once and chased a phantom tie-out failure.
