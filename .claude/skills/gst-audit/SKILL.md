---
name: gst-audit
description: >
  Working method for the GST-audit automation (gst-audit-engine + accountic
  gst_audit module). Use whenever editing run_amar.py / engine/*, the working
  Excel deliverables, the portal module, or deriving audit queries for a GST
  annual-return engagement (GSTR 9 / 9C).
---

# GST-audit engine — working method

Repo: `C:\Users\rashi\OneDrive\Desktop\Amar Impex - Automation\gst-audit-engine`
(origin: github.com/RashidFaisal76/gst-audit-engine). Portal module:
`backend/app/services/gst_audit/` + `backend/app/api/gst_audit.py`.

## Non-negotiables (learned the hard way — do not relearn)

1. **Backtest-first.** Any change to the engine must keep the FY 24-25 golden
   at **28/28** (`compare_golden.py`), and any new *automation of judgment*
   must **reproduce the golden's manual decisions exactly** before it ships —
   run the golden year through AI/portal mode (config with real filed figures)
   and demand rupee-exact equality on the automated cells. If the automation
   can't reproduce last year's elections, the rule is wrong — go read the rule
   off the golden file itself.
2. **Read rules off the golden, never invent them.** Every convention (tag
   vocabularies, which rows are 6A1 vs "Refer Anupam Sheet", Table-13 Yes/No,
   Rule 42 splits) exists somewhere in the golden workbook's own columns.
   Deriving a rule from first principles without checking the golden costs a
   full review round every time.
3. **No plugs, ever.** Golden formulas may carry year-specific rounding plugs
   (`-8.32`, `+1`, `-952.46`) — strip them when porting. Unexplained
   differences stay VISIBLE (a FINDING flag or a 0.36 residual), never
   absorbed. Elections (filed figures, adjustments) enter via config, tagged.
4. **AI proposes, Rashid elects.** Every automated judgment cell gets a
   visible "AI automated - kindly confirm & check" tag cell beside it
   (`_ai_tag`) plus a comment. Never silently concede/elect (see
   no-unelected-concessions). Client-dependent figures get `_draft_note`.
5. **Never harvest judgment cells** from a final working into a new year
   (narrative reasons, plugs, elections). Templates are copied, then
   year-token-shifted (`_shift_sheet_years`) and judgment content blanked.
6. **Formula-driven, not value-pasted.** Anything the office had as a formula
   must be a live formula in the output — layout-independent where layouts
   move (SUMIF by label, VLOOKUP on "Grand Total") because `live_pivot.py`
   replaces Rate wise / HSN Pivot with real PivotTables whose rows differ.
   Rashid election (2026-08-03, GENERALISED): **every derived cell that
   mirrors or aggregates another cell must be a LIVE formula**, even where
   the golden pasted values. Applied so far (repeat on every workbook):
   - 'Sales Books-GSTR1 and 3B' 3B block (R27:R38) -> `='GSTR-3B'!` 3.1(a)
     rows (`fix26_3b_formulas.py`);
   - 'ITC-Books-GSTR3B' 3B block (R21:R32): B/C/D/E<-4A(5), G<-4A(1),
     H<-4A(2) — ClearTax label is "(2) Import of Service", SINGULAR —
     I<-3.1(d) taxable, J/K/L<-4A(3) (`fix26_itcbooks_formulas.py`);
   - GSTR-2B availed cols AF/AG/AH -> `=<raw tax col><row>` ("availed as
     reflected"); partial-availment judgment cells STAY literal.
   Method: locate rows by LABEL, self-check mapped value == literal being
   replaced (24-25/backtested files); on template-replay drafts the old
   literals are STALE — replace regardless, sanity-check the anchor against
   the report's FY-total. **Self-check must cover EVERY cell INCLUDING
   BLANKS** — goldens deliberately leave some block cells empty where the
   filed report has values; revert any formula whose recalc value differs
   from the golden literal (`fix26_itcbooks_parity.py`), then re-run gates.
7. **ALWAYS build on the client's own template — never invent a layout.**
   (Learned on Amar International FY 25-26: a "clean" self-designed workbook
   was rejected on sight.) Every output workbook starts as a COPY of the
   client's prior working (the golden, or last year's GENERATED which carries
   the golden presentation): full sheet vocabulary (incl. trailing-space form
   names), letterheads, R1 diff/R2 as-per-filed/R3 SUBTOTAL top rows, A..Q
   register geometry with the TEXT-month/rate/Gross formulas, Summary pivot
   grids at their load-bearing anchors (ITC-Books refs `Summary!C26` etc.),
   styles/panes/widths. New-year data is swapped IN at the same coordinates;
   year tokens shift in one pass; judgment content blanks (rule 5). This
   applies to every new FY *and* every new deliverable — pick the template
   first, then write data into it.

## The verification gauntlet (run before showing anything)

```bash
ENGINE="C:\...\gst-audit-engine"; PY=backend/.venv/Scripts/python.exe
# 1. golden year, standalone
cd $ENGINE && $PY -X utf8 run_amar.py && $PY -X utf8 compare_golden.py   # => 28/28
# 2. golden year, AI/portal mode (real figures via GST_AUDIT_CONFIG json)
GST_AUDIT_CONFIG=<fy2425_ai_cfg.json> $PY -X utf8 run_amar.py && $PY compare_golden.py
# 3. current-year deliverables (draft + template + data request)
$PY -X utf8 <scratchpad>/generate_deliverables.py
# 4. Excel COM full recalc: zero error cells + key cross-foots
# 5. scans: stale year-tokens, stale narratives (Rs|INR|DRC|ARN), hardcoded
#    totals, broken refs (NB sheet name 'GSTR 9 Form ' has a TRAILING SPACE)
```

Verify the **newest file by mtime** — Excel locks force a DRAFT→v2→v3 ladder
and stale-file verification has burned us repeatedly.

**`output\` holds ONLY deliverables** (Rashid 2026-08-04): the compare gates
write their full-recalc snapshots to **`output\_gate\...- recalc.xlsx`** —
scratch, regenerated every gate run, safe to delete. Never review, share, or
edit a `- recalc` twin (Rashid reviewed one and saw pre-fix state — "duplicate
files" / stale-blanks confusion). Superseded DRAFT versions get deleted once
their successor verifies. If a fix errors with PermissionError, a deliverable
is open in Excel — ask Rashid to close it, never write to a new twin.

## Trap list

- **Template-replay (new-FY roll-forward) traps** (Amar Intl FY 25-26):
  (1) **sheet-rename collision** — year-shifting sheet names makes the
  incoming spill sheet's new name equal the old outgoing sheet's current
  name; openpyxl silently appends '1' and any `if name in wb.sheetnames`
  paste-guard then skips — rename in dependency order or via temp names,
  and verify every expected sheet name after the shift. (2) **`[1]`/`[2]`
  external-workbook refs** — formulas copied from a working may reference an
  EXTERNAL copy of a sheet (`'[1]FY 24-25 claimed…'!E59`); they survive the
  template copy and error or silently read stale data — scan for `[N]` in
  every formula and rewire to the in-workbook sheet (spill totals: incoming
  layout E59/F59/G59 vs outgoing layout F101 — repoint, don't assume).
  (3) form sheets' live refs into pasted report sheets (3B) only compute
  once the paste matches the golden ROW layout — until then blank + DRAFT
  flag those cells; never leave error cells or fake values. (4) a pasted
  sheet may contain its own Total row — column-summing it double-counts;
  find the sheet's own total row instead. (5) **THEME-COLOR REMAP** — cells
  styled with theme-indexed fills (accentN + tint) change hue when the
  workbook's theme differs: openpyxl-written files get the DEFAULT 2007
  theme (accent4 = 8064A2 PURPLE) while the goldens use the modern palette
  (accent4 = FFC000 AMBER) — the office's amber banding renders purple.
  Fix: transplant the golden's `xl/theme/theme1.xml` into every output
  (zip surgery; values/styles untouched), then re-run the gates. Check
  accent palettes whenever output colours look wrong. (7) **a TB is a GRID,
  not a voucher register** — `read_ledger_positional`/voucher readers return
  ZERO rows on it SILENTLY and the data swap writes an empty sheet; paste
  the TB export directly and ASSERT row counts after every sheet swap.
  (8) **month-token year-shift cascade** — sequential replaces double-shift
  ('Apr-24'→'Apr-25'→'Apr-26' in one pass); use placeholder substitution for
  month tokens exactly as for FY tokens. Prefer pasting the new-year export's
  own letterhead/period rows over shifting the template's.
  Rashid rule (2026-08-03): **TB Eligibility/RCM/Remarks columns are
  AI-generated for EVERY ledger row** — precedent map first, then the GST
  rule engine (9(3) RCM heads, 17(5) blocks, NA/Ignore for BS & control
  ledgers), bare 'Query' only when genuinely unresolvable; every cell tagged.
  Classifier keywords need BOTH word boundaries — `rent\b` matched
  "cur**rent** A/c" and mis-tagged partner Current accounts as RCM; use
  `\brent\b` (fixed 2026-08-04, 21 rows corrected across all 4 books).
  Retro-filling a BACKTESTED year is gate-neutral IF only blank cells get
  TEXT (gates flag only generated-only NUMERIC cells) and office-tagged rows
  are never touched (`fix26_tb_retro.py`).
- **Spill-out / lapsed sheets are DERIVED, never hardcoded** (Rashid): port
  of Impex `extract_held_itc` — anchor the March tab of the client's monthly
  Conso Reco on the label "Amount not considered in GSTR 3B" (text-located,
  never fixed cells), rows beneath until a stop label. Backtested: GJ 24-25
  conso ties golden spill-out F101 135,639.99 EXACT (month level); MH 24-25
  = the golden's 3 Apeda rows / 351 EXACT. Goldens keep their invoice-level
  listings; new years generate from the conso (`derive26_spillout.py`).
  Filename traps: 'Mar' search must be case-sensitive ('Amar' contains
  'mar'); names carry DOUBLE SPACES ('_Mar  25') — search loosely, filter.
- **'FS vs GST returns' Comments are AI-GENERATED, never replayed** (Rashid
  2026-08-04): precedent map `fs_comments` in `{gj,mh}26_precedent.json`
  (line label -> office comment, mined from 24-25 by `mine26_fs_comments.py`)
  + rules: diff==0 -> 'ok'; known non-GST heads (drawback, forex, capital
  gain/loss, interest, IT refund, write-offs) -> 'Non-GST Transaction';
  quality claim/weight shortage -> purchase-DN note; new lines -> Query.
  All flagged in the AI Review column; office/client comment cols stay
  theirs per ownership below.
- **Comment-column ownership (Rashid 2026-08-04):** register sheets carry
  O 'Comments' (office analysis), P 'Amar comments' (CLIENT fills), Q 'DPS
  comments' (DPS fills) — in a new-year draft ALL THREE stay BLANK for
  humans; AI review notes go in their OWN 'AI Review' column appended after
  the template columns, never in P/Q. Also restore pivot-HELPER columns
  after a data swap (GJ Sales Data R 'Pivot Remark', MH Purchase Data W
  'Pivot Remarks') as live IF-formulas over the tag column — the live
  PivotTables group on them.
- **After any template-replay build, run a COMPLETENESS SWEEP** — every
  sheet must be rebuilt from this year's data, derived, or explicitly
  NIL/PENDING-bannered. Stale remnants hide in: office-input listing sheets
  (spill/lapsed/Queries/Additional impact), Summary/Ratewise/RCM-Summary
  pivots, HSN, Table 8A, and month-label DATETIME cells (string year-shift
  skips datetimes). Audit: scan for prior-FY datetime labels and for sheets
  whose values didn't change from the template. (6) **HIDDEN ROWS
  ride in from the golden** — goldens get saved mid-filter, and the replay
  copies row-hidden state verbatim (GJ carried 9,231 hidden rows). Rashid's
  standing instruction (2026-08-03): **deliverables ship with ALL ROWS
  VISIBLE** — unhide every row dimension + clear active autofilter criteria
  before handover (keep the filter dropdowns, keep deliberately-hidden
  SHEETS like Sheet3/Queries(hidden), keep hidden helper COLUMNS).
  EXCEPTION (Rashid 2026-08-04): the **'GSTR 9 Form '/'GSTR 9C Form '
  canvases keep the golden's COLLAPSED break-up rows** (ClearTax drill-down
  remnants with dead-hyperlink styling — hidden by the office by design;
  unhiding them reads as "generated hyperlinks that weren't in the actual
  sheet"). Restore their hidden state from the golden fx file's
  row_dimensions (`fix26_form_rows.py`).
- **Golden workbooks contain their own defects — classify before reproducing**
  (learned on Amar International, both GSTINs): (1) **stale pasted pivots** —
  the office re-tags/edits register rows after the last pivot refresh, so the
  golden's Summary pivot can disagree with the golden's OWN register; verify
  by re-summing the golden's register per tag; generate LIVE pivots and record
  the pivot cells as delta-enforced accepted deviations. (2) **Excel 15-digit
  mangling** — >15-digit numeric voucher cells lose their last digits in the
  golden; engine output keeps true text, the COMPARE layer tolerates
  first-15-digit equality. (3) **SUBTOTAL filter snapshots** — a golden saved
  with an active autofilter shows filtered SUBTOTALs; a regenerated file has
  no live filter criteria, so freeze those header cells to cached values.
  (4) **circular self-range formulas** (=SUM(P92:P96) inside its own range) —
  freeze to cached value. (5) **paste type-coercion** — golden pastes turn
  digit-string vouchers into numbers and downstream VLOOKUPs are type-strict;
  writers must coerce identically. Dr/Cr number formats must be resolved
  **per cell**, never per file (a file can have Cr formats on Gross Total but
  neutral accounting formats on the ledger columns).

- **openpyxl**: `ws.cell(row, col, value=None)` is a **silent no-op** — use
  `.value = None` to blank. MergedCell writes need the anchor (or unmerge).
  `read_only=True` yields ragged rows truncated per-row; normal mode pads.
  `load_wb(data_only=True)` flattens formulas — copy formulas needs
  `data_only=False`.
- **Tally exports**: Dr/Cr sign lives in the cell **number format**
  (`"" Dr"` / `"" Cr"`), values always positive — authoritative over
  narrations. Duplicate header columns: FIRST occurrence wins (far-right
  copies can carry the opposite side). Ledger rate names: `@` is optional
  (`Output CGST 6%`). A ledger is NOT homogeneous — classify rows by
  **voucher type** (sales rows inside expense ledgers are re-billings that
  belong in Sales Data Base).
- **Octa exports**: real report hides behind a near-empty "Overview" first
  sheet (`_octa_report_sheet`). 3B monthly matrix: sections in col B, heads in
  col C, months D..O; some blocks (4(C) Net ITC) carry the section label only
  on their header row. Reversal rows are NEGATIVE in the report — the form
  wants positive. Overview must be rebuilt (ClearTax semantics: taxable =
  3.1(a) only).
- **Excel COM** is used for: xlsb→xlsx conversion (reading golden formula
  text — strip pivotCache parts + .rels first), live pivots, and the recalc
  gate. On this box render PDFs with PyMuPDF+pytesseract (no poppler).

## Input contracts

- `inputs/ai_suggestions.json` — the future AI-API contract:
  `rcm_nature` (narration substring→label), `tb_eligibility`
  (collapsed-lowercase ledger→[remark, eligibility, rcm]), `tnf_remarks`
  (party-lowercase→FCM/RCM/Query), `queries` ([area, text, amount]).
  Loaded only on portal runs (`_cfg_path`).
- `inputs/golden_form_formulas.json` — golden formulas ported verbatim
  (fixed-layout refs only; token-shifted at port time). Plugs stripped.
- `GST_AUDIT_CONFIG` json — base/out_dir/gstin/prefilled/out_name +
  `config_9c` figures (unsupplied figures are ZEROED on portal runs, never
  defaulted to last year's).
- Working-inputs template = **client inputs only** (Anupam, 4B1, FS-pnl,
  2B-workcols pre-filled with raw 2B, prior-year spillover paste, Table 8A SS,
  Rate Diff ledger). Engine-built sheets (spill-out, TB check, 9/9C canvas)
  are `required=False` in the manifest and never gate readiness.

## References

- `references/gstr9-9c-reference.md` — distilled GSTR 9 / 9C table-by-table
  rules, cross-checks and FY 2024-25 changes (from Bharat's "Analysis of GST
  Returns"). Read it before touching form logic or drafting queries.
- `references/amar-impex.md` — client-specific conventions and elections.
  One file per client; promote rules that repeat across clients into this
  SKILL.md.
- `references/gst-manual/INDEX.md` — the **GST law knowledge base**: CGST Act +
  Rules, IGST Act + Rules, GSTAT and Cess rules, split one file per
  section/rule from the Garg & Garg manual (7th half-yearly edn, updated
  01-Jul-2026). Read the relevant section BEFORE drafting any law-dependent
  check (invoice compliance/S.31, credit-note time bar/S.34, ITC/S.16-17).
  **Personal-use copy: gitignored — never commit or redistribute its text.**
  Refresh each half-year from the latest edition PDF (`build_gst_kb.py`).
- `references/amar-international.md` — Amar International (2 GSTINs, own
  template ≠ Impex). Full golden decode + build notes live in
  `C:\Users\pawar\Downloads\Amar International - Automation\ENGAGEMENT_NOTES.md`.
- **VEL (Vikran Engineering, 19 GSTINs)** has its own per-phase skills — sales phase:
  `.claude/skills/gst-audit-vel-sales/SKILL.md`; engine + contracts at
  `c:\PROJECTS\accountic\gst-audit-engine\vel\`. RCM/ITC skills follow in their phases.

## After every engagement

Append new lessons here (traps, conventions, corrections Rashid gave) and to
the client file. Additions that change engine behavior must survive the
gauntlet first.
