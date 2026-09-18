# VEL — ITC phase source contracts (FY 2025-26 baseline)

Same doctrine as sources.md / rcm-sources.md: survey-diff before build; drift stops for rulings.

## 1. ITC All State (the register's source)

`Clients Data\29.6.2026 Initial Data\ITC All state FY 2025-26.xlsx`, sheet **`Working`**,
header row 6, 31 cols, **41,508 real rows** (used range 1M+ = phantom; stream with an
empty-run break on `Document Number`). Filter `Company = VEL` (user ruling: drop "Others" —
none present in FY 25-26 but filter anyway). TYPE: ITC 40,891 / RCM 446 / ISD 171 — the
RCM rows are the claim-month source for the RCM phase's post-ITC hooks.
- **AMOUNTS STORED AS TEXT** ('20000.0') — coerce every numeric column.
- State names: FOUR variants seen across files — "Madhya Pradesh"/"Madya Pradesh",
  "Tamil Nadu"/"TamilNadu"/"Tamilnadu" — canonicalize to GSTIN immediately.
- `Reference` = vendor invoice no (ITC rows) / nature text (RCM rows).
- `Type for GSTR9` fully empty (derive later); 617 rows blank GST-CREDIT flag; 523 blank
  2B-periods — all flagged, not guessed.
- If data is missing here, fetch from the monthly workings (flag first — user ruling).

## 2. Target format

Last year's final `ITC Register 2024-25` sheet (FY 24-25 xlsb): **83 columns**, header
row 5 — core register, 2B match block (Invoice Level Match / KEY / Countif / B_ / 2B_ / D_),
correct-invoice columns, POS block, GL/eligibility/query block, 8A & 6A1 tail. Last year's
2B_ semantics: matched rows carry the books-level amounts (doc totals APPORTIONED across the
invoice's register lines; doc sums tie; naive doc-total repeats inflate 3.5x). Countif held
decision labels ("Consider"), not counts.

## 3. GSTR-2B (cumulative, the matching universe)

`Audit data of FY 2026-27\04 July 2026\Final\<State>\GSTR-3B <State> July 2026.xlsx`
(names inconsistent: "GSTR 3B Madhyapradesh..." — match loosely), sheet **`GSTR 2B`**:
summary junk rows 1-4, header row found by "GSTIN of supplier" (row 5), data row 6+,
**cumulative from FY 2020-21**. 48,124 rows over 15 states.
- `GSTR 3B Month` = the claim marking ("01 APR 2022" / "Claimed in Feb 2022" / "Lapsed") —
  **maintained ONLY for Bihar and Tamil Nadu**; everywhere else derive claim status by
  matching against the registers.
- Invoice numbers carry a leading apostrophe; `F.Y` column 54% blank — derive FY from
  `2B Return Period` ("01.APR'2022" pattern).
- **Arunachal, J&K, Kerala workings have NO 'GSTR 2B' sheet** — fallback: the FY 25-26
  Portal Reports 2B files (claim markings absent there; FY-limited).
- ISD credits live in the separate `ISD` sheets of the Portal Reports 2B files.
- Raw portal per-month 2B exports (072026_<GSTIN>_GSTR2B_*.xlsx) exist per month folder.

## 4. GSTR-3B Table 4 (Net ITC)

Octa files (same folder as sales/RCM): rows `4.A.*` and `4.B.*` per tax head per month.
**4B reversal rows are NEGATIVE as reported** → Net ITC = 4A + 4B (never A − B).
`4(C)` is a header-only row — always compute. 4A(3) (RCM) and 4A(4) (ISD) captured
separately onto 3B Data — they drive the RCM paid-vs-claimed comparison.
NOTE: added columns on 3B Data currently carry their headers on ROW 1 (the reserved button
row) while original headers sit on row 2 — scan both rows when resolving headers.

## 5. Last year's claims (for unclaimed analysis)

FY 24-25 xlsb `ITC Register 2024-25`: (vendor GSTIN col R, invoice col L) → 8,227 claim
keys. Any FY 24-25 2B row unmatched by BOTH registers = genuine unclaimed candidate.
FY 26-27 Apr-Jul claims NOT yet netted — standing caveat on the candidates sheet.

## 6. Receivables open items

`Clients Data\GLs\Inward GL\Recievables - Open Items.xlsx` — FBL3N style, 42,429 rows,
**no vendor / no Reference columns** — GL-side listing only; cannot match 2B directly
(checklist Sr 58 review item, not a matching input).

## 7. Checklist (Type=ITC)

41 checkpoints (Sr 25-65); ONLY FOUR ordered: 1 data-accuracy (ITCR vs 2B invoice level),
2 ITCR vs 3B Net ITC MoM, 3 ITC POS review, 4 ITCR vs 2B GSTIN level ("Less in 2B" +
vendor recovery). `File to be used` column names each step's inputs. The other 37 are the
unordered backlog (many inquiry-type).
