# gst-audit-engine

Deterministic automation for DPS & Co's GSTR-9 / 9C annual-return audits (Thalia Technologies).
The audit deliverable is one Excel master workbook per client-year; everything here builds, checks
and rolls that workbook forward. No LLM calls anywhere in the pipeline.

## Layout

| Path | What |
|---|---|
| `.claude/skills/gst-audit/` | General working method (template-first, backtest-first, no plugs, AI-proposes-CA-elects). |
| `.claude/skills/gst-audit-vel-sales/` | VEL engagement - Sales phase: recipe, CA rulings log, traps, formats. |
| `.claude/skills/gst-audit-vel-rcm/` | VEL - RCM phase. |
| `.claude/skills/gst-audit-vel-itc/` | VEL - ITC phase (register, 2B reco, 6A1 extract, ITC Summary, Table 13/12). |
| `vel/contracts/` | Source-data contracts (paths, header contracts, vocabularies) and per-sheet format specs. |
| `vel/reference/` | Style dumps of last year's formats (`formats.json`, `index_dump.json`, `cn_fill_log.json`). |
| `vel/scripts/` | Every build / fix / verify script that produced the FY 2025-26 master, in the order the skills describe. |

Open this folder in Claude Code and the four skills load automatically; the VEL skills point at the
scripts by relative path (`vel/scripts/...`).

## Engagement-local files (deliberately NOT in this repo)

Client figures and meeting content stay on the engagement machine. `.gitignore` excludes:

- `vel/contracts/goldens_*.md` - the per-FY backtest targets (every verified figure, to the rupee)
- `vel/reference/salesreco_dump.json`, `vel/reference/ly_itc_formulas_fy2425.txt` - last year's sheet dumps (contain values)
- `vel/reference/transcripts/` - translated CA-session transcripts
- `.claude/skills/gst-audit/references/amar-*.md` - Amar engagement notes
- all `*.xlsx / *.xlsb / *.pkl` - client workbooks and intermediate pickles

A year-roll therefore needs the engagement machine (or a copy of those files) to run the backtest.
The scripts themselves carry the client's GSTINs and server paths, so this repository must stay private.

## Running

Python 3.12 with `openpyxl`, `pandas`, `pywin32` (Excel COM for recalc / buttons / pivots), `pyxlsb`.
Run scripts with `python -X utf8 <script>` from `vel/scripts/`. Each script's docstring states its
inputs, outputs and where it sits in the pipeline; the skills give the order.
