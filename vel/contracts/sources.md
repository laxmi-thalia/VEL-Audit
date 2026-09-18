# VEL — source data contracts (FY 2025-26 baseline)

Every year-roll starts by DIFFING the new files against this contract. Any deviation
(missing sheet, renamed header, new vocabulary value) STOPS the build and is listed for a
one-line human ruling. Rulings are appended here — never patched silently into code.

Server root (both work; share flaps — retry): `\\server\GST FOLDER` = `\\192.168.1.69\GST FOLDER`
Engagement root: `<root>\GST Returns\GST Audit & Annual Return\FY 2025-26\1. Corporate Clients\VEL\`

## 1. ClearTax monthly sales registers (the register's source of truth)

Path: `Clients Data\05.08.2026 Main Data\<NN Month YYYY>\Cleartax Sales Register <Mon> <YYYY>.xlsx`
- Filenames are INCONSISTENT (case, spacing) — search loosely, filter exact. Mar-26 file was
  supplied separately by Pawan ("Cleartax sales register Mar 2026.xlsx").
- Use the **"Working" sheet** (header row varies 1-2); IGNORE "Cancelled" sheets.
- Some months have 2 header rows (see `sr_compare2.py` SRC dict: sheet + header-row per month).
- **Map columns by HEADER NAME, never position.**
- Jul-25 file is legacy `.xls` → converted copy `July2025_converted.xlsx`.
- One document can span multiple rows (page split) — collapse phantom duplicates.
- Vocabularies (closed — new value = STOP):
  - Document Type Code: `INV, CRN, DBN, MOB ADV REC, MOB ADV ADJ, MOB ADV REV`
  - Supply Type Code: `B2B, B2C, SEZWP, EXPWP, DEXP` (as seen; B2C rows are summary-level in GSTR-1)
- **MOB ADV REV sign rule (CA-approved):** positive taxable → Received, negative → Adjusted.
  This is the ONLY classification that ties GSTR-1 (Rs 9.49) AND the advance GL (exact, all 8 states).

## 2. SAP monthly sales registers (enrichment: GL name, PC, CC, PO/SO)

Path: `Clients Data\05.08.2026 Main Data\<NN Month YYYY>\Sales Register <Mon> <YYYY>.XLSX`
- Exact filenames have double-space / case quirks — recorded in `sr_compare2.py` SAPF list.
- Join keys: **IRN first, then ODN** (ODN = ClearTax Document Number).
- Fields used: GL No/Name, Profit Centre, Cost Centre, Business Place, `Contract No` (= PO),
  `Sales Doc Number` (= SO).
- 540 register rows have no SAP match (advances + B2C summary rows) — expected, not an error.

## 3. GSTR-1 as filed (ClearTax "Filed" exports)

- Document-level (`Sales-Net`) + summary-level (`SalesSummary-Net`: B2CS, advances) rows.
- IRN = 64-char hash; B2C and advances exist ONLY at summary level (61 rows) — never
  document-match them.

## 4. GSTR-3B (Octa portal exports)

- Table 3.1(a) single clubbed line per GSTIN-month. Real report hides behind a near-empty
  "Overview" first sheet.
- Known-good totals: FY total 3.1(a) taxable 8,669,828,675.34.

## 5. SAP GL — output tax ledgers

Path: `Clients Data\GLs\Outward Tax\` (+ `GST Advance.xlsx` for the advance ledger)
- Ledgers: 2610080100/101/102 = CGST/SGST/IGST output.
- `Reference` = invoice number (the SR match key); `Business place` = state key (AP01, AR01...).
- Sales-origin doc types: RV, DA, DG, DR, XD → only these enter the document match
  (**`Sales-origin? = "Y"` criterion is REQUIRED on every GL SUMIFS** — dropping it once cost
  the whole reco).
- Advance postings carry TEXT references (MOB*/COLLECTION) → state-level reco only.
- **Year/Month filter `2025/01`..`2025/12` is mandatory** (SAP fiscal year: 2025/01 = Apr-25).

## 6. Credit Note Statement (client-maintained)

Path: `Clients Data\Credit note statement\Credit Note Statement.xlsx`, sheet `FY 2025-26`, header row 3.
- Filter col P ("Nature of deviation #") contains "Credit Note against".
- Key: col H `Bill No.` = register CN Document Number.
- Original ref: col Y, pattern `<doc> Dated dd.mm.yyyy` → register BF/BG.
- 40 rows unparseable by design ("Credit Note against Negative PV" etc.); 14 conflicts vs
  ClearTax-sourced originals were kept as ClearTax (see reference/cn_fill_log.json).

## 7. Cost centre master

Path: `Clients Data\Cost Centres_New.xlsx`, tab **"Sheet4"**: col A = PC code, col F = CC narration
(location, e.g. "BHOPAL RO"). 124 pairs; 4 duplicate PCs (first kept — confirm with Priyesh).

## 8. FS revenue extract

Client's tagged SAP revenue extract, state/ledger level, 133 rows. NO document-level FS data
exists → register column AV can only ever be a state-level pointer.

## 9. Reference workbooks

- Last year's audit: `Audit Data of FY 2024-25\VEL_GSTR 9_9C FY 24-25.xlsb`
  (Index + Sales Reco formats replicated from here; styles in reference/*.json dumps).
- Filed GSTR-9C PDF (Gujarat FY 24-25) confirms the advance opening 5C (figure recorded in goldens_fy2526.md, engagement-local).
- HSN/SAC master: `HSN_SAC.xlsx` from CA (HSN_MSTR 21,935 codes / SAC_MSTR 681).

## 10. The 19 GSTINs

01 J&K, 03 Punjab, 06 Haryana, 08 Rajasthan, 09 UP, 10 Bihar, 12 Arunachal, 18 Assam,
19 West Bengal, 20 Jharkhand, 22 Chhattisgarh, 23 MP, 24 Gujarat, 27 Maharashtra,
29 Karnataka, 32 Kerala, 33 Tamil Nadu, 36 Telangana, 37 Andhra Pradesh
(all `AAECR0503Q` PAN; exact GSTINs in every build script's GSTINS list).
State labels must match the register's "My State" EXACTLY ("Tamil Nadu" with space;
"Jammu & Kashmir" with & — key on GSTIN wherever possible, a label mismatch once silently
dropped Rs 4.23L).
