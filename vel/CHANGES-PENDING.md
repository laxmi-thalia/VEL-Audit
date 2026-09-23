# VEL GSTR-9 / 9C — pending changes list

The running list of changes agreed for `VEL_GST_Audit_FY2025-26_MASTER (2).xlsx`. Nothing here is
applied to the master until Pawan says go, item by item.

- **Authoritative source:** the Word file dated 23-09-2026 in Rashid's shared folder. This file is
  the working copy of that list plus everything captured from the recordings.
- **Status values:** `todo` (agreed, not started), `wip`, `done` (applied and verified in the
  master), `blocked` (waiting on data or a decision), `open` (a question, not yet a change).
- Add a dated line under an item when its status changes. Do not delete an item; mark it `done`.

Sources: M1 = meeting 23-09-2026 14:55 IST (2h01m), M2 = meeting 23-09-2026 16:59 IST (13m).
Transcripts in `C:\PROJECTS\accountic\reports\meeting-2026-09-23-*.md`.

---

## A. Changes to the workbook

| # | Change | Sheets it touches | Source | Status |
|---|---|---|---|---|
| A1 | **Reco remarks must name the side at fault, on BOTH sheets.** A line that matches on invoice number but differs on GSTIN should say whether the SUPPLIER (vendor) or the RECIPIENT GSTIN is wrong, not only "recipient GSTIN differs". The vendor-GSTIN-mismatch remark must appear on the ITC Register AND on the 2B sheet (Pawan 23-09). Worked example from Pawan: `GSTR-2B Apr25-Aug26` row 18, company `01AAECR0503Q1ZM` (J&K), supplier GSTIN `06AABCF4798C12N`, doc `HR/FF/16676` dated 18-04-25 — the supplier GSTIN carries `2` where the 14th character must be `Z`, so it is malformed yet still passes a loose 15-character shape test. The remark should say the 2B's own vendor GSTIN is wrong, not the register's. | ITC Register 2025-26, GSTR-2B Apr25-Aug26 | M1, Pawan 23-09 | todo |
| A2 | **Carry forward every prior-year remark**, not the three columns currently provided. Hard requirement: when new remark rows are inserted, the existing formulas must not shift. | ITC Register 2025-26 and wherever LY remarks are mirrored | M2 | todo |
| A3 | **Two remarks columns on the differences**, so each identified difference carries its reason. Last year's reason list was compared with this year's and one reason was missing. | Table 13 & 6A1 differences (and the tax-comparison differences) | M1 (after the break) | todo |
| A4 | **Table 13 must include RCM, not only ITC.** The FY 26-27 claims of FY 25-26 invoices belong there. ITC was brought from the monthly workings, RCM never was. Source is the RCM register's 3B-claim-month column (April 2026 entries), in the same format as the ITC lines. Unclaimed items come through the receivable GL reconciliation. | ITC Register 2026-27, RCM Register, ITC Summary (Table 13 block) | M1 | todo |
| A5 | **Table 8C from the FY 26-27 ITC register** filtered to invoices dated 25-26, instead of from 2B as last year. (8A stays from 2B updated to August; 8D from the tax comparison report.) | ITC Summary (Table 8 blocks) | M1 | todo |
| A6 | **RCM paid vs ITC claimed is built wrong.** The claim lags the payment by one month (March paid, April claimed). Show BOTH months side by side instead of leaving the corresponding month implicit, and split by tax head (IGST / CGST / SGST) instead of total tax only. | RCM Paid vs ITC Claimed | M2 | todo |
| A7 | **New report: ITC Register vs GSTR-2B, summary level, state-wise and month-wise**, carrying the GSTR-3B claim month. Register side excludes correction entries and RCM; 2B side filters Reverse Charge = No and ITC Available = Yes; years compared like for like (25-26 vs 25-26, 24-25 vs 24-25). | new sheet | M1 | blocked — needs A12 (the FY 24-25 3B claim months) |
| A8 | **Match key must include state.** It is supplier GSTIN + invoice number today, so one supplier billing the same invoice number in two states cross-matches and invents differences. Fix: state + supplier GSTIN + invoice number. | wherever the 2B match key is built | M1 | todo |
| A9 | **Drop B2B where B2BA exists.** Both arrive in 2B and the practice is to keep the amendment. Duplicates surfaced in Chhattisgarh. ClearTax's report does not zero the amended rows, so they are being removed by hand today. | GSTR-2B sheets | M1 | todo |
| A10 | **ISD 2B to be reconciled against GSTR-3B.** Imports are still pending too. | 2B ISD sheets, a new reco | M1, M2 | todo |
| A11 | **Remove duplicate invoices.** Same supplier, same invoice number, same amount repeated across 2B return periods. Screenshot (Pawan 23-09) shows supplier `23ADFFS9422E1Z*` invoices `VEEPL23-24RA03/05/06/07`, 33,000 taxable / 5,940 IGST, FY 2023-24, repeating across 2B return periods 01-09-24 and 05-09-24, the repeats highlighted. **Open question:** Pawan said VEL **Assam**, but the `State folder` column on every row reads **Chhattisgarh** (and M1 put the amendment duplicates in Chhattisgarh). Confirm which state before touching anything. Check first whether these are B2B/B2BA amendment pairs (see A9) rather than true duplicates — the fix differs. | GSTR-2B sheets | Pawan 23-09 | todo |
| A12 | **Fill `GSTR 3B Month` on the FY 24-25 2B sheet.** Column E of `GSTR-2B ITC Data` (the sheet rebuilt from the PAN-level export `PAN GSTR2B 2024-25.xlsx`, Inv+CDN document level) is blank on every row while `2B Return Period` is populated (01-04-24, 01-05-24 ...). The claim month is the month the credit was actually taken in GSTR-3B, which is not the 2B return period. **Decide the source first:** the client's monthly workings, or derived by matching each document to the ITC register's 3B claim month. Unblocks A7. | GSTR-2B ITC Data | Pawan 23-09 | todo |

## B. Findings that change numbers, not yet decided

| # | Finding | Source | Status |
|---|---|---|---|
| B1 | **Permanent reversals net to zero.** The 2B carries the invoice positive and the credit note negative, so the pair cancels, while the firm's working reversed only the credit note. Whatever sits in Table 4A(5) must flow to 6(7). Reading: the client under-reversed and owes the difference with interest, put at 18% a year across two years. Rectify this year or repeat last year's treatment was escalated, not settled. | M1 | open — needs the CA's decision |
| B2 | **The 6,200 item belongs in Table 4D(1) at its gross figure.** The October Arunachal Pradesh working carried about half, and the same wrong figure went into the filed 3B, so the working and the return agreed and no reconciliation could flag it. Counting it once extra had doubled a 1,24,000 difference, which nets to zero. | M1, M2 | done in the review — confirm the master reflects the gross figure |
| B3 | **Excel converts invoice numbers into dates.** "06/24-25" became June 2025 and silently broke every lookup on it. Needs a guard wherever invoice numbers are read or written. | M1 | todo |

## C. Still outstanding (tracked, not changes to make yet)

- Tables 12B, 12C and 13.
- The GSTR-9C tables, to be checked at filing time.
- Table 5O, blocked on the client providing financials.
- The GL reconciliation.
- The refund being claimed will be included even though it is not normally shown.

## D. Process

- Keep this list current. It exists because decisions were being forgotten between sessions (M1).
- Meetings are recorded from the start (M1).
- Changes are delivered before the review meeting, not during it (M2).

---

## Change history

| Date | Item | What happened |
|---|---|---|
| 23-09-2026 | — | List created from the two 23-09 recordings. Nothing applied to the master. |
| 23-09-2026 | A11 | Added from Pawan's screenshot: duplicate invoices to remove. State to be confirmed (Assam vs Chhattisgarh). |
| 23-09-2026 | A12 | Added: fill the GSTR 3B Month column on the FY 24-25 2B sheet. A7 re-pointed at it as the blocker. |
| 23-09-2026 | A1 | Extended from Pawan's screenshot: the vendor-GSTIN-mismatch remark must appear on the register AND the 2B sheet; worked example recorded. |
