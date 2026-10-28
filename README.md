# Bordereaux Data Cleaning & Validation

**Messy insurance bordereaux in — a clean, validated file out, with every correction recorded and every judgement call flagged for you, never buried.**

Built by **Waqar Ahmed**, spreadsheet & data specialist (7+ years: data cleaning, reconciliation, validation and reporting).

## Who this is for

MGAs, brokers, coverholders and insurance operations teams who receive bordereaux that arrive with wrong years of account, malformed UMRs, broken dates, duplicate certificates and totals that do not reconcile — and who need a defensible clean file plus an audit trail, not a quietly "tidied" spreadsheet.

## What the proof in this repo shows

A 200-row sample bordereau with **74 deliberately planted errors** was cleaned and validated end-to-end. The results were then checked by an **independent verifier** (`verify.py`) that does not reuse the cleaning pipeline's logic — expected values come from the error ledger written when the sample was generated, consequential corrections are proven arithmetically, and structural sweeps compare raw vs clean field by field.

| Measure | Result |
|---|---|
| Rows | 200 in → 200 out (order and certificate references preserved) |
| Planted errors | **74 of 74 accounted for (100%)** |
| — Fixed to the expected value | **55**, each recorded in `exceptions-corrected.csv` |
| — Flagged for client action (values *not* altered) | **19**, each annotated in `exceptions-client-action.csv` |
| Consequential recalculations | 11, each proven delta-for-delta |
| Unexplained numeric changes | **0** |
| Gross totals | Reconcile exactly: 2,450,433.56 raw = clean |

Error classes covered include: wrong year of account, malformed UMRs, duplicated certificates (both rows of each pair flagged), expiry-before-inception dates, missing sectors, blank currency, unauthorised risk codes, and commission/net-premium recalculations — with net and settlement totals changing **only** on the rows the ledger says they should.

Files: the raw file (CSV + XLSX), the clean file (CSV + XLSX), the error ledger, both exception lists, a mapping template, the validation report, the full method (`METHOD.md`), the verification record (`VERIFICATION.md`), and the re-runnable pipeline + verifier scripts.

## What you receive on a real engagement

- A clean bordereau (XLSX + CSV) in your template structure
- A **Corrections list**: every change made, row by row, with before/after values
- A **Client Action list**: every item that needs your decision, annotated in the file itself, values untouched
- A validation summary: totals reconciliation and the checks run
- Trial available: **up to 1,000 rows for a fixed £495** — so you can judge the work on your own format before any ongoing arrangement

## How the work runs

1. **You send** a sample file (or your template) and the rules that matter to you (UMR format, section/risk-code rules, currency handling).
2. **I map** your file to the validation rules and confirm the mapping with you in writing.
3. **Clean + validate** — correctable errors fixed and recorded; anything needing a decision is flagged, never guessed.
4. **Independent verification** — totals and sweeps re-checked before delivery.
5. **Delivery** — clean file + both exception lists + validation summary.

## Tools

Python · Excel / Google Sheets · CSV/XLSX bordereaux formats

## About this sample

Proof of capability on **100% synthetic data** — fictional insureds, certificates and amounts; no real bordereaux or client data is used or shown. The method and verification approach are exactly what runs on a client file.

## Work with me

I take on freelance projects through Contra: **https://contra.com/mr_world_z7lsrx86**

Communication is **written-only** (chat/email) — clear briefs, documented handovers, no calls needed. Payment via Payoneer or bank transfer.
