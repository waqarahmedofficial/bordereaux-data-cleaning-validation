# Validation & Exception Report — Summary (SAMPLE)

**Standard:** Lloyd's Coverholder Reporting Standards v5.2 (20 August 2019, Market Bulletin Y5261) — premium bordereau  
**Coverholder:** Harbourline Insurance Services Ltd  
**Binder:** UMR B0876HLB250001X · Agreement BND-2024-1187 · Year of Account 2024  
**Reporting period:** 2025-12-01 to 2025-12-31  
**Rows processed:** 200 · **Exceptions corrected:** 66 · **Flagged for client action:** 24 items on 24 rows

## Check results

| Check | Ref | Exceptions | Corrected | Flagged | Result |
|---|---|---|---|---|---|
| CHK-01 UMR format & contract match | CR0005 | 8 | 8 | 0 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-02 Agreement reference present | CR0006 | 0 | 0 | 0 | PASS — no exceptions found |
| CHK-03 Reporting period complete | CR0001/CR0002 | 0 | 0 | 0 | PASS — no exceptions found |
| CHK-04 Section present (multi-section binder) | CR0007 | 7 | 7 | 0 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-05 Class of business present | CR0017 | 0 | 0 | 0 | PASS — no exceptions found |
| CHK-06 Risk code present, valid & authorised | CR0016 | 16 | 15 | 1 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-07 Year of account = binder YoA | CR0010 | 10 | 10 | 0 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-08 Certificate reference unique | CR0029 | 10 | 0 | 10 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-09 Insured name & sector completeness | CR0035 / v5.2 | 6 | 0 | 6 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-10 Date logic | CR0030/CR0031/CR0057 | 6 | 0 | 6 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-11 Currency codes (ISO 4217) | CR0020 | 7 | 6 | 1 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-12 Premium arithmetic | CR0059/61/62/64/65 | 11 | 11 | 0 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-13 FX completeness & arithmetic | CR0067/CR0068 | 9 | 9 | 0 | PASS — all exceptions corrected or carried on the Client Action list |
| CHK-14 Mandatory-field completeness sweep | v5.2 core set | 1 | 0 | 1 | PASS — all mandatory fields populated except items on the Client Action list |

## Totals reconciliation

| Measure | As keyed (raw) | Clean file | Difference |
|---|---|---|---|
| Gross premium paid this time | 2,450,433.56 | 2,450,433.56 | 0.00 |
| Commission amount | 576,599.54 | 576,695.29 | 95.75 |
| Total taxes and levies | 248,849.12 | 248,849.12 | 0.00 |
| Net premium to London (original ccy) | 1,625,328.40 | 1,624,889.15 | -439.25 |
| Net premium to London (settlement GBP) | 1,586,187.25 | 1,567,311.16 | -18,876.09 |

Gross premium and tax totals are unchanged — no gross or tax errors were present. Net differences are fully explained by the CHK-12/CHK-13 arithmetic corrections listed in the workbook. Recorded corrections include consequential recalculations: where a commission amount was corrected, the dependent net premium (and, on one EUR row, the settlement-currency net) was recalculated from the corrected components. The clean settlement-currency total excludes the one row flagged for unconfirmed original currency; that row is carried in the original-currency totals only. All other flagged rows are carried at their keyed amounts pending coverholder confirmation.

## Items requiring client action

- Row 3 · HLB/2024/0164 · Certificate Ref: Certificate reference is shared with row(s) 20 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 5 · HLB/2024/0053 · Certificate Ref: Certificate reference is shared with row(s) 182 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 8 · HLB/2024/0104 · Certificate Ref: Certificate reference is shared with row(s) 97 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 9 · HLB/2024/0086 · Risk Expiry Date: Expiry date 2025-07-19 is on or before inception 2025-07-20 → Coverholder to confirm the correct period of cover (source-record fact; not altered by this service)
- Row 10 · HLB/2024/0085 · Risk Expiry Date: Expiry date 2024-12-26 is on or before inception 2025-12-25 → Coverholder to confirm the correct period of cover (source-record fact; not altered by this service)
- Row 20 · HLB/2024/0164 · Certificate Ref: Certificate reference is shared with row(s) 3 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 38 · HLB/2024/0061 · Insured Trade or Industry Sector: Trade/industry sector blank on a commercial risk → Coverholder to supply the insured's trade or industry sector (mandatory for commercial risks under v5.2)
- Row 52 · HLB/2024/0127 · Risk Expiry Date: Expiry date 2024-12-15 is on or before inception 2025-12-14 → Coverholder to confirm the correct period of cover (source-record fact; not altered by this service)
- Row 56 · HLB/2024/0045 · Insured Trade or Industry Sector: Trade/industry sector blank on a commercial risk → Coverholder to supply the insured's trade or industry sector (mandatory for commercial risks under v5.2)
- Row 82 · HLB/2024/0093 · Risk Expiry Date: Expiry date 2024-12-09 is on or before inception 2025-12-08 → Coverholder to confirm the correct period of cover (source-record fact; not altered by this service)
- Row 95 · HLB/2024/0128 · Risk Code: Risk code 'ZZ' is not a Lloyd's risk code / not authorised for section 2 → Coverholder to confirm the correct Lloyd's risk code for this risk
- Row 96 · HLB/2024/0120 · Certificate Ref: Certificate reference is shared with row(s) 140 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 97 · HLB/2024/0104 · Certificate Ref: Certificate reference is shared with row(s) 8 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 106 · HLB/2024/0098 · Certificate Ref: Certificate reference is shared with row(s) 130 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 118 · HLB/2024/0122 · Insured Trade or Industry Sector: Trade/industry sector blank on a commercial risk → Coverholder to supply the insured's trade or industry sector (mandatory for commercial risks under v5.2)
- Row 126 · HLB/2024/0106 · Insured Trade or Industry Sector: Trade/industry sector blank on a commercial risk → Coverholder to supply the insured's trade or industry sector (mandatory for commercial risks under v5.2)
- Row 127 · HLB/2024/0022 · Insured Trade or Industry Sector: Trade/industry sector blank on a commercial risk → Coverholder to supply the insured's trade or industry sector (mandatory for commercial risks under v5.2)
- Row 130 · HLB/2024/0098 · Certificate Ref: Certificate reference is shared with row(s) 106 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 140 · HLB/2024/0120 · Certificate Ref: Certificate reference is shared with row(s) 96 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 152 · HLB/2024/0141 · Insured Trade or Industry Sector: Trade/industry sector blank on a commercial risk → Coverholder to supply the insured's trade or industry sector (mandatory for commercial risks under v5.2)
- Row 153 · HLB/2024/0080 · Original Currency: Original currency '—' cannot be determined from the source file → Coverholder to confirm the original currency of this transaction
- Row 162 · HLB/2024/0081 · Risk Expiry Date: Expiry date 2025-05-15 is on or before inception 2025-12-01 → Coverholder to confirm the correct period of cover (source-record fact; not altered by this service)
- Row 182 · HLB/2024/0053 · Certificate Ref: Certificate reference is shared with row(s) 5 on Original Premium transactions → Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other
- Row 188 · HLB/2024/0069 · Risk Expiry Date: Expiry date 2023-12-27 is on or before inception 2024-07-14 → Coverholder to confirm the correct period of cover (source-record fact; not altered by this service)

Full detail: validation-report.xlsx (Summary / Exceptions — Corrected / Exceptions — Client Action / Check Definitions).
