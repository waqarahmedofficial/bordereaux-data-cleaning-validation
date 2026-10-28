# VERIFICATION — Bordereaux Sample Gate

**Verdict: PASS**

## Counts

- Rows: raw 200 → clean 200
- Planted errors (error-ledger.csv): 74 total
- Planted correctable: 55 — verified fixed to the expected value: 55
- Planted flagged-for-client: 19 — verified on the Client Action list, annotated, values unaltered: 19
- Corrections recorded by the pipeline: 66 = 55 planted fixes + 11 consequential recalculations (each proven delta-for-delta)
- Client Action items recorded: 24 (the 19 planted items plus the counterpart rows of each duplicated certificate pair, which are flagged as well — a duplicated reference implicates both rows)
- Planted errors accounted for: 74 of 74 (100%)

## Independent checks

- PASS — row counts: raw=200, clean=200 (expected 200/200)
- PASS — row order and certificate references preserved from raw to clean
- PASS — all 55 planted correctable errors fixed to the expected value and recorded (verified 55)
- PASS — all 19 planted flagged errors are on the Client Action list, annotated in the clean file, and not silently altered (verified 19)
- PASS — all 11 non-ledger corrections are consequential recalculations proven delta-for-delta from a ledgered C7 correction on the same row
- PASS — no unexplained numeric changes between raw and clean (found 0: [])
- PASS — clean file: Year of Account = 2024 on all 200 rows
- PASS — YoA sweep: raw rows with wrong YoA are exactly the C1 ledger rows
- PASS — clean file: UMR = contract UMR on all 200 rows
- PASS — UMR sweep: raw rows with malformed UMR are exactly the C4 ledger rows
- PASS — duplicate sweep: every row of every duplicated Original Premium certificate pair is flagged (10 rows across 5 pairs; 5 planted + 5 counterpart rows)
- PASS — date sweep: rows with expiry <= inception in the clean file are exactly the C10 flagged rows (carried, not altered)
- PASS — sector sweep: commercial rows missing sector in the clean file are exactly the C8 flagged rows
- PASS — currency sweep: the only blank original currency in the clean file is the flagged indeterminate row
- PASS — risk-code sweep: every clean risk code is authorised for its section, except the flagged 'ZZ' row
- PASS — gross totals reconcile exactly: raw 2,450,433.56 = clean 2,450,433.56
- PASS — net premium changes occur on exactly the 7 C7 commission/net rows and nowhere else
- PASS — settlement amounts change on exactly the 9 C7 rows and nowhere else
- PASS — settlement totals identity holds: clean 1,567,311.16 = keyed 1,586,187.25 - flagged-currency row held out 18,480.98 + C7 deltas -395.11
- INFO — totals: gross raw 2,450,433.56 = clean 2,450,433.56; net raw 1,625,328.40 -> clean 1,624,889.15 (delta -439.25, on C7 rows only); settlement raw 1,586,187.25 -> clean 1,567,311.16 (C7 deltas -395.11; the flagged-currency row's keyed settlement 18,480.98 is held out of the clean settlement total pending confirmation)

## Method of verification
verify.py re-reads raw-bordereau.csv, error-ledger.csv, clean-bordereau-v52.csv and the two machine-readable exception lists. It does not import or trust the cleaning pipeline's logic: expected values come from the ledger written at generation time, consequential corrections are proven arithmetically, and structural sweeps compare the raw and clean files field by field.
