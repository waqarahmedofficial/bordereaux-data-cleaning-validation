# METHOD — Bordereaux Mapping, Standardisation & Validation Sample

**Service demonstrated:** a coverholder's messy premium bordereau is mapped to the
Lloyd's Coverholder Reporting Standards core data set, standardised, validated
against 14 named checks, corrected where a deterministic rule fixes the error,
and flagged — never silently guessed — where the coverholder's judgement is
required. Pure file-in / file-out. **Final submission into market systems
(e.g. the Delegated Data Manager) and all underwriting judgement remain with
the client** — this service delivers mapping, standardisation, validation and
an exception report only.

**Sample data:** 100% synthetic. The coverholder ("Harbourline Insurance
Services Ltd"), the binder (UMR B0876HLB250001X, Agreement BND-2024-1187,
incepted 1 Jul 2024, 18-month term, Year of Account 2024), all insured names,
certificate references and amounts were invented for this sample. No real
company's data is used.

---

## 1. Sources used (all public, read 24 Oct 2026)

1. **Lloyd's — Reporting Standards page** (lloyds.com/conducting-business/
   delegated-authorities/compliance-and-operations/reporting-standards):
   Lloyd's mandates that syndicates request and collect data consistent with
   the Coverholder Reporting Standards user guide; core requirements apply to
   binding authority / coverholder appointment agreements incepting since
   July 2017, covering regulatory, tax, premium and claims data.
2. **Lloyd's — Coverholder Reporting Standards User Guide, Version 5.2,
   20 August 2019** (announced by Market Bulletin Y5261), as the current
   version linked from the Lloyd's Reporting Standards page; and the
   Reporting Standards User Guide field tables (CiteSeerX copy of the Lloyd's
   guide), which publish the CR-coded premium field set used here
   (CR0001–CR0077) with mandatory/conditional status.
3. **Lloyd's — "Premium reporting: a standard for coverholders" User Guide**
   (28 Oct 2010; field definitions carried into the later standards),
   reproduced at manualzilla.com/doc/6890675: field-by-field definitions and
   conditions — UMR format ("All UMRs start with B, followed by the Lloyd's
   broker's number… only numbers 0-9 and letters A-Z; maximum length 17
   characters"); Coverholder PIN is a 9-digit number; Year of Account =
   "Lloyd's year of account in which the binding authority was incepted";
   original currency — ISO 4217 three-letter codes recommended; countries —
   ISO 3166 alpha-2 recommended; Net Premium to London = the premium paid to
   the London broker with all deductions/additions listed in the tax, levy,
   commission or accessori fields; rate of exchange convention (worked
   example: EUR original, GBP settlement, 1 EUR = 0.82 GBP → rate stated as
   1.22, i.e. original units per one settlement unit; settlement amount =
   original amount ÷ rate); the Lloyd's settlement currency list (AUD, CAD,
   DKK, EUR, HKD, NZD, NOK, GBP, ZAR, SGD, SEK, CHF, USD, JPY); transaction
   types (Original premium, Additional premium, Return premium, …); brokerage
   and final net premium are completed by the London broker.
4. **Lloyd's — Risk Codes: Guidance and Mappings, July 2021**
   (assets.lloyds.com): risk codes are alphanumeric Lloyd's codes; the
   sample binder's authorised codes are taken from it — B5 (binder
   commercial property outside the USA), B3 (same, USA), HP (UK household),
   B4 (binder private property outside the USA), W3 (UK employers'
   liability, per Market Bulletin Y4056).
5. **DDM Standard Operating Procedures (LIMOSS)**, premium bordereaux:
   field references CR0010, CR0035, CR0056 in operational use; file rules —
   XLSX recommended, no merged cells, static heading order, no duplicate
   headings; transformation is by a once-only mapping questionnaire.
6. **Pharos Production — "Bordereaux Reporting"** (pharosproduction.com/
   insights/engineering/bordereaux-reporting): field-level failure-mode
   analysis — UMR keyed by hand not matching the contract record; section
   left blank on a multi-section binder; year of account taken from the
   policy issue date instead of the agreement; commercial risk with no
   industry sector (mandatory under the v5.2 changes); net premium out of
   step with gross/commission/brokerage components (a consistency check);
   expiry before inception (no date-ordering rule is published in the
   standard — the ordering check is the pipeline's own, documented as such);
   exceptions routed back to the party that owns the source record.

## 2. The v5.2 premium core field set used (46 fields)

Header/contract group: Coverholder Name (CR0013, M), Coverholder PIN
(CR0014, C), Unique Market Reference (CR0005, M — at least one of UMR /
Agreement No), Agreement No (CR0006, C), Reporting Period Start Date
(CR0001, C), Reporting Period End Date (CR0002, M), Section No (CR0007, C),
Class of Business (CR0017, C), Risk Code (CR0016, C), Type of Insurance
(CR0019, M), London Broker Reference (CR0011, C), Year of Account (CR0010, C).

Transaction detail: Certificate Ref (CR0029, M), Insured First Name
(CR0034, C), Insured Full Name / Company Name (CR0035, M), Insured Trade or
Industry Sector (mandatory for commercial risks under the v5.2 changes — see
Limitations), Insured Country Subdivision (CR0039, C), Insured Country
(CR0041, M), Type of Business New or Renewal (CR0022, C), Risk Inception
Date (CR0030, M), Risk Expiry Date (CR0031, M), Sum Insured (CR0052, C),
Location of Risk Country (CR0050, M), Location of Risk Country Subdivision
(CR0048, C), Transaction Type (CR0056, M), Transaction Effective Date
(CR0057, C), Transaction Expiry Date (CR0058, C), Original Currency
(CR0020, M), Gross Annual Premium (CR0021, C), Gross Premium Paid This Time
(CR0059, M), Terrorism Premium (CR0060, C), Commission % (CR0061, M),
Commission Amount (CR0062, M), Total Taxes and Levies (CR0064, C), Tax Type
(tax block, C), Tax Jurisdiction (CR0077, C), Net Premium to London —
Original Currency (CR0065, M), Settlement Currency (M), Rate of Exchange
(CR0067, C), Net Premium to London — Settlement Currency (CR0068, C), and
the broker block reserved for the London broker: Brokerage % (C), Brokerage
Amount original currency (C), Final Net Premium original currency (C),
Brokerage Amount settlement currency (CR0074, C), Final Net Premium
settlement currency (CR0075, C), Final Net Premium % for Lloyd's (CR0076, C).

## 3. The 14 validation checks

CHK-01 UMR format & contract match · CHK-02 Agreement reference present ·
CHK-03 Reporting period complete · CHK-04 Section present (multi-section
binder) · CHK-05 Class of business present · CHK-06 Risk code present, valid
& authorised for the section · CHK-07 Year of account = binder YoA ·
CHK-08 Certificate reference unique (Original Premium rows; both rows of a
duplicated pair are flagged) · CHK-09 Insured name & industry-sector
completeness · CHK-10 Date logic (expiry after inception) · CHK-11 Currency
codes ISO 4217 / Lloyd's settlement list · CHK-12 Premium arithmetic
(commission = gross × %; net = gross − commission − taxes; tolerance 0.01) ·
CHK-13 FX completeness & arithmetic (rate present where currencies differ;
settlement = original ÷ rate, per the guide's convention; same-currency
settlement follows the corrected original net) · CHK-14 Mandatory-field
completeness sweep.

**Correction policy.** A value is corrected only where a published rule or
the contract record determines it uniquely (contract UMR, binder year of
account, section from class, risk code from the section's authorised set +
location, ISO code normalisation of unambiguous abbreviations, arithmetic
recomputation). Anything else — duplicate certificate references, missing
industry sector, impossible policy dates, a non-existent risk code, an
indeterminate currency — is carried unchanged, annotated in the clean file,
and listed for the client with the action required.

## 4. Sample arithmetic conventions (agreement parameters, not Lloyd's rules)

- Net Premium to London = Gross Premium Paid This Time − Commission Amount −
  Total Taxes and Levies, following the user guide's instruction that all
  deductions be listed in those fields. In this sample binder, taxes (UK
  Insurance Premium Tax at 12% on GB risks; nil on IE risks) are deductions.
  On a live engagement the sign treatment of taxes is a per-agreement
  parameter agreed with the managing agent.
- FX: EUR original → GBP settlement at rate 1.16 (original units per one
  settlement unit), the guide's stated convention.
- Commission rates (25.0% / 20.0% / 17.5% by section) and the IPT rate are
  sample binder data, not Lloyd's requirements.

## 5. Honest limitations

- The full v5.2 user guide PDF is published by Lloyd's for market
  participants; the field set here is assembled from Lloyd's own published
  field tables and guides (sources above), which are public. If a managing
  agent's own template differs in detail, the mapping template is adjusted
  per agreement — that is exactly what the reusable questionnaire is for.
- The CR code for **Insured Trade or Industry Sector** and for **Settlement
  Currency** could not be confirmed from the public extracts consulted; the
  fields are included (the sector field is evidenced as mandatory for
  commercial risks under the v5.2 changes per source 6) but no CR code is
  asserted for them. No Lloyd's requirement has been invented anywhere in
  this sample; where a check is the pipeline's own consistency rule rather
  than a published Lloyd's rule (CHK-10 date ordering, CHK-12/13 arithmetic
  reconciliation), the report says so.
- Conditional blocks not exercised by this sample binder: US surplus-lines
  block, accessori (Italy), instalment schedules. A binder writing those
  territories/classes would add the corresponding conditional fields.
- This sample validates a premium bordereau only. Risk and claims
  bordereaux share the standard's contract-identity fields but have their
  own detail fields.

## 6. Files & reproducibility

`build_raw.py` (generates the raw bordereau + error-ledger.csv),
`clean_and_validate.py` (mapping, standardisation, corrections, flags,
validation report, mapping template), `verify.py` (independent gate; writes
VERIFICATION.md). All scripts are idempotent with a fixed seed; re-running
the three in order reproduces every deliverable and the verdict.
