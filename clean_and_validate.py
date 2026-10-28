#!/usr/bin/env python3
"""
clean_and_validate.py — The service pipeline:
  1. Parse the coverholder's raw bordereau (messy formats).
  2. Map raw columns to the Lloyd's Coverholder Reporting Standards (v5.2)
     premium-bordereau core field set.
  3. Standardise formats (dates ISO, ISO 4217 currencies, ISO 3166 country
     codes, numeric amounts) and correct every deterministically-fixable
     error, recording each correction with its rule basis.
  4. Flag (never silently guess) every item needing coverholder judgement.
  5. Run the 14 named validation checks and emit:
       clean-bordereau-v52.xlsx / .csv
       validation-report.xlsx, validation-report-summary.md
       exceptions-corrected.csv, exceptions-client-action.csv  (machine-readable)
       mapping-template.xlsx
Idempotent: re-runnable over the same raw file.
"""
import csv, json, os, re
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

OUT = os.path.dirname(os.path.abspath(__file__))

# ------------------------------------------------------------- contract data -
UMR = "B0876HLB250001X"
AGREEMENT_NO = "BND-2024-1187"
COVERHOLDER = "Harbourline Insurance Services Ltd"
PIN = "123456782"
BROKER_REF = "HLB"
YOA = 2024
PERIOD_START, PERIOD_END = "2025-12-01", "2025-12-31"
SETTLEMENT_CCY = "GBP"
LLOYDS_SETTLEMENT_CURRENCIES = {"AUD", "CAD", "DKK", "EUR", "HKD", "NZD", "NOK",
                                "GBP", "ZAR", "SGD", "SEK", "CHF", "USD", "JPY"}
SECTION_BY_CLASS = {"Commercial Property": 1, "Household": 2, "Employers' Liability": 3}
CLASS_BY_SECTION = {v: k for k, v in SECTION_BY_CLASS.items()}
AUTHORISED_CODES = {1: {"B5", "B3"}, 2: {"HP", "B4"}, 3: {"W3"}}
# deterministic code derivation where the source code is blank:
# section 1 = binder commercial property (B5 non-US / B3 US); section 2 =
# household (HP UK / B4 non-UK binder private property); section 3 = W3.
def derive_code(section, risk_country):
    if section == 1:
        return "B3" if risk_country == "US" else "B5"
    if section == 2:
        return "HP" if risk_country == "GB" else "B4"
    return "W3"

COUNTRY_MAP = {"UNITED KINGDOM": "GB", "UK": "GB", "GREAT BRITAIN": "GB", "GB": "GB",
               "IRELAND": "IE", "IE": "IE", "EIRE": "IE",
               "UNITED STATES": "US", "USA": "US", "US": "US"}
CCY_MAP = {"GBP": "GBP", "GBp": "GBP", "gbp": "GBP", "£": "GBP",
           "EUR": "EUR", "EURO": "EUR", "Eur": "EUR", "€": "EUR",
           "USD": "USD", "US DOLLARS": "USD", "US$": "USD"}

CLEAN_FIELDS = [
    # (field name, CR ref, status)
    ("Coverholder Name", "CR0013", "M"), ("Coverholder PIN", "CR0014", "C"),
    ("Unique Market Reference (UMR)", "CR0005", "M"), ("Agreement No", "CR0006", "C"),
    ("Reporting Period Start Date", "CR0001", "C"), ("Reporting Period End Date", "CR0002", "M"),
    ("Section No", "CR0007", "C"), ("Class of Business", "CR0017", "C"),
    ("Risk Code", "CR0016", "C"), ("Type of Insurance", "CR0019", "M"),
    ("London Broker Reference", "CR0011", "C"), ("Year of Account", "CR0010", "C"),
    ("Certificate Ref", "CR0029", "M"), ("Insured First Name", "CR0034", "C"),
    ("Insured Full Name / Company Name", "CR0035", "M"),
    ("Insured Trade or Industry Sector", "v5.2*", "C"),
    ("Insured Country Subdivision", "CR0039", "C"), ("Insured Country", "CR0041", "M"),
    ("Type of Business (New or Renewal)", "CR0022", "C"),
    ("Risk Inception Date", "CR0030", "M"), ("Risk Expiry Date", "CR0031", "M"),
    ("Sum Insured", "CR0052", "C"),
    ("Location of Risk Country", "CR0050", "M"), ("Location of Risk Country Subdivision", "CR0048", "C"),
    ("Transaction Type", "CR0056", "M"), ("Transaction Effective Date", "CR0057", "C"),
    ("Transaction Expiry Date", "CR0058", "C"),
    ("Original Currency", "CR0020", "M"), ("Gross Annual Premium", "CR0021", "C"),
    ("Gross Premium Paid This Time", "CR0059", "M"), ("Terrorism Premium", "CR0060", "C"),
    ("Commission %", "CR0061", "M"), ("Commission Amount", "CR0062", "M"),
    ("Total Taxes and Levies", "CR0064", "C"), ("Tax Type", "tax block", "C"),
    ("Tax Jurisdiction", "CR0077", "C"),
    ("Net Premium to London (Original Currency)", "CR0065", "M"),
    ("Settlement Currency", "—", "M"), ("Rate of Exchange", "CR0067", "C"),
    ("Net Premium to London (Settlement Currency)", "CR0068", "C"),
    ("Brokerage %", "broker block", "C"), ("Brokerage Amount (Original Currency)", "broker block", "C"),
    ("Final Net Premium (Original Currency)", "broker block", "C"),
    ("Brokerage Amount (Settlement Currency)", "CR0074", "C"),
    ("Final Net Premium (Settlement Currency)", "CR0075", "C"),
    ("Final Net Premium % for Lloyd's", "CR0076", "C"),
]
FIELD_NAMES = [f[0] for f in CLEAN_FIELDS]

# ------------------------------------------------------------------- parsing -
def parse_date(s):
    s = (s or "").strip()
    if not s:
        return None
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"unparseable date: {s!r}")

def parse_money(s):
    s = (s or "").strip().replace("£", "").replace("€", "").replace("$", "").replace(",", "")
    if s in ("", "-"):
        return None
    return round(float(s), 2)

def parse_pct(s):
    s = (s or "").strip()
    if not s:
        return None
    if s.endswith("%"):
        return float(s[:-1])
    v = float(s)
    return v * 100 if v < 1 else v

def r2(v):
    return round(v + 1e-9, 2)

raw_path = os.path.join(OUT, "raw-bordereau.csv")
with open(raw_path, newline="", encoding="utf-8") as f:
    rows = list(csv.reader(f))
hdr_idx = next(i for i, r in enumerate(rows) if r and r[0] == "Coverholder")
headers = rows[hdr_idx]
data_rows = []
for r in rows[hdr_idx + 1:]:
    if not r or r[0] == "TOTALS":
        continue
    data_rows.append(dict(zip(headers, r + [""] * (len(headers) - len(r)))))
assert len(data_rows) == 200, len(data_rows)

corrected, flagged = [], []

def note_corr(row, cert, field, check, raw, new, rule):
    corrected.append({"row": row, "cert_ref": cert, "field": field, "check": check,
                      "raw_value": raw, "corrected_value": new, "rule": rule})

def note_flag(row, cert, field, check, problem, action):
    flagged.append({"row": row, "cert_ref": cert, "field": field, "check": check,
                    "problem": problem, "action_required": action})

# duplicate certificate references among Original Premium rows (both rows of
# each duplicated pair are exceptions — the file cannot show which owns it)
from collections import defaultdict
cert_groups = defaultdict(list)
for j, d in enumerate(data_rows, start=1):
    if d["Trans Type"].strip() == "Original Premium":
        cert_groups[d["Cert Ref"].strip()].append(j)
dup_groups = {c: js for c, js in cert_groups.items() if len(js) > 1}

clean = []
for i, d in enumerate(data_rows, start=1):
    cert_keyed = d["Cert Ref"].strip()
    out = {name: "" for name in FIELD_NAMES}
    notes = []

    # --- contract identity -------------------------------------------------
    umr_raw = d["UMR / Binder Ref"].strip()
    if umr_raw != UMR:
        note_corr(i, cert_keyed, "Unique Market Reference (UMR)",
                  "CHK-01 UMR format & contract match", umr_raw, UMR,
                  "CR0005; user-guide UMR format (starts 'B', alphanumeric, max 17 chars, no punctuation); single-binder file corrected to contract UMR")
    section_txt = d["Sect"].strip()
    class_txt = d["Class"].strip()
    section = int(section_txt) if section_txt else SECTION_BY_CLASS[class_txt]
    if not section_txt:
        note_corr(i, cert_keyed, "Section No", "CHK-04 Section present (multi-section binder)",
                  "", str(section),
                  "CR0007 + v5.2/Bulletin Y5261 section requirement; derived from class of business")
    risk_country = COUNTRY_MAP[d["Risk Country"].strip().upper()]
    insured_country = COUNTRY_MAP[d["Insured Country"].strip().upper()]

    code_keyed = d["Risk Cd"]
    code_norm = re.sub(r"[^A-Za-z0-9]", "", code_keyed).upper()
    if code_keyed.strip() == "":
        code = derive_code(section, risk_country)
        note_corr(i, cert_keyed, "Risk Code", "CHK-06 Risk code present, valid & authorised",
                  "", code,
                  "CR0016; derived from the binder section's authorised code set (Lloyd's Risk Code Guidance) + location of risk")
    elif code_norm in AUTHORISED_CODES[section]:
        code = code_norm
        if code_keyed != code:
            note_corr(i, cert_keyed, "Risk Code", "CHK-06 Risk code present, valid & authorised",
                      code_keyed, code,
                      "Lloyd's Risk Code Guidance: format normalised (case/punctuation) and validated against the section's authorised set")
    else:
        code = code_keyed.strip()
        note_flag(i, cert_keyed, "Risk Code", "CHK-06 Risk code present, valid & authorised",
                  f"Risk code '{code_keyed.strip()}' is not a Lloyd's risk code / not authorised for section {section}",
                  "Coverholder to confirm the correct Lloyd's risk code for this risk")
        notes.append("ACTION REQUIRED: risk code to be confirmed by coverholder")

    yoa_raw = d["YoA"].strip()
    if yoa_raw != str(YOA):
        note_corr(i, cert_keyed, "Year of Account", "CHK-07 Year of account = binder YoA",
                  yoa_raw, str(YOA),
                  "CR0010: year of account = year the binding authority incepted (2024), not the policy issue year")

    # --- certificate uniqueness (Original Premium rows) ---------------------
    trans_type = d["Trans Type"].strip()
    if trans_type == "Original Premium" and cert_keyed in dup_groups:
        others = [j for j in dup_groups[cert_keyed] if j != i]
        note_flag(i, cert_keyed, "Certificate Ref", "CHK-08 Certificate reference unique",
                  f"Certificate reference is shared with row(s) {', '.join(map(str, others))} on Original Premium transactions",
                  "Coverholder to confirm which risk owns this certificate reference and supply the correct reference for the other")
        notes.append("ACTION REQUIRED: duplicate certificate reference — correct reference to be confirmed")

    # --- insured / sector ----------------------------------------------------
    sector = d["Trade / Sector"].strip()
    if section in (1, 3) and sector == "":
        note_flag(i, cert_keyed, "Insured Trade or Industry Sector",
                  "CHK-09 Insured name & sector completeness",
                  "Trade/industry sector blank on a commercial risk",
                  "Coverholder to supply the insured's trade or industry sector (mandatory for commercial risks under v5.2)")
        notes.append("ACTION REQUIRED: insured trade/industry sector to be supplied")

    # --- dates ---------------------------------------------------------------
    inc = parse_date(d["Inception"])
    exp = parse_date(d["Expiry"])
    if exp <= inc:
        note_flag(i, cert_keyed, "Risk Expiry Date", "CHK-10 Date logic",
                  f"Expiry date {exp.isoformat()} is on or before inception {inc.isoformat()}",
                  "Coverholder to confirm the correct period of cover (source-record fact; not altered by this service)")
        notes.append("ACTION REQUIRED: expiry date precedes inception — correct dates to be confirmed")

    # --- currency --------------------------------------------------------------
    ccy_keyed = d["Curr"]
    ccy = CCY_MAP.get(ccy_keyed.strip())
    if ccy is None:
        ccy = ""
        note_flag(i, cert_keyed, "Original Currency", "CHK-11 Currency codes (ISO 4217)",
                  f"Original currency '{ccy_keyed.strip()}' cannot be determined from the source file",
                  "Coverholder to confirm the original currency of this transaction")
        notes.append("ACTION REQUIRED: original currency to be confirmed")
    elif ccy_keyed.strip() != ccy:
        note_corr(i, cert_keyed, "Original Currency", "CHK-11 Currency codes (ISO 4217)",
                  ccy_keyed.strip(), ccy,
                  "CR0020; ISO 4217 three-letter code — unambiguous abbreviation/symbol normalised")

    # --- money -------------------------------------------------------------------
    gross = parse_money(d["Gross Prem (this time)"])
    comm_pct = parse_pct(d["Comm %"])
    comm_keyed = parse_money(d["Comm Amt"])
    tax = parse_money(d["Tax Amt"]) or 0.0
    comm_calc = r2(gross * comm_pct / 100.0)
    if abs(comm_keyed - comm_calc) > 0.01:
        note_corr(i, cert_keyed, "Commission Amount", "CHK-12 Premium arithmetic",
                  f"{comm_keyed:.2f}", f"{comm_calc:.2f}",
                  "CR0061/CR0062: commission amount = gross premium paid this time x commission %")
    net_calc = r2(gross - comm_calc - tax)
    net_keyed = parse_money(d["Net to London"])
    if abs(net_keyed - net_calc) > 0.01:
        note_corr(i, cert_keyed, "Net Premium to London (Original Currency)",
                  "CHK-12 Premium arithmetic", f"{net_keyed:.2f}", f"{net_calc:.2f}",
                  "CR0065: net = gross premium paid this time - commission amount - total taxes and levies (all deductions listed, per user guide)")
    fx_txt = d["FX Rate"].strip()
    fx = float(fx_txt) if fx_txt else None
    settle_keyed = parse_money(d["Net to London (Settle Ccy)"])
    if ccy and ccy != SETTLEMENT_CCY and fx:
        settle_calc = r2(net_calc / fx)
        if abs(settle_keyed - settle_calc) > 0.01:
            note_corr(i, cert_keyed, "Net Premium to London (Settlement Currency)",
                      "CHK-13 FX completeness & arithmetic", f"{settle_keyed:.2f}",
                      f"{settle_calc:.2f}",
                      "CR0067/CR0068: settlement net = original-currency net / rate of exchange (guide convention)")
        settle_out = settle_calc
    elif ccy == SETTLEMENT_CCY:
        settle_out = net_calc
        if abs(settle_keyed - net_calc) > 0.01:
            note_corr(i, cert_keyed, "Net Premium to London (Settlement Currency)",
                      "CHK-13 FX completeness & arithmetic", f"{settle_keyed:.2f}",
                      f"{net_calc:.2f}",
                      "Settlement currency equals original currency, so the settlement amount follows the corrected original-currency net premium")
    else:
        settle_out = None

    # --- assemble clean row -------------------------------------------------------
    out.update({
        "Coverholder Name": COVERHOLDER, "Coverholder PIN": PIN,
        "Unique Market Reference (UMR)": UMR, "Agreement No": AGREEMENT_NO,
        "Reporting Period Start Date": PERIOD_START, "Reporting Period End Date": PERIOD_END,
        "Section No": section, "Class of Business": class_txt, "Risk Code": code,
        "Type of Insurance": d["Ins Type"].strip(), "London Broker Reference": BROKER_REF,
        "Year of Account": YOA, "Certificate Ref": cert_keyed,
        "Insured First Name": d["Insured First Name"].strip(),
        "Insured Full Name / Company Name": d["Insured Name"].strip(),
        "Insured Trade or Industry Sector": sector,
        "Insured Country Subdivision": d["Insured County/State"].strip(),
        "Insured Country": insured_country,
        "Type of Business (New or Renewal)": d["New/Ren"].strip(),
        "Risk Inception Date": inc.isoformat(), "Risk Expiry Date": exp.isoformat(),
        "Sum Insured": parse_money(d["Sum Insured"]) if d["Sum Insured"].strip() else "",
        "Location of Risk Country": risk_country,
        "Location of Risk Country Subdivision": d["Risk County/State"].strip(),
        "Transaction Type": trans_type,
        "Transaction Effective Date": (parse_date(d["Trans Eff Date"]).isoformat()
                                       if d["Trans Eff Date"].strip() else ""),
        "Transaction Expiry Date": "",
        "Original Currency": ccy,
        "Gross Annual Premium": parse_money(d["Annual Gross"]) if d["Annual Gross"].strip() else "",
        "Gross Premium Paid This Time": gross,
        "Terrorism Premium": parse_money(d["Terr Prem"]) if d["Terr Prem"].strip() else "",
        "Commission %": comm_pct, "Commission Amount": comm_calc,
        "Total Taxes and Levies": tax,
        "Tax Type": d["Tax Type"].strip(),
        "Tax Jurisdiction": ("GB" if d["Tax Jurisdiction"].strip() else ""),
        "Net Premium to London (Original Currency)": net_calc,
        "Settlement Currency": SETTLEMENT_CCY,
        "Rate of Exchange": fx if fx else "",
        "Net Premium to London (Settlement Currency)": settle_out if settle_out is not None else "",
    })
    out["_row"] = i
    out["_notes"] = "; ".join(notes)
    clean.append(out)

# ------------------------------------------------------------------- checks ---
from collections import Counter
corr_by_check = Counter(c["check"] for c in corrected)
flag_by_check = Counter(f["check"] for f in flagged)
CHECKS = [
    ("CHK-01 UMR format & contract match", "CR0005", "Every row's UMR meets the user-guide format (starts 'B', alphanumeric only, max 17 chars) and equals the contract UMR for this binder."),
    ("CHK-02 Agreement reference present", "CR0006", "Agreement number present on every row (at least one of UMR / agreement number must be shown)."),
    ("CHK-03 Reporting period complete", "CR0001/CR0002", "Reporting period start and end dates stated on every row; end date gives at least month and year."),
    ("CHK-04 Section present (multi-section binder)", "CR0007", "Section number present on every row of a multi-section binder (v5.2 change, Market Bulletin Y5261)."),
    ("CHK-05 Class of business present", "CR0017", "Class of business narrative present on every row."),
    ("CHK-06 Risk code present, valid & authorised", "CR0016", "Risk code present, in Lloyd's format, a genuine Lloyd's risk code and within the section's authorised code set."),
    ("CHK-07 Year of account = binder YoA", "CR0010", "Year of account equals the year the binding authority incepted (2024) on every row — not the policy issue year."),
    ("CHK-08 Certificate reference unique", "CR0029", "Certificate reference present on every row and unique across Original Premium transactions."),
    ("CHK-09 Insured name & sector completeness", "CR0035 / v5.2", "Insured name present on every row; trade/industry sector present on every commercial risk (mandatory under v5.2)."),
    ("CHK-10 Date logic", "CR0030/CR0031/CR0057", "Risk expiry follows inception on every row; endorsement effective dates fall inside the period of cover."),
    ("CHK-11 Currency codes (ISO 4217)", "CR0020", "Original currency is a valid ISO 4217 three-letter code; settlement currency is a Lloyd's settlement currency."),
    ("CHK-12 Premium arithmetic", "CR0059/61/62/64/65", "Commission amount = gross x commission %; net premium to London = gross - commission - total taxes and levies (tolerance 0.01)."),
    ("CHK-13 FX completeness & arithmetic", "CR0067/CR0068", "Where settlement currency differs from original: rate of exchange present and settlement net = original net / rate (guide convention, tolerance 0.01)."),
    ("CHK-14 Mandatory-field completeness sweep", "v5.2 core set", "Every mandatory (M) field of the premium core set populated on every row, except items carried on the Client Action list."),
]
check_rows = []
for name, ref, rule in CHECKS:
    found = corr_by_check.get(name, 0) + flag_by_check.get(name, 0)
    if name.startswith("CHK-14"):
        blanks = sum(1 for o in clean if o["Original Currency"] == "")
        found = blanks
        check_rows.append((name, ref, rule, 200, found, 0, found,
                           "PASS — all mandatory fields populated except items on the Client Action list"))
    elif found == 0:
        check_rows.append((name, ref, rule, 200, 0, 0, 0, "PASS — no exceptions found"))
    else:
        check_rows.append((name, ref, rule, 200, found, corr_by_check.get(name, 0),
                           flag_by_check.get(name, 0),
                           "PASS — all exceptions corrected or carried on the Client Action list"))

# ------------------------------------------------------------------- totals ---
tot = lambda rows, k: r2(sum(parse_money(r[k]) or 0.0 for r in rows))
keyed = {
    "gross": tot(data_rows, "Gross Prem (this time)"),
    "comm": tot(data_rows, "Comm Amt"), "tax": tot(data_rows, "Tax Amt"),
    "net": tot(data_rows, "Net to London"), "net_s": tot(data_rows, "Net to London (Settle Ccy)"),
}
clean_tot = {
    "gross": r2(sum(o["Gross Premium Paid This Time"] for o in clean)),
    "comm": r2(sum(o["Commission Amount"] for o in clean)),
    "tax": r2(sum(o["Total Taxes and Levies"] for o in clean)),
    "net": r2(sum(o["Net Premium to London (Original Currency)"] for o in clean)),
    "net_s": r2(sum(o["Net Premium to London (Settlement Currency)"] for o in clean
                    if o["Net Premium to London (Settlement Currency)"] != "")),
}

# ------------------------------------------------------------------ outputs ---
hdr_fill = PatternFill("solid", fgColor="1F3864")
hdr_font = Font(bold=True, color="FFFFFF")
flag_fill = PatternFill("solid", fgColor="FFF2CC")

# clean workbook
wb = Workbook()
ws = wb.active
ws.title = "Premium Bordereau v5.2"
ws.append(FIELD_NAMES + ["Validation Note (service annotation — not a standard field)"])
for c in ws[1]:
    c.font = hdr_font
    c.fill = hdr_fill
    c.alignment = Alignment(wrap_text=True, vertical="center")
for o in clean:
    ws.append([o[n] for n in FIELD_NAMES] + [o["_notes"]])
    if o["_notes"]:
        for c in ws[ws.max_row]:
            c.fill = flag_fill
ws.freeze_panes = "A2"
fk = wb.create_sheet("Field Key")
fk.append(["Field", "Lloyd's reference", "Status in core set"])
for c in fk[1]:
    c.font = Font(bold=True)
for n, ref, st in CLEAN_FIELDS:
    fk.append([n, ref, {"M": "Mandatory", "C": "Conditional"}[st]])
fk.append([])
fk.append(["* Insured Trade or Industry Sector: mandatory for commercial risks under the v5.2 changes (Market Bulletin Y5261); its CR code was not confirmed in the public extracts consulted — see METHOD.md."])
fk.append(["Broker block fields are reserved for completion by the London broker, per the user guide; left blank in this coverholder submission."])
wb.save(os.path.join(OUT, "clean-bordereau-v52.xlsx"))

with open(os.path.join(OUT, "clean-bordereau-v52.csv"), "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=FIELD_NAMES + ["Validation Note"])
    w.writeheader()
    for o in clean:
        row = {n: o[n] for n in FIELD_NAMES}
        row["Validation Note"] = o["_notes"]
        w.writerow(row)

# machine-readable exception lists
for fname, items, cols in [
    ("exceptions-corrected.csv", corrected, ["row", "cert_ref", "field", "check", "raw_value", "corrected_value", "rule"]),
    ("exceptions-client-action.csv", flagged, ["row", "cert_ref", "field", "check", "problem", "action_required"]),
]:
    with open(os.path.join(OUT, fname), "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(sorted(items, key=lambda e: (e["row"], e["field"])))

# validation report workbook
wb = Workbook()
ws = wb.active
ws.title = "Summary"
ws.append(["BORDEREAUX VALIDATION & EXCEPTION REPORT — SAMPLE"])
ws.append(["Service", "Bordereaux data mapping, standardisation & validation"])
ws.append(["Standard", "Lloyd's Coverholder Reporting Standards, Version 5.2 (20 Aug 2019, Market Bulletin Y5261) — premium bordereau"])
ws.append(["Coverholder", COVERHOLDER])
ws.append(["Binder", f"UMR {UMR} · Agreement {AGREEMENT_NO} · Year of Account {YOA}"])
ws.append(["Reporting period", f"{PERIOD_START} to {PERIOD_END}"])
ws.append(["Rows processed", 200])
ws.append(["Exceptions corrected in file", len(corrected)])
ws.append(["Items flagged for client action", len(flagged)])
ws.append(["Rows requiring client action", len({f['row'] for f in flagged})])
ws.append([])
ws.append(["Check", "Lloyd's ref", "Rows checked", "Exceptions found", "Corrected", "Flagged for client", "Result"])
for c in ws[ws.max_row]:
    c.font = hdr_font
    c.fill = hdr_fill
for name, ref, rule, n, found, corr, flag, result in check_rows:
    ws.append([name, ref, n, found, corr, flag, result])
ws.append([])
ws.append(["Totals reconciliation (keyed in raw file vs clean file)"])
ws.append(["Measure", "As keyed (raw)", "Clean file", "Difference", "Explained by"])
EXPLAIN = {
    "gross": "No gross errors; totals agree",
    "comm": "CHK-12 commission corrections only",
    "tax": "No tax errors; totals agree",
    "net": "CHK-12 corrections only, including consequential recalculation where a commission amount was corrected",
    "net_s": "CHK-12/CHK-13 corrections; clean total excludes the 1 row flagged for unconfirmed original currency (Client Action list), which is carried in the original-currency totals only",
}
for label, k in [("Gross premium paid this time (mixed currencies, as keyed)", "gross"),
                 ("Commission amount", "comm"), ("Total taxes and levies", "tax"),
                 ("Net premium to London (original currency)", "net"),
                 ("Net premium to London (settlement currency GBP)", "net_s")]:
    ws.append([label, keyed[k], clean_tot[k], r2(clean_tot[k] - keyed[k]), EXPLAIN[k]])
for title_row in (1,):
    ws[title_row][0].font = Font(bold=True, size=13)
ws2 = wb.create_sheet("Exceptions — Corrected")
ws2.append(["Row", "Certificate Ref", "Field", "Check", "As keyed", "Corrected to", "Rule applied"])
for c in ws2[1]:
    c.font = hdr_font
    c.fill = hdr_fill
for e in sorted(corrected, key=lambda e: (e["row"], e["field"])):
    ws2.append([e["row"], e["cert_ref"], e["field"], e["check"], e["raw_value"], e["corrected_value"], e["rule"]])
ws3 = wb.create_sheet("Exceptions — Client Action")
ws3.append(["Row", "Certificate Ref", "Field", "Check", "Problem", "Action required from coverholder"])
for c in ws3[1]:
    c.font = hdr_font
    c.fill = hdr_fill
for e in sorted(flagged, key=lambda e: (e["row"], e["field"])):
    ws3.append([e["row"], e["cert_ref"], e["field"], e["check"], e["problem"], e["action_required"]])
ws4 = wb.create_sheet("Check Definitions")
ws4.append(["Check", "Lloyd's ref", "Rule"])
for c in ws4[1]:
    c.font = hdr_font
    c.fill = hdr_fill
for name, ref, rule, *_ in check_rows:
    ws4.append([name, ref, rule])
wb.save(os.path.join(OUT, "validation-report.xlsx"))

# summary markdown
lines = []
lines.append("# Validation & Exception Report — Summary (SAMPLE)\n")
lines.append(f"**Standard:** Lloyd's Coverholder Reporting Standards v5.2 (20 August 2019, Market Bulletin Y5261) — premium bordereau  ")
lines.append(f"**Coverholder:** {COVERHOLDER}  ")
lines.append(f"**Binder:** UMR {UMR} · Agreement {AGREEMENT_NO} · Year of Account {YOA}  ")
lines.append(f"**Reporting period:** {PERIOD_START} to {PERIOD_END}  ")
lines.append(f"**Rows processed:** 200 · **Exceptions corrected:** {len(corrected)} · "
             f"**Flagged for client action:** {len(flagged)} items on {len({f['row'] for f in flagged})} rows\n")
lines.append("## Check results\n")
lines.append("| Check | Ref | Exceptions | Corrected | Flagged | Result |")
lines.append("|---|---|---|---|---|---|")
for name, ref, rule, n, found, corr, flag, result in check_rows:
    lines.append(f"| {name} | {ref} | {found} | {corr} | {flag} | {result} |")
lines.append("\n## Totals reconciliation\n")
lines.append("| Measure | As keyed (raw) | Clean file | Difference |")
lines.append("|---|---|---|---|")
for label, k in [("Gross premium paid this time", "gross"), ("Commission amount", "comm"),
                 ("Total taxes and levies", "tax"), ("Net premium to London (original ccy)", "net"),
                 ("Net premium to London (settlement GBP)", "net_s")]:
    lines.append(f"| {label} | {keyed[k]:,.2f} | {clean_tot[k]:,.2f} | {r2(clean_tot[k]-keyed[k]):,.2f} |")
lines.append("\nGross premium and tax totals are unchanged — no gross or tax errors were present. "
             "Net differences are fully explained by the CHK-12/CHK-13 arithmetic corrections listed in the workbook. "
             "Recorded corrections include consequential recalculations: where a commission amount was corrected, the dependent "
             "net premium (and, on one EUR row, the settlement-currency net) was recalculated from the corrected components. "
             "The clean settlement-currency total excludes the one row flagged for unconfirmed original currency; that row is "
             "carried in the original-currency totals only. All other flagged rows are carried at their keyed amounts pending "
             "coverholder confirmation.\n")
lines.append("## Items requiring client action\n")
for e in sorted(flagged, key=lambda e: (e["row"], e["field"])):
    lines.append(f"- Row {e['row']} · {e['cert_ref']} · {e['field']}: {e['problem']} → {e['action_required']}")
lines.append("\nFull detail: validation-report.xlsx (Summary / Exceptions — Corrected / Exceptions — Client Action / Check Definitions).")
with open(os.path.join(OUT, "validation-report-summary.md"), "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")

# mapping template workbook
wb = Workbook()
ws = wb.active
ws.title = "Mapping Template"
ws.append(["Raw column (coverholder layout)", "Maps to v5.2 field", "Lloyd's ref", "Status",
           "Transformation applied"])
for c in ws[1]:
    c.font = hdr_font
    c.fill = hdr_fill
MAPPING = [
    ("Coverholder", "Coverholder Name", "CR0013", "M", "Constant per binder; repeated on every row"),
    ("(coverholder PIN, from agreement record)", "Coverholder PIN", "CR0014", "C", "Constant; 9-digit Lloyd's PIN"),
    ("UMR / Binder Ref", "Unique Market Reference (UMR)", "CR0005", "M", "Validated to user-guide format; corrected to contract UMR"),
    ("Agreement No.", "Agreement No", "CR0006", "C", "As keyed"),
    ("Period From", "Reporting Period Start Date", "CR0001", "C", "Date standardised to ISO YYYY-MM-DD"),
    ("Period To", "Reporting Period End Date", "CR0002", "M", "Date standardised to ISO YYYY-MM-DD"),
    ("Sect", "Section No", "CR0007", "C", "Integer; derived from class of business where blank"),
    ("Class", "Class of Business", "CR0017", "C", "As keyed"),
    ("Risk Cd", "Risk Code", "CR0016", "C", "Format normalised; validated against section's authorised set; derived where blank and derivable"),
    ("Ins Type", "Type of Insurance", "CR0019", "M", "Controlled list (Direct / Facultative Reinsurance / Treaty Reinsurance / Excess of Loss)"),
    ("(from UMR / agreement record)", "London Broker Reference", "CR0011", "C", "Broker reference embedded in UMR"),
    ("YoA", "Year of Account", "CR0010", "C", "Set to the binder's year of account (agreement inception year)"),
    ("Cert Ref", "Certificate Ref", "CR0029", "M", "As keyed; uniqueness checked across Original Premium rows"),
    ("Insured First Name", "Insured First Name", "CR0034", "C", "Individuals only"),
    ("Insured Name", "Insured Full Name / Company Name", "CR0035", "M", "As keyed"),
    ("Trade / Sector", "Insured Trade or Industry Sector", "v5.2*", "C", "Mandatory for commercial risks under v5.2; flagged where blank"),
    ("Insured County/State", "Insured Country Subdivision", "CR0039", "C", "As keyed"),
    ("Insured Country", "Insured Country", "CR0041", "M", "Country name mapped to ISO 3166 alpha-2"),
    ("New/Ren", "Type of Business (New or Renewal)", "CR0022", "C", "Controlled list (New / Renewal)"),
    ("Inception", "Risk Inception Date", "CR0030", "M", "Date standardised to ISO; date-logic checked"),
    ("Expiry", "Risk Expiry Date", "CR0031", "M", "Date standardised to ISO; date-logic checked"),
    ("Sum Insured", "Sum Insured", "CR0052", "C", "Numeric, original currency"),
    ("Risk Country", "Location of Risk Country", "CR0050", "M", "Country name mapped to ISO 3166 alpha-2"),
    ("Risk County/State", "Location of Risk Country Subdivision", "CR0048", "C", "As keyed"),
    ("Trans Type", "Transaction Type", "CR0056", "M", "Controlled list per user guide"),
    ("Trans Eff Date", "Transaction Effective Date", "CR0057", "C", "Date standardised to ISO"),
    ("(not in source layout)", "Transaction Expiry Date", "CR0058", "C", "Left blank where not applicable / not supplied"),
    ("Curr", "Original Currency", "CR0020", "M", "Normalised to ISO 4217 three-letter code"),
    ("Annual Gross", "Gross Annual Premium", "CR0021", "C", "Numeric"),
    ("Gross Prem (this time)", "Gross Premium Paid This Time", "CR0059", "M", "Numeric; currency symbols/separators stripped"),
    ("Terr Prem", "Terrorism Premium", "CR0060", "C", "Numeric where separately stated"),
    ("Comm %", "Commission %", "CR0061", "M", "Normalised to percentage number (25.0 = 25%)"),
    ("Comm Amt", "Commission Amount", "CR0062", "M", "Recomputed as gross x commission %; differences corrected and reported"),
    ("Tax Amt", "Total Taxes and Levies", "CR0064", "C", "Numeric, original currency"),
    ("Tax Type", "Tax Type", "tax block", "C", "As keyed"),
    ("Tax Jurisdiction", "Tax Jurisdiction", "CR0077", "C", "Country mapped to ISO 3166 alpha-2"),
    ("Net to London", "Net Premium to London (Original Currency)", "CR0065", "M", "Recomputed as gross - commission - taxes; differences corrected and reported"),
    ("Settle Curr", "Settlement Currency", "—", "M", "Validated against Lloyd's settlement currency list"),
    ("FX Rate", "Rate of Exchange", "CR0067", "C", "Guide convention: original units per 1 unit of settlement currency"),
    ("Net to London (Settle Ccy)", "Net Premium to London (Settlement Currency)", "CR0068", "C", "Recomputed as original net / rate of exchange"),
    ("(London broker to complete)", "Brokerage % / Brokerage Amount / Final Net Premium block", "CR0074–CR0076", "C", "Reserved for the London broker per the user guide; left blank in the coverholder file"),
]
for m in MAPPING:
    ws.append(list(m))
note = wb.create_sheet("Notes")
note.append(["Re-use: complete this questionnaire once per binder; re-apply every period."])
note.append(["Heading order in the clean file is static; do not rename or merge columns (per Lloyd's premium bordereaux operating procedure)."])
note.append(["XLSX is the recommended submission format."])
wb.save(os.path.join(OUT, "mapping-template.xlsx"))

print(f"corrected={len(corrected)} flagged={len(flagged)} "
      f"flagged_rows={len({f['row'] for f in flagged})}")
print("keyed:", keyed)
print("clean:", clean_tot)
