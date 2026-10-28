#!/usr/bin/env python3
"""
verify.py — INDEPENDENT GATE for the bordereaux sample.
Re-reads the raw file, the error ledger and the delivered clean file, and
proves, without trusting the cleaning pipeline's own claims:
  1. Every planted CORRECTABLE error is fixed in the clean file to exactly
     the ledger's expected value, and is recorded in the corrections list.
  2. Every planted FLAGGED error appears in the Client Action list and its
     row carries an annotation in the clean file; the value is NOT silently
     altered.
  3. Every pipeline correction that is not itself a planted error is a
     consequential recalculation fully explained by a planted commission
     correction on the same row (delta-for-delta).
  4. No other value changed: numeric fields differ between raw and clean
     ONLY where the ledger or rule 3 explains it; structural sweeps (YoA,
     UMR, duplicates, date logic, sector, currency) match the ledger exactly.
  5. Totals reconcile: gross identical; net/settlement deltas fully explained.
Writes VERIFICATION.md with exact counts and a PASS/FAIL verdict.
Exit code 0 = PASS, 1 = FAIL.
"""
import csv, os, sys
from datetime import datetime

OUT = os.path.dirname(os.path.abspath(__file__))
TOL = 0.011
fails, notes = [], []

def chk(cond, msg):
    if cond:
        notes.append(f"PASS — {msg}")
    else:
        fails.append(msg)

def pf(v):
    try:
        return float(str(v).replace(",", "").replace("£", "").replace("€", ""))
    except (TypeError, ValueError):
        return None

def parse_date(s):
    s = (s or "").strip()
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None

# --------------------------------------------------------------------- load --
with open(os.path.join(OUT, "raw-bordereau.csv"), newline="", encoding="utf-8") as f:
    rows = list(csv.reader(f))
hdr = next(i for i, r in enumerate(rows) if r and r[0] == "Coverholder")
headers = rows[hdr]
raw = []
for r in rows[hdr + 1:]:
    if not r or r[0] == "TOTALS":
        continue
    raw.append(dict(zip(headers, r + [""] * (len(headers) - len(r)))))
clean = list(csv.DictReader(open(os.path.join(OUT, "clean-bordereau-v52.csv"), encoding="utf-8")))
ledger = list(csv.DictReader(open(os.path.join(OUT, "error-ledger.csv"), encoding="utf-8")))
corr = list(csv.DictReader(open(os.path.join(OUT, "exceptions-corrected.csv"), encoding="utf-8")))
flag = list(csv.DictReader(open(os.path.join(OUT, "exceptions-client-action.csv"), encoding="utf-8")))

chk(len(raw) == 200 and len(clean) == 200, f"row counts: raw={len(raw)}, clean={len(clean)} (expected 200/200)")
chk([r["Cert Ref"] for r in raw] == [r["Certificate Ref"] for r in clean],
    "row order and certificate references preserved from raw to clean")

corr_keys = {(c["row"], c["field"]) for c in corr}
flag_keys = {(c["row"], c["field"]) for c in flag}
led_corr = [e for e in ledger if e["disposition"] == "CORRECTED"]
led_flag = [e for e in ledger if e["disposition"] == "FLAGGED_FOR_CLIENT"]

# ------------------------------------------------- 1. corrected errors fixed -
fixed = 0
for e in led_corr:
    row = clean[int(e["raw_row"]) - 1]
    val = row[e["field"]]
    exp = e["expected_correct_value"]
    ok = (pf(val) is not None and pf(exp) is not None and abs(pf(val) - pf(exp)) <= TOL) \
        or str(val).strip() == exp.strip()
    recorded = (e["raw_row"], e["field"]) in corr_keys
    if ok and recorded:
        fixed += 1
    else:
        fails.append(f"ledger {e['error_id']} row {e['raw_row']} field {e['field']}: "
                     f"clean value {val!r} vs expected {exp!r}; recorded={recorded}")
chk(fixed == len(led_corr),
    f"all {len(led_corr)} planted correctable errors fixed to the expected value and recorded (verified {fixed})")

# --------------------------------------------------- 2. flagged errors handled -
flagged_ok = 0
for e in led_flag:
    row = clean[int(e["raw_row"]) - 1]
    in_list = (e["raw_row"], e["field"]) in flag_keys
    annotated = "ACTION REQUIRED" in row["Validation Note"]
    # value must NOT be silently altered:
    if e["error_class"] == "C6":
        unchanged = row["Certificate Ref"] == e["raw_value"]
    elif e["error_class"] == "C8":
        unchanged = row["Insured Trade or Industry Sector"] == ""
    elif e["error_class"] == "C10":
        unchanged = row["Risk Expiry Date"] == e["raw_value"]
    elif e["error_class"] == "C9":
        unchanged = row["Original Currency"] == ""
    elif e["error_class"] == "C3":
        unchanged = row["Risk Code"] == e["raw_value"]
    else:
        unchanged = False
    if in_list and annotated and unchanged:
        flagged_ok += 1
    else:
        fails.append(f"ledger {e['error_id']} row {e['raw_row']} field {e['field']}: "
                     f"in_client_action={in_list} annotated={annotated} value_unchanged={unchanged}")
chk(flagged_ok == len(led_flag),
    f"all {len(led_flag)} planted flagged errors are on the Client Action list, annotated in the clean file, and not silently altered (verified {flagged_ok})")

# ------------------------------------------- 3. consequential corrections only -
led_keys = {(e["raw_row"], e["field"]) for e in ledger}
extra = [c for c in corr if (c["row"], c["field"]) not in led_keys]
conseq_ok = 0
for c in extra:
    i = int(c["row"]) - 1
    c7_on_row = [e for e in ledger if e["raw_row"] == c["row"] and e["error_class"] == "C7"]
    comm_led = [e for e in c7_on_row if e["field"] == "Commission Amount"]
    net_delta = pf(clean[i]["Net Premium to London (Original Currency)"]) - pf(raw[i]["Net to London"])
    delta = pf(clean[i][c["field"]]) - pf(raw[i]["Net to London"] if "Settlement" not in c["field"]
                                           else raw[i]["Net to London (Settle Ccy)"])
    ok = False
    if c["field"] == "Net Premium to London (Original Currency)" and comm_led:
        comm_delta = pf(comm_led[0]["expected_correct_value"]) - pf(comm_led[0]["raw_value"])
        ok = abs(delta - (-comm_delta)) <= TOL
    elif c["field"] == "Net Premium to London (Settlement Currency)" and c7_on_row:
        if clean[i]["Original Currency"] == "GBP":
            ok = abs(delta - net_delta) <= TOL          # settlement mirrors original (same currency)
        elif raw[i]["FX Rate"].strip():
            ok = abs(delta - net_delta / float(raw[i]["FX Rate"])) <= TOL
    if ok:
        conseq_ok += 1
    else:
        fails.append(f"unexplained or miscalculated consequential correction at row {c['row']} field {c['field']}")
chk(conseq_ok == len(extra),
    f"all {len(extra)} non-ledger corrections are consequential recalculations proven delta-for-delta from a ledgered C7 correction on the same row")

# --------------------------------------------- 4. nothing else changed (money) -
NUMERIC = [("Gross Prem (this time)", "Gross Premium Paid This Time"),
           ("Comm Amt", "Commission Amount"), ("Tax Amt", "Total Taxes and Levies"),
           ("Net to London", "Net Premium to London (Original Currency)"),
           ("Net to London (Settle Ccy)", "Net Premium to London (Settlement Currency)")]
allowed_delta = {(e["raw_row"], e["field"]) for e in ledger} | {(c["row"], c["field"]) for c in extra}
unexplained = []
for i, (r, c) in enumerate(zip(raw, clean), start=1):
    for rk, ck in NUMERIC:
        rv, cv = pf(r[rk]), pf(c[ck])
        if rv is None or cv is None:
            if (rv is None) != (cv is None) and ck == "Net Premium to London (Settlement Currency)":
                continue  # flagged-currency row: settlement legitimately blank
            if rv is None and cv is None:
                continue
            unexplained.append((i, ck, rv, cv))
        elif abs(rv - cv) > TOL and (str(i), ck) not in allowed_delta:
            unexplained.append((i, ck, rv, cv))
chk(not unexplained,
    f"no unexplained numeric changes between raw and clean (found {len(unexplained)}: {unexplained[:5]})")

# ------------------------------------------------------- structural sweeps ---
def rows_where(pred, rows_):
    return {str(i) for i, r in enumerate(rows_, start=1) if pred(r)}

chk(all(c["Year of Account"] == "2024" for c in clean), "clean file: Year of Account = 2024 on all 200 rows")
chk(rows_where(lambda r: r["YoA"].strip() != "2024", raw)
    == {e["raw_row"] for e in ledger if e["error_class"] == "C1"},
    "YoA sweep: raw rows with wrong YoA are exactly the C1 ledger rows")
chk(all(c["Unique Market Reference (UMR)"] == "B0876HLB250001X" for c in clean),
    "clean file: UMR = contract UMR on all 200 rows")
chk(rows_where(lambda r: r["UMR / Binder Ref"].strip() != "B0876HLB250001X", raw)
    == {e["raw_row"] for e in ledger if e["error_class"] == "C4"},
    "UMR sweep: raw rows with malformed UMR are exactly the C4 ledger rows")
from collections import defaultdict
groups = defaultdict(list)
for i, r in enumerate(raw, start=1):
    if r["Trans Type"].strip() == "Original Premium":
        groups[r["Cert Ref"].strip()].append(str(i))
dup_all = {j for js in groups.values() if len(js) > 1 for j in js}
flagged_cert_rows = {f["row"] for f in flag if f["field"] == "Certificate Ref"}
c6_ledger_rows = {e["raw_row"] for e in ledger if e["error_class"] == "C6"}
chk(flagged_cert_rows == dup_all and c6_ledger_rows <= dup_all,
    f"duplicate sweep: every row of every duplicated Original Premium certificate pair is flagged "
    f"({len(dup_all)} rows across {sum(1 for js in groups.values() if len(js)>1)} pairs; "
    f"{len(c6_ledger_rows)} planted + {len(dup_all - c6_ledger_rows)} counterpart rows)")
chk(rows_where(lambda c: c["Risk Expiry Date"] <= c["Risk Inception Date"], clean)
    == {e["raw_row"] for e in ledger if e["error_class"] == "C10"},
    "date sweep: rows with expiry <= inception in the clean file are exactly the C10 flagged rows (carried, not altered)")
chk(rows_where(lambda c: c["Insured Trade or Industry Sector"] == "" and c["Section No"] in ("1", "3"), clean)
    == {e["raw_row"] for e in ledger if e["error_class"] == "C8"},
    "sector sweep: commercial rows missing sector in the clean file are exactly the C8 flagged rows")
chk(rows_where(lambda c: c["Original Currency"] == "", clean)
    == {e["raw_row"] for e in ledger if e["error_class"] == "C9" and e["disposition"] == "FLAGGED_FOR_CLIENT"},
    "currency sweep: the only blank original currency in the clean file is the flagged indeterminate row")
AUTH = {"1": {"B5", "B3"}, "2": {"HP", "B4"}, "3": {"W3"}}
bad_codes = {str(i) for i, c in enumerate(clean, start=1)
             if c["Risk Code"] not in AUTH[c["Section No"]]}
chk(bad_codes == {e["raw_row"] for e in ledger if e["error_class"] == "C3"
                  and e["disposition"] == "FLAGGED_FOR_CLIENT"},
    "risk-code sweep: every clean risk code is authorised for its section, except the flagged 'ZZ' row")

# ------------------------------------------------------------------- totals --
def total(rows_, key, parser=pf):
    return round(sum(parser(r[key]) or 0.0 for r in rows_), 2)

g_raw, g_clean = total(raw, "Gross Prem (this time)"), total(clean, "Gross Premium Paid This Time")
chk(abs(g_raw - g_clean) <= TOL, f"gross totals reconcile exactly: raw {g_raw:,.2f} = clean {g_clean:,.2f}")
c7_rows = {e["raw_row"] for e in ledger if e["error_class"] == "C7"}
c7_net_rows = {e["raw_row"] for e in ledger if e["error_class"] == "C7"
               and e["field"] in ("Commission Amount", "Net Premium to London (Original Currency)")}
n_delta_rows = {str(i) for i, (r, c) in enumerate(zip(raw, clean), start=1)
                if abs((pf(r["Net to London"]) or 0) - (pf(c["Net Premium to London (Original Currency)"]) or 0)) > TOL}
chk(n_delta_rows == c7_net_rows,
    f"net premium changes occur on exactly the {len(c7_net_rows)} C7 commission/net rows and nowhere else")
s_delta_rows = {str(i) for i, (r, c) in enumerate(zip(raw, clean), start=1)
                if pf(r["Net to London (Settle Ccy)"]) is not None
                and pf(c["Net Premium to London (Settlement Currency)"]) is not None
                and abs(pf(r["Net to London (Settle Ccy)"]) - pf(c["Net Premium to London (Settlement Currency)"])) > TOL}
chk(s_delta_rows == c7_rows,
    f"settlement amounts change on exactly the {len(c7_rows)} C7 rows and nowhere else")
n_raw, n_clean = total(raw, "Net to London"), total(clean, "Net Premium to London (Original Currency)")
s_raw = total(raw, "Net to London (Settle Ccy)")
s_clean = round(sum(pf(c["Net Premium to London (Settlement Currency)"]) or 0.0 for c in clean), 2)
flag_ccy_row = [e for e in ledger if e["error_class"] == "C9" and e["disposition"] == "FLAGGED_FOR_CLIENT"][0]
flag_i = int(flag_ccy_row["raw_row"]) - 1
flag_keyed_settle = pf(raw[flag_i]["Net to London (Settle Ccy)"])
s_deltas = round(sum((pf(clean[i]["Net Premium to London (Settlement Currency)"]) or 0.0)
                     - (pf(raw[i]["Net to London (Settle Ccy)"]) or 0.0)
                     for i in range(200) if str(i + 1) in c7_rows), 2)
chk(abs(s_clean - (s_raw - flag_keyed_settle + s_deltas)) <= 0.05,
    f"settlement totals identity holds: clean {s_clean:,.2f} = keyed {s_raw:,.2f} "
    f"- flagged-currency row held out {flag_keyed_settle:,.2f} + C7 deltas {s_deltas:,.2f}")
notes.append(f"INFO — totals: gross raw {g_raw:,.2f} = clean {g_clean:,.2f}; "
             f"net raw {n_raw:,.2f} -> clean {n_clean:,.2f} (delta {round(n_clean-n_raw,2):,.2f}, on C7 rows only); "
             f"settlement raw {s_raw:,.2f} -> clean {s_clean:,.2f} (C7 deltas {s_deltas:,.2f}; the flagged-currency "
             f"row's keyed settlement {flag_keyed_settle:,.2f} is held out of the clean settlement total pending confirmation)")

# ------------------------------------------------------------------- verdict -
verdict = "PASS" if not fails else "FAIL"
n_corr = len(led_corr)
n_flag = len(led_flag)
lines = [
    "# VERIFICATION — Bordereaux Sample Gate\n",
    f"**Verdict: {verdict}**\n",
    "## Counts\n",
    f"- Rows: raw 200 → clean 200",
    f"- Planted errors (error-ledger.csv): {len(ledger)} total",
    f"- Planted correctable: {n_corr} — verified fixed to the expected value: {fixed}",
    f"- Planted flagged-for-client: {n_flag} — verified on the Client Action list, annotated, values unaltered: {flagged_ok}",
    f"- Corrections recorded by the pipeline: {len(corr)} = {n_corr} planted fixes + {len(extra)} consequential recalculations (each proven delta-for-delta)",
    f"- Client Action items recorded: {len(flag)} (the {n_flag} planted items plus the counterpart rows of each "
    f"duplicated certificate pair, which are flagged as well — a duplicated reference implicates both rows)",
    f"- Planted errors accounted for: {fixed + flagged_ok} of {len(ledger)} "
    f"({'100%' if fixed + flagged_ok == len(ledger) else 'INCOMPLETE'})\n",
    "## Independent checks\n",
]
lines += [f"- {n}" for n in notes]
if fails:
    lines.append("\n## Failures\n")
    lines += [f"- FAIL — {f}" for f in fails]
lines.append("\n## Method of verification\n"
             "verify.py re-reads raw-bordereau.csv, error-ledger.csv, clean-bordereau-v52.csv and the two "
             "machine-readable exception lists. It does not import or trust the cleaning pipeline's logic: "
             "expected values come from the ledger written at generation time, consequential corrections are "
             "proven arithmetically, and structural sweeps compare the raw and clean files field by field.")
with open(os.path.join(OUT, "VERIFICATION.md"), "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
print(f"VERDICT: {verdict} | planted={len(ledger)} fixed={fixed}/{n_corr} flagged={flagged_ok}/{n_flag} "
      f"consequential={conseq_ok}/{len(extra)} unexplained_numeric={len(unexplained)}")
for f_ in fails:
    print("FAIL:", f_)
sys.exit(0 if verdict == "PASS" else 1)
