from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/mnt/user-data/outputs/Capstone1_Payments_Failure_Reduction/04_Data_Model/Data_Model_Specification.xlsx"
F = "Arial"
hf = Font(name=F, bold=True, color="FFFFFF", size=10); hfill = PatternFill("solid", start_color="1F3864")
body = Font(name=F, size=10); bold = Font(name=F, size=10, bold=True)
wrap = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="BFBFBF"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
alt = PatternFill("solid", start_color="F2F5FA")
fact_fill = PatternFill("solid", start_color="D9E2F3")
wb = Workbook()

def sheet(name, headers, rows, widths, first=False, boldcols=1, shade_fn=None):
    ws = wb.active if first else wb.create_sheet()
    ws.title = name
    ws.append(headers)
    for r in rows: ws.append(list(r))
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]: c.font = hf; c.fill = hfill; c.alignment = Alignment(wrap_text=True, vertical="center")
    for ri, row in enumerate(ws.iter_rows(min_row=2), 0):
        for ci, c in enumerate(row, 1):
            c.font = bold if ci <= boldcols else body; c.alignment = wrap; c.border = box
            if ri % 2: c.fill = alt
    ws.freeze_panes = "B2"
    return ws

# ---------------- README ----------------
ws = wb.active; ws.title = "README"
L = [
    ("Data Model Specification - Capstone 1 (Workstream 4)", True),
    ("Draft v1, 11 Nov 2026. Technology-neutral logical model. Names, keys and thresholds are proposals to be adapted to the bank's real source systems.", False),
    ("", False),
    ("Sheets", True),
    ("Tables: every fact, dimension and bridge with its grain and SCD type. Columns: column-level specification. Bus_Matrix: which dimensions each fact uses. KPI_Mapping: how K1 to K10 are computed in the model. Dataset_Mapping: how each field in the synthetic dataset would load. DQ_Gates: load-time quality checks with the score the current data would get.", False),
    ("", False),
    ("Status labels: values marked PROPOSED are analyst proposals. Data-quality scores come from the Workstream 3 analysis workbook (synthetic data).", False),
]
for i, (t, b) in enumerate(L, 1):
    c = ws.cell(row=i, column=1, value=t); c.font = Font(name=F, size=12 if i == 1 else 10, bold=b); c.alignment = wrap
ws.column_dimensions["A"].width = 140

# ---------------- Tables ----------------
T = [
    ("fact_payment_event", "Fact (transaction)", "One row per state change of a payment attempt (initiated, routed, authorised, pending, settled, failed, reversed, retried).", "n/a", "Payments Operations", "Source of truth for timelines, pending duration and late-arriving events."),
    ("fact_payment", "Fact (accumulating snapshot)", "One row per payment instruction, identified by payment_id (the idempotency key). Updated as events arrive.", "n/a", "Payments", "Home of K1, K2, K3, K9. Retries are attempts under the same payment_id."),
    ("fact_fraud_event", "Fact (transaction)", "One row per fraud event on a payment: alert, decision, review outcome, confirmed fraud, loss booked, recovery.", "n/a", "Fraud / Risk", "Home of K4 and K5. Loss is signed so recoveries net off."),
    ("fact_investigation_case", "Fact (accumulating snapshot)", "One row per investigation or reconciliation case.", "n/a", "Operations", "Home of K6, K7, K8 cost. Tracks reopen and SLA applied."),
    ("fact_complaint", "Fact (transaction)", "One row per customer complaint about a payment.", "n/a", "Customer Service", "Home of K10."),
    ("bridge_case_payment", "Bridge", "One row per case-to-payment link, with an allocation weight that sums to 1 per case.", "n/a", "Operations", "Lets one case cover many payments without double counting cost or hours."),
    ("dim_date", "Dimension", "One row per calendar date.", "SCD0", "Data platform", "date_key = yyyymmdd."),
    ("dim_time_of_day", "Dimension", "One row per minute of day.", "SCD0", "Data platform", "Holds time_band (Night, Morning, Afternoon, Evening)."),
    ("dim_customer", "Dimension", "One row per customer version.", "SCD2", "Customer data owner", "Segment and risk band change over time; history needed (for example New_Customer)."),
    ("dim_merchant", "Dimension", "One row per merchant version.", "SCD2 on category, SCD1 on name", "Payments", "Merchant category re-mapping must not rewrite history."),
    ("dim_channel", "Dimension", "One row per channel.", "SCD1", "Payments", "Mobile, Web, Branch_Assisted."),
    ("dim_device", "Dimension", "One row per device environment.", "SCD1", "Digital channels", "Android, iOS, Desktop_Browser, Branch_Terminal."),
    ("dim_payment_type", "Dimension", "One row per payment method.", "SCD1", "Payments", "Instant_Transfer, Card, Bank_Transfer, Direct_Debit."),
    ("dim_geography", "Dimension", "One row per country and currency pair.", "SCD1", "Finance", "Currency conversion uses a daily rate table, not a fixed rate."),
    ("dim_processor", "Dimension", "One row per processor version.", "SCD2", "Payments", "Contract tier and routing group change over time. Unknown member = -1."),
    ("dim_status", "Dimension", "One row per final payment status.", "SCD0", "Payments", "outcome_category enforces the K1/K2/K3 identity (success + technical + business + risk + pending = 100%)."),
    ("dim_event_type", "Dimension", "One row per lifecycle event type.", "SCD0", "Payments Operations", "Defines valid event order."),
    ("dim_failure_reason", "Dimension", "One row per standard failure reason.", "SCD1 (mapping table is versioned)", "Payments Operations", "Standard taxonomy; raw processor codes map to it. Includes 'Unclassified' (-1) and 'Not applicable' (0)."),
    ("map_failure_code", "Mapping table", "One row per processor and raw failure code.", "Versioned", "Payments Operations", "Controls the mapping to dim_failure_reason; unmapped codes are reported daily."),
    ("dim_policy_version", "Dimension", "One row per version of routing, retry and fraud-rule configuration in force.", "SCD2", "Change management", "Lets KPI changes be attributed to interventions and supports benefit tracking."),
    ("dim_rule_family", "Dimension", "One row per fraud rule family.", "SCD1", "Fraud / Risk", "Cut for K5."),
    ("dim_case_trigger", "Dimension", "One row per case trigger (Reconciliation, FraudAlert, Complaint, Duplicate).", "SCD1", "Operations", "Holds the SLA hours applied per trigger and its version."),
]
sheet("Tables", ["Table", "Type", "Grain / purpose", "SCD", "Owner (steward)", "Notes"], T, [26, 26, 70, 22, 22, 70])

# ---------------- Columns ----------------
C = []
def col(t, c, ty, key, nullable, desc, src="", note=""): C.append((t, c, ty, key, nullable, desc, src, note))
# fact_payment_event
col("fact_payment_event", "event_key", "bigint", "PK", "No", "Surrogate key of the event.")
col("fact_payment_event", "payment_id", "text", "Degenerate dimension", "No", "Idempotency key of the payment instruction.", "Transaction_ID (needs a true idempotency key)", "Current Transaction_ID repeats for 36 rows.")
col("fact_payment_event", "attempt_no", "int", "", "No", "Attempt sequence within the payment (1 = first).", "retry_count + 1 (summary only)")
col("fact_payment_event", "event_type_key", "int", "FK dim_event_type", "No", "Lifecycle event.", "not available")
col("fact_payment_event", "event_ts", "timestamp", "", "No", "When the event happened at source.", "timestamp (initiation only)")
col("fact_payment_event", "ingest_ts", "timestamp", "", "No", "When the event reached the platform.", "not available", "Needed to detect late-arriving events.")
col("fact_payment_event", "processor_key", "int", "FK dim_processor", "No", "Processor that handled this event (-1 if unknown).", "processor")
col("fact_payment_event", "failure_reason_key", "int", "FK dim_failure_reason", "No", "Standard reason (0 = not applicable, -1 = unclassified).", "failure_code")
col("fact_payment_event", "stage_latency_ms", "int", "", "Yes", "Latency of this stage.", "latency_ms (end-to-end only)")
col("fact_payment_event", "source_batch_id", "text", "", "No", "Load batch for lineage.")
col("fact_payment_event", "record_hash", "text", "", "No", "Hash of source record for change detection and dedupe.")
col("fact_payment_event", "is_late_arriving", "boolean", "", "No", "True if ingest_ts - event_ts exceeds the late threshold (PROPOSED 1 hour).")
col("fact_payment_event", "restatement_version", "int", "", "No", "Increments when the event is corrected.")
# fact_payment
col("fact_payment", "payment_key", "bigint", "PK", "No", "Surrogate key.")
col("fact_payment", "payment_id", "text", "Unique business key", "No", "Idempotency key; unique per customer instruction.", "Transaction_ID")
col("fact_payment", "initiated_date_key", "int", "FK dim_date", "No", "Initiation date (attribution date for payment KPIs).", "timestamp")
col("fact_payment", "initiated_time_key", "int", "FK dim_time_of_day", "No", "Initiation minute.", "timestamp / time_band")
col("fact_payment", "customer_key", "bigint", "FK dim_customer (version in force)", "No", "Customer version at initiation.", "customer_segment")
col("fact_payment", "merchant_key", "bigint", "FK dim_merchant", "No", "Merchant version at initiation.", "merchant_category")
col("fact_payment", "channel_key", "int", "FK dim_channel", "No", "Initiation channel.", "channel")
col("fact_payment", "device_key", "int", "FK dim_device", "No", "Device environment.", "device_type")
col("fact_payment", "payment_type_key", "int", "FK dim_payment_type", "No", "Payment method.", "payment_type")
col("fact_payment", "geography_key", "int", "FK dim_geography", "No", "Country and currency.", "country, currency")
col("fact_payment", "final_processor_key", "int", "FK dim_processor", "No", "Processor of the final event (-1 if unknown).", "processor", "96 blank in current data (0.8%).")
col("fact_payment", "final_status_key", "int", "FK dim_status", "No", "Final status and outcome category.", "payment_status")
col("fact_payment", "final_failure_reason_key", "int", "FK dim_failure_reason", "No", "Standard reason for the final non-success.", "failure_code", "About 47% of technical and business failures blank in current data.")
col("fact_payment", "policy_version_key", "int", "FK dim_policy_version", "No", "Routing, retry and rule versions in force at initiation.", "not available")
col("fact_payment", "amount_local", "decimal", "", "No", "Amount in local currency.", "amount")
col("fact_payment", "amount_usd", "decimal", "", "No", "Amount converted at the daily rate.", "amount_usd_equiv (fixed rate)")
col("fact_payment", "fx_rate_used", "decimal", "", "No", "Rate applied.", "not available", "Current data uses fixed illustrative rates.")
col("fact_payment", "attempt_count", "int", "", "No", "Number of attempts (retries + 1).", "retry_count + 1")
col("fact_payment", "retry_count", "int", "", "No", "Retries after the first attempt.", "retry_count")
col("fact_payment", "pending_minutes", "decimal", "", "Yes", "Total time spent pending.", "pending_minutes")
col("fact_payment", "end_to_end_latency_ms", "int", "", "Yes", "Latency from initiation to final event.", "latency_ms", "Missing for 25% of branch-assisted rows.")
col("fact_payment", "duplicate_group_id", "text", "", "Yes", "Groups payments suspected to be duplicates of each other.", "not available")
col("fact_payment", "duplicate_suspected_flag", "boolean", "", "No", "Suspected duplicate from matching logic.", "duplicate_suspected_flag")
col("fact_payment", "has_fraud_alert", "boolean", "", "No", "Derived from fact_fraud_event; convenience only, never an independent source.", "fraud_alert_flag")
col("fact_payment", "first_event_ts", "timestamp", "", "No", "First event time.", "timestamp")
col("fact_payment", "final_event_ts", "timestamp", "", "Yes", "Final event time.", "not available")
col("fact_payment", "load_batch_id", "text", "", "No", "Lineage: batch that last updated the row.")
# fact_fraud_event
col("fact_fraud_event", "fraud_event_key", "bigint", "PK", "No", "Surrogate key.")
col("fact_fraud_event", "payment_key", "bigint", "FK fact_payment", "No", "Payment the event relates to.")
col("fact_fraud_event", "fraud_event_type", "text", "", "No", "ALERT_RAISED, DECISION_APPROVE, DECISION_DECLINE, REVIEW_OUTCOME, FRAUD_CONFIRMED, LOSS_BOOKED, RECOVERY.", "fraud_alert_flag, confirmed_fraud_flag")
col("fact_fraud_event", "event_ts", "timestamp", "", "No", "When the event happened.")
col("fact_fraud_event", "rule_family_key", "int", "FK dim_rule_family", "No", "Rule family that fired.", "not available")
col("fact_fraud_event", "reviewed_outcome_label", "text", "", "No", "GENUINE, FRAUD or UNKNOWN after analyst review.", "not available", "Needed for a true false-positive rate on declines (open item O-5).")
col("fact_fraud_event", "loss_amount_usd", "decimal", "", "Yes", "Signed amount: loss positive, recovery negative. Only on LOSS_BOOKED and RECOVERY rows.", "fraud_loss_usd")
col("fact_fraud_event", "cohort_date_key", "int", "FK dim_date", "No", "Initiation date of the payment (attribution cohort for K4).")
col("fact_fraud_event", "days_since_initiation", "int", "", "No", "Maturity of the event relative to initiation.", "not available", "Cohorts are 'matured' after the confirmation window (PROPOSED 90 days).")
# fact_investigation_case
col("fact_investigation_case", "case_key", "bigint", "PK", "No", "Surrogate key.")
col("fact_investigation_case", "case_id", "text", "Unique business key", "No", "Case identifier in the case system.")
col("fact_investigation_case", "case_trigger_key", "int", "FK dim_case_trigger", "No", "Single primary trigger of the case.", "reconciliation_flag, fraud_alert_flag, customer_complaint_flag")
col("fact_investigation_case", "opened_ts", "timestamp", "", "No", "Case opened.", "investigation_start")
col("fact_investigation_case", "closed_ts", "timestamp", "", "Yes", "Case closed (null while open); must be on or after opened_ts.", "investigation_end", "25 rows end before start in current data.")
col("fact_investigation_case", "sla_hours_applied", "decimal", "", "No", "SLA threshold in force for the trigger.", "Inputs: 8 h fraud, 24 h other (PROPOSED)")
col("fact_investigation_case", "within_sla_flag", "boolean", "", "Yes", "Closed within the SLA threshold.")
col("fact_investigation_case", "reopen_count", "int", "", "No", "Number of reopenings.", "not available")
col("fact_investigation_case", "staff_hours_logged", "decimal", "", "Yes", "Staff hours from time tracking.", "not available", "Needed to test the 1,200 hours per month fact.")
col("fact_investigation_case", "handling_cost_usd", "decimal", "", "No", "Case handling cost.", "handling_cost_usd (payment level)")
# bridge, complaint
col("bridge_case_payment", "case_key", "bigint", "PK part, FK fact_investigation_case", "No", "Case.")
col("bridge_case_payment", "payment_key", "bigint", "PK part, FK fact_payment", "No", "Payment.")
col("bridge_case_payment", "allocation_weight", "decimal", "", "No", "Share of the case cost and hours assigned to this payment; sums to 1 per case.")
col("fact_complaint", "complaint_key", "bigint", "PK", "No", "Surrogate key.")
col("fact_complaint", "payment_key", "bigint", "FK fact_payment", "No", "Payment complained about.", "customer_complaint_flag")
col("fact_complaint", "received_ts", "timestamp", "", "No", "Complaint received.")
col("fact_complaint", "category", "text", "", "Yes", "Complaint category.")
# dims
col("dim_customer", "customer_key", "bigint", "PK", "No", "Surrogate key of a customer version.")
col("dim_customer", "customer_id_token", "text", "Business key (tokenised)", "No", "Tokenised customer identifier; no direct personal data.")
col("dim_customer", "segment", "text", "", "No", "Retail_Mass, Retail_Affluent, SME, New_Customer.", "customer_segment")
col("dim_customer", "tenure_band", "text", "", "No", "Months since onboarding, banded.")
col("dim_customer", "risk_band", "text", "", "Yes", "Customer risk band.")
col("dim_customer", "effective_from / effective_to / is_current", "date / date / boolean", "", "No", "SCD2 validity range.")
col("dim_merchant", "merchant_key", "bigint", "PK", "No", "Surrogate key of a merchant version.")
col("dim_merchant", "merchant_category", "text", "", "No", "Category (from MCC).", "merchant_category")
col("dim_processor", "processor_key", "int", "PK", "No", "Surrogate key; -1 = Unknown.")
col("dim_processor", "processor_id / route_group / contract_tier", "text", "Business key", "No", "Processor identity and grouping.", "processor")
col("dim_failure_reason", "failure_reason_key", "int", "PK", "No", "Surrogate key; 0 = Not applicable, -1 = Unclassified.")
col("dim_failure_reason", "reason_category", "text", "", "No", "Technical, Business, Risk, Pending, Unclassified.", "derived from failure_code")
col("dim_failure_reason", "is_bank_controllable", "boolean", "", "No", "True for failures the bank can influence (technical).", "")
col("map_failure_code", "processor_id + raw_code", "text", "PK", "No", "Raw processor code.", "failure_code")
col("map_failure_code", "failure_reason_key / mapping_version", "int / int", "FK", "No", "Standard reason and mapping version.")
col("dim_policy_version", "policy_version_key", "int", "PK", "No", "Surrogate key.")
col("dim_policy_version", "routing_version / retry_rule_version / fraud_rule_set_version", "text", "", "No", "Versions in force.")
col("dim_policy_version", "effective_from / effective_to / change_ticket_id", "timestamp / timestamp / text", "", "No", "Validity and the change that created it.")
col("dim_status", "status_key", "int", "PK", "No", "Surrogate key.")
col("dim_status", "final_status / outcome_category", "text", "", "No", "SUCCESS, FAILED_TECHNICAL, FAILED_BUSINESS, DECLINED_RISK, PENDING and their category.", "payment_status")
sheet("Columns", ["Table", "Column", "Data type", "Key", "Nullable", "Description", "Source in current dataset", "Note"], C, [26, 40, 16, 30, 10, 70, 38, 56], boldcols=2)

# ---------------- Bus matrix ----------------
facts = ["fact_payment_event", "fact_payment", "fact_fraud_event", "fact_investigation_case", "fact_complaint"]
dims = ["dim_date", "dim_time_of_day", "dim_customer", "dim_merchant", "dim_channel", "dim_device", "dim_payment_type", "dim_geography",
        "dim_processor", "dim_status", "dim_event_type", "dim_failure_reason", "dim_policy_version", "dim_rule_family", "dim_case_trigger"]
use = {
    "fact_payment_event": {"dim_date", "dim_time_of_day", "dim_processor", "dim_event_type", "dim_failure_reason"},
    "fact_payment": {"dim_date", "dim_time_of_day", "dim_customer", "dim_merchant", "dim_channel", "dim_device", "dim_payment_type", "dim_geography", "dim_processor", "dim_status", "dim_failure_reason", "dim_policy_version"},
    "fact_fraud_event": {"dim_date", "dim_rule_family"},
    "fact_investigation_case": {"dim_date", "dim_case_trigger"},
    "fact_complaint": {"dim_date", "dim_channel"},
}
ws = wb.create_sheet("Bus_Matrix")
ws.append(["Fact \\ Dimension"] + dims)
for f_ in facts:
    ws.append([f_] + ["X" if d in use[f_] else "" for d in dims])
ws.append([])
ws.append(["Facts without a direct dimension reach it through payment_key (join to fact_payment first, then to the dimension). That keeps each fact at its own grain."])
for c in ws[1]:
    c.font = hf; c.fill = hfill; c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
for row in ws.iter_rows(min_row=2, max_row=1 + len(facts)):
    for ci, c in enumerate(row, 1):
        c.font = bold if ci == 1 else body; c.border = box; c.alignment = Alignment(horizontal="center" if ci > 1 else "left", vertical="center")
        if ci > 1 and c.value == "X": c.fill = fact_fill
ws.column_dimensions["A"].width = 30
for i in range(2, len(dims) + 2): ws.column_dimensions[get_column_letter(i)].width = 14
ws.row_dimensions[1].height = 42
ws.freeze_panes = "B2"

# ---------------- KPI mapping ----------------
K = [
    ("K1", "Payment success rate", "fact_payment", "count(payment_id) where dim_status.outcome_category = 'Success' / count(payment_id) in period (by initiated_date_key)", "dim_status, dim_channel, dim_processor, dim_geography, dim_payment_type, dim_time_of_day, dim_device", "One row per payment_id, so retries cannot inflate the denominator."),
    ("K2", "Technical decline rate", "fact_payment", "count(payment_id) where outcome_category = 'Technical' / same denominator as K1", "dim_failure_reason, dim_processor, dim_time_of_day, dim_device", "Same denominator as K1 keeps the identity."),
    ("K3", "Pending rate (over 15 minutes)", "fact_payment (pending_minutes from fact_payment_event)", "count(payment_id) where pending_minutes > 15 / same denominator", "dim_processor, dim_geography, dim_payment_type", "pending_minutes is rolled up from events, so late events restate it."),
    ("K4", "Fraud loss rate", "fact_fraud_event + fact_payment", "sum(loss_amount_usd) by cohort_date_key (matured) / sum(amount_usd) of payments with outcome 'Success' in the same cohort", "dim_customer, dim_merchant, dim_geography, dim_channel (via payment_key)", "Each fact aggregated to cohort date first, then divided. Loss attributed to the payment's initiation date."),
    ("K5", "False-positive rate (alerts)", "fact_fraud_event", "count(payments with ALERT_RAISED and reviewed_outcome_label = 'GENUINE') / count(payments with ALERT_RAISED and a reviewed label)", "dim_rule_family, dim_customer, dim_merchant", "Uses reviewed labels, which fixes the label bias seen in the current data."),
    ("K6", "Investigation SLA attainment", "fact_investigation_case", "count(case_key) where within_sla_flag = true / count(closed cases with closed_ts >= opened_ts)", "dim_case_trigger, dim_channel (via payments)", "One row per case, one primary trigger."),
    ("K7", "Reconciliation effort", "fact_investigation_case (+ bridge)", "sum(staff_hours_logged) for trigger 'Reconciliation' per month; cases per 1,000 payments from bridge-weighted counts", "dim_case_trigger", "Bridge weights sum to 1 per case, so hours are not double counted."),
    ("K8", "Cost per successful payment", "fact_investigation_case + fact_payment", "sum(handling_cost_usd) / count(payment_id) where outcome_category = 'Success'", "dim_channel, dim_processor, dim_payment_type", "Cost sits only in the case fact; aggregated to period before dividing."),
    ("K9", "Duplicate payment rate", "fact_payment", "count(payment_id) where duplicate_suspected_flag = true / count(payment_id) where retry_count >= 1", "dim_processor, dim_channel", "duplicate_group_id lets duplicates be counted once per group."),
    ("K10", "Complaint rate", "fact_complaint + fact_payment", "count(complaint_key) per 1,000 payments in the period", "dim_channel, dim_status (via payment_key)", "Complaints counted in their own fact; aggregated before dividing."),
]
sheet("KPI_Mapping", ["KPI", "Name", "Computed in", "Calculation in the model", "Main dimensions", "Double-counting control"], K, [7, 30, 36, 90, 56, 70], boldcols=2)

# ---------------- Dataset mapping ----------------
Dm = [
    ("Transaction_ID", "fact_payment.payment_id; fact_payment_event.payment_id", "Replace with a true idempotency key; de-duplicate at load; quarantine repeats", "36 repeated IDs (0.3%)"),
    ("timestamp", "fact_payment.initiated_date_key, initiated_time_key, first_event_ts", "Split into date and minute keys; keep full timestamp", "Initiation only, no event times"),
    ("channel", "dim_channel via fact_payment.channel_key", "Look up; unknown = new member", ""),
    ("country", "dim_geography via fact_payment.geography_key", "Look up country and currency pair", ""),
    ("payment_type", "dim_payment_type via fact_payment.payment_type_key", "Look up", ""),
    ("amount", "fact_payment.amount_local", "Load as is", ""),
    ("currency", "dim_geography.currency_code", "Look up", ""),
    ("amount_usd_equiv", "fact_payment.amount_usd, fx_rate_used", "Recompute with the daily rate for the initiation date; store the rate", "Fixed illustrative rates"),
    ("customer_segment", "dim_customer.segment (SCD2 version at initiation)", "Look up the customer version in force at initiation", "Static label would be overwritten on change"),
    ("merchant_category", "dim_merchant.merchant_category", "Look up merchant version", ""),
    ("payment_status", "dim_status via fact_payment.final_status_key", "Look up status and outcome category", "Final status only, no history"),
    ("failure_code", "map_failure_code -> dim_failure_reason", "Map raw code to the standard reason; blank on a failed payment = -1 Unclassified and counted in DQ gate G3", "About 47% of technical and business failures blank"),
    ("processor", "dim_processor via final_processor_key", "Blank = -1 Unknown; count in DQ gate G2", "96 blank (0.8%)"),
    ("latency_ms", "fact_payment.end_to_end_latency_ms; fact_payment_event.stage_latency_ms", "Null if not captured; flag values above the plausible bound", "14 outliers; branch rows 25% missing"),
    ("retry_count", "fact_payment.retry_count, attempt_count", "attempt_count = retry_count + 1; real attempts need event rows", "Summary only"),
    ("device_type", "dim_device via fact_payment.device_key", "Look up", ""),
    ("time_band", "dim_time_of_day.time_band", "Derived from time; not stored on the fact", ""),
    ("pending_minutes", "fact_payment.pending_minutes", "Recompute from events when available", "Cannot tell resolved from unresolved"),
    ("fraud_alert_flag", "fact_fraud_event (ALERT_RAISED); fact_payment.has_fraud_alert (derived)", "Create an ALERT_RAISED event per flagged payment", ""),
    ("confirmed_fraud_flag", "fact_fraud_event (FRAUD_CONFIRMED)", "Create an event per confirmed payment", "No confirmation date or reviewed label"),
    ("fraud_loss_usd", "fact_fraud_event.loss_amount_usd (LOSS_BOOKED)", "Create a LOSS_BOOKED event; recoveries as negative rows", "No recoveries in data"),
    ("duplicate_suspected_flag", "fact_payment.duplicate_suspected_flag, duplicate_group_id", "Load the flag; derive groups from matching logic", "Flag is a model output"),
    ("investigation_start", "fact_investigation_case.opened_ts", "One case per flagged payment (assumption: 1:1 until real case data)", "Cases may cover many payments"),
    ("investigation_end", "fact_investigation_case.closed_ts", "Reject or flag if before opened_ts (DQ gate G4)", "25 rows end before start"),
    ("reconciliation_flag", "fact_investigation_case.case_trigger_key = Reconciliation", "Primary trigger rule: pending, duplicate, technical failure, other", ""),
    ("handling_cost_usd", "fact_investigation_case.handling_cost_usd (via bridge weights)", "Allocate to the case; sum of weights = 1", "Cost drivers are assumptions"),
    ("customer_complaint_flag", "fact_complaint (one row per complaint)", "Create a complaint row per flagged payment", "Complaint time not in data"),
]
sheet("Dataset_Mapping", ["Dataset field", "Target in the model", "Load rule / transformation", "Issue in current data"], Dm, [28, 66, 80, 50])

# ---------------- DQ gates ----------------
G = [
    ("G1", "payment_id unique", "0 repeated keys", "Fail: 36 repeated IDs (0.30% of rows)", "Quarantine repeats; alert the data owner; load only the latest valid record"),
    ("G2", "Processor present", "Unknown processor at most 0.1% of rows (PROPOSED)", "Fail: 0.80% blank", "Load as Unknown (-1) and alert; fail the batch above the limit"),
    ("G3", "Failure reason mapped for every non-success", "At least 99% classified (PROPOSED)", "Fail: about 53% classified for technical and business failures", "Load as Unclassified (-1); publish the daily unmapped-code list; block KPI release for K2 if below the limit"),
    ("G4", "Case closed on or after opened", "0 cases with closed_ts before opened_ts", "Fail: 25 cases (3.5% of investigations)", "Quarantine and send to Operations to correct; exclude from K6 until fixed"),
    ("G5", "Latency within plausible bound", "At most 0.1% above 20,000 ms (PROPOSED)", "Fail marginally: 0.12% above the cap", "Flag as outlier; use median for latency KPIs"),
    ("G6", "Latency captured for branch-assisted payments", "At least 95% (PROPOSED)", "Fail: 75% captured", "Fix capture at the terminal; do not compare latency across channels until fixed"),
    ("G7", "Event order valid (terminal event after initiated)", "0 invalid sequences", "Not testable: no event history", "Quarantine the payment's events; flag for review"),
    ("G8", "Late-arriving share monitored", "Report share with ingest_ts - event_ts above 1 hour (threshold PROPOSED)", "Not testable: no ingest time", "Restate affected KPI periods; show data freshness on dashboards"),
    ("G9", "FX rate present for every payment date", "100% (PROPOSED)", "Not testable: fixed rate used", "Block conversion; use the previous business day rate and flag"),
    ("G10", "Bridge allocation weights sum to 1 per case", "Exactly 1 (tolerance 0.0001)", "Not testable: no bridge in current data", "Reject the case load"),
]
sheet("DQ_Gates", ["Gate", "Check", "Rule / threshold", "Score on the current (synthetic) data", "Action on failure"], G, [7, 46, 52, 52, 80])

wb.save(OUT)
print("saved")
