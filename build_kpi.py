from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

out = "/mnt/user-data/outputs/Capstone1_Payments_Failure_Reduction/02_KPIs_and_Requirements/KPI_Dictionary.xlsx"
F = "Arial"
hf = Font(name=F, bold=True, color="FFFFFF", size=10)
hfill = PatternFill("solid", start_color="1F3864")
body = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
wrap = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
alt = PatternFill("solid", start_color="F2F5FA")

wb = Workbook()

def sheet(name, headers, rows, widths, first=False, boldcol=1):
    ws = wb.active if first else wb.create_sheet()
    ws.title = name
    ws.append(headers)
    for r in rows: ws.append(r)
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.font = hf; c.fill = hfill; c.alignment = Alignment(wrap_text=True, vertical="center")
    for ri, row in enumerate(ws.iter_rows(min_row=2), 0):
        for ci, c in enumerate(row, 1):
            c.font = bold if ci <= boldcol else body
            c.alignment = wrap; c.border = box
            if ri % 2: c.fill = alt
    ws.freeze_panes = "C2" if len(headers) > 6 else "A2"
    return ws

# ---------------- README ----------------
ws = wb.active; ws.title = "README"
lines = [
    ("KPI Dictionary - Capstone 1 (Workstream 2)", True),
    ("Draft v1, 11 Nov 2026. Definitions are written so that every KPI can be computed from the synthetic dataset (Payments_Synthetic_Dataset.xlsx) and later from a real event-level feed.", False),
    ("", False),
    ("Sheets", True),
    ("KPI_Dictionary: 10 KPIs (the 8 required by the brief plus 2 guardrail KPIs: duplicate payment rate and complaint rate). Formula, numerator, denominator, grain, refresh, owner, source fields.", False),
    ("Denominator_Design: the counting rules and the identity that ties success, technical, business, risk and pending outcomes together.", False),
    ("Guardrails_and_Limits: which KPIs must be read together, the KPI not to optimise alone, and the method for measurable control limits.", False),
    ("Open_Items: values that are proposals or need agreement with stakeholders before they are used.", False),
    ("", False),
    ("Status labels: values marked 'PROPOSED' are analyst proposals, not agreed targets. Nothing here is a real client result.", False),
]
for i, (t, b) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t); c.font = Font(name=F, size=12 if i == 1 else 10, bold=b); c.alignment = wrap
ws.column_dimensions["A"].width = 140

# ---------------- KPI dictionary ----------------
H = ["ID", "KPI", "Decision it informs", "Formula (plain words)", "Numerator", "Denominator (and exclusions)", "Grain / counting unit",
     "Refresh", "Owner", "Source fields (dataset names)", "Standard cuts", "Read together with", "Known pitfalls / bias"]
K = [
    ("K1", "Payment success rate", "Overall health of end-to-end completion; headline for Payments and COO.",
     "Successful payments divided by all payment attempts in the period.",
     "Attempts with final payment_status = SUCCESS.",
     "All attempts with a final status (SUCCESS, FAILED_TECHNICAL, FAILED_BUSINESS, DECLINED_RISK, PENDING). Each Transaction_ID counted once. Test/internal transfers excluded where identifiable.",
     "One payment instruction (customer intent); retries are rolled up into the same instruction, not counted as new attempts.",
     "Hourly for operations monitoring; daily and monthly for reporting.", "Head of Payments",
     "payment_status, Transaction_ID, timestamp", "channel, processor, country, payment_type, time_band, device_type, customer_segment",
     "K4 fraud loss rate, K5 false-positive rate, K9 duplicate rate",
     "PENDING counts as not yet successful, so a delayed settlement lowers the rate until resolved. A falling share of business declines (customer-side) can move the rate with no change in the bank's performance, so also report the controllable success rate = SUCCESS / (attempts - FAILED_BUSINESS)."),
    ("K2", "Technical decline rate", "Size of the fixable failure pool; targets routing, retry and processor fixes.",
     "Technically failed payments divided by all payment attempts.",
     "Attempts with payment_status = FAILED_TECHNICAL.",
     "Same denominator as K1 (all attempts) so that K1 and K2 decompose cleanly. Business declines and risk declines stay in the denominator but are not in the numerator.",
     "One payment instruction.", "Hourly (alerting) and daily.", "Head of Payments / Payment Operations",
     "payment_status, failure_code, processor, latency_ms", "processor, time_band, device_type, failure_code, payment_type, country",
     "K1 success rate, K9 duplicate rate",
     "Failed rows with a blank failure_code must be reported as 'Unclassified', never dropped (about 2% of rows in the synthetic data). Technical vs business classification depends on the mapping of each processor's codes."),
    ("K3", "Pending rate (over 15 minutes)", "Customer uncertainty and reconciliation load driver.",
     "Payments that stayed pending for more than 15 minutes divided by all attempts.",
     "Attempts with pending_minutes > 15 (resolved later or still pending).",
     "All attempts (same as K1). Report separately the count still unresolved at close of day.",
     "One payment instruction.", "Every 5 to 15 minutes for alerting (PROPOSED); daily for reporting.", "Payment Operations lead",
     "pending_minutes, payment_status, processor, country, payment_type", "processor x country x payment_type, time_band",
     "K7 reconciliation effort, K10 complaint rate",
     "Pending time is only reliable with event history (timestamps of each state change). A single status field hides how long an item was pending (links to the WS4 event-level model)."),
    ("K4", "Fraud loss rate", "Whether approval gains are being paid for in fraud losses; core Risk control.",
     "Confirmed fraud losses divided by the value of approved payments, for a given initiation cohort.",
     "Sum of fraud_loss_usd (losses booked where confirmed fraud was approved), in USD-equivalent.",
     "Sum of amount_usd_equiv for payments with payment_status = SUCCESS. Cohort is the payment initiation date, matured for the confirmation window (PROPOSED 90 days); younger cohorts shown as provisional. Like-for-like variant for testing the brief's 0.07% fact: same numerator over the value of ALL attempts (see assumption A-005).",
     "One approved payment; loss counted once per Transaction_ID. Fraud events and recoveries sit in a separate fact table to avoid double counting.",
     "Weekly, on matured cohorts; daily provisional view.", "Chief Risk Officer / Head of Fraud",
     "fraud_loss_usd, confirmed_fraud_flag, amount_usd_equiv, payment_status", "customer_segment, merchant_category, country, channel, amount band",
     "K1 success rate, K5 false-positive rate",
     "Fraud is confirmed late (chargebacks, customer reports), so recent cohorts understate loss. Loss is by value, so a few large cases dominate; show counts and values. Recoveries are not in the dataset."),
    ("K5", "False-positive rate (fraud alerts)", "Customer friction caused by fraud rules; guides rule tuning.",
     "Fraud alerts that were not confirmed as fraud divided by all fraud alerts.",
     "Items with fraud_alert_flag = 1 and confirmed_fraud_flag = 0 (after the confirmation window).",
     "Items with fraud_alert_flag = 1. Alerts still under review excluded and reported as a count. Decline variant: DECLINED_RISK items not confirmed fraud / all DECLINED_RISK items.",
     "One alerted payment.", "Weekly on matured data; daily provisional.", "Head of Fraud / CRO",
     "fraud_alert_flag, confirmed_fraud_flag, payment_status", "customer_segment, merchant_category, rule family (not in dataset), country",
     "K4 fraud loss rate, K1 success rate",
     "Label bias: a declined payment rarely gets a confirmed-fraud label, because the customer never lost money. Without a reviewed-outcome label the decline false-positive rate is overstated and unknowable from this data alone. Needs analyst review outcomes."),
    ("K6", "Investigation SLA attainment", "Whether cases are resolved fast enough; targets Operations capacity and automation.",
     "Share of closed investigation cases resolved within the SLA threshold; report median and 90th percentile duration alongside.",
     "Closed cases with (investigation_end - investigation_start) <= threshold T.",
     "Closed cases with valid timestamps. Rows with end before start are excluded and counted (25 in the synthetic data). Threshold T is PROPOSED at 24 hours for reconciliation and complaint cases and 8 hours for fraud alerts, to be agreed.",
     "One investigation case.", "Daily; monthly trend.", "Head of Operations",
     "investigation_start, investigation_end, reconciliation_flag, fraud_alert_flag, customer_complaint_flag", "channel, case trigger, payment_type",
     "K8 cost per successful payment, K10 complaint rate",
     "Hitting the SLA by closing cases carelessly. Add a reopen rate once data exists. A case can have several triggers; assign a single primary trigger so it is counted once."),
    ("K7", "Reconciliation effort", "Size of the manual workload; business case for automated reconciliation.",
     "Staff-hours per month spent on manual reconciliation = reconciled cases x average handling minutes per case / 60. Also reconciliation cases per 1,000 attempts.",
     "Count of items with reconciliation_flag = 1 (cases); staff-hours need time-tracking data.",
     "For the intensity measure: all attempts. For hours: not a ratio.",
     "One reconciliation case.", "Monthly; weekly trend.", "Head of Operations",
     "reconciliation_flag, handling_cost_usd (proxy), timesheet data (not in dataset)", "cause (pending, duplicate, technical failure), channel, processor",
     "K3 pending rate, K8 cost per successful payment",
     "The brief's 1,200 staff-hours per month cannot be measured from transaction rows. It needs monthly volume x minutes per case from time tracking. Treat as an assumption (A-004) and show the working."),
    ("K8", "Cost per successful payment", "Efficiency; links operational failure cost to the business case.",
     "Total handling cost of all attempts (including failed and pending) divided by the number of successful payments.",
     "Sum of handling_cost_usd across all attempts.",
     "Count of attempts with payment_status = SUCCESS.",
     "Period total (month) with segment cuts.", "Monthly.", "COO office with Finance",
     "handling_cost_usd, payment_status", "channel, processor, payment_type, country",
     "K6 SLA attainment, K10 complaint rate",
     "Handling cost only. Platform, support and change costs are added in the TCO (WS7) for a fully loaded view. Unit cost drivers in the dataset are assumptions (A-007)."),
    ("K9", "Duplicate payment rate (retry-induced)", "Guardrail for retry rules and customer harm from double charging.",
     "Suspected duplicate payments divided by payments that were retried at least once.",
     "Items with duplicate_suspected_flag = 1.",
     "Items with retry_count >= 1. Also report per 1,000 successful payments.",
     "One payment instruction.", "Daily.", "Head of Payments",
     "duplicate_suspected_flag, retry_count, payment_status", "prior status before retry (pending vs failed), processor, channel",
     "K1 success rate, K2 technical decline rate",
     "The flag is itself the output of matching logic (same beneficiary, amount, time window), so false negatives are possible. Idempotency keys give a stronger control (see requirement FR4)."),
    ("K10", "Complaint rate", "Customer impact of failures and pending items; checks notification and fraud rules.",
     "Complaints per 1,000 payment attempts; also complaints per failed or pending item.",
     "Items with customer_complaint_flag = 1.",
     "All attempts (per 1,000); failed and pending attempts (per item).",
     "One payment instruction.", "Daily; monthly trend.", "Head of Customer Service",
     "customer_complaint_flag, payment_status", "payment_status, channel, customer_segment, processor",
     "K3 pending rate, K1 success rate",
     "Under-reporting: only some customers complain, so the rate understates impact. Complaints can lag the payment by days."),
]
sheet("KPI_Dictionary", H, K, [6, 26, 34, 40, 34, 52, 30, 26, 24, 34, 34, 30, 60], boldcol=2)

# ---------------- Denominator design ----------------
D = [
    ("Counting unit", "One row per payment instruction (customer intent). Retries are summarised in retry_count on the same instruction. A real event-level feed would hold one row per state change and roll up to this grain.", "Prevents a retried payment from being counted as several attempts and distorting success rate."),
    ("De-duplication", "Count each Transaction_ID once, keeping the latest valid record. Duplicate IDs are reported as a data-quality count (36 in the synthetic data).", "Transaction_ID is not guaranteed unique; double counting would inflate every rate."),
    ("Common denominator", "K1, K2 and K3 share the same denominator (all attempts). Outcomes decompose as: success + technical failure + business decline + risk decline + pending = 100% of attempts.", "Makes the KPIs reconcile to each other, so a gain in K1 must show up as a fall in another category."),
    ("Controllable view", "Controllable success rate = SUCCESS / (all attempts - FAILED_BUSINESS). Shown beside K1, never instead of it.", "Separates what the bank can influence from customer-side declines."),
    ("Pending", "PENDING stays in the denominator and out of the numerator until a final status is received. Pending over 15 minutes is its own KPI (K3).", "Avoids hiding unresolved payments."),
    ("Fraud vs payment KPIs", "K4 uses approved value as denominator and the initiation cohort for attribution. Fraud loss is booked once per Transaction_ID; recoveries and chargebacks sit in a separate fact. K5 uses alerts as denominator.", "Prevents double counting between payment outcomes and fraud outcomes."),
    ("Value vs count", "Rates are by count unless stated (K4 is by value). Show value-weighted success rate as a secondary view.", "A few large payments can behave differently from the many small ones."),
    ("Time attribution", "Payment KPIs are attributed to initiation time. Investigation KPIs to case start. Fraud KPIs to initiation cohort after maturity.", "Avoids mixing periods when events arrive late."),
    ("Missing values", "Missing processor or failure_code is reported as 'Unknown' / 'Unclassified' in every cut, with the count displayed.", "Silent dropping of rows biases segment comparisons."),
]
sheet("Denominator_Design", ["Topic", "Rule", "Why it matters"], D, [24, 100, 70])

# ---------------- Guardrails and limits ----------------
G = [
    ("K1 Payment success rate", "K4 fraud loss rate, K5 false-positive rate, K9 duplicate rate", "A success-rate gain counts only if fraud loss stays under its ceiling and duplicate payments do not rise.", "REFUSE TO OPTIMISE IN ISOLATION: success rate can be raised by loosening fraud rules or retrying aggressively, both of which push harm elsewhere."),
    ("K2 Technical decline rate", "K9 duplicate rate", "More retries lower technical declines but can create duplicates.", "Retry rules are accepted only if K9 stays within its control limit."),
    ("K5 False-positive rate", "K4 fraud loss rate", "Fewer false positives is good only if confirmed fraud does not rise.", "Do not tune rules on K5 alone."),
    ("K6 Investigation SLA", "K8 cost, K10 complaints, reopen rate", "SLA can be hit by closing cases carelessly.", "Track reopen rate once data is available."),
    ("K8 Cost per successful payment", "K6 SLA, K10 complaints", "Cost can be cut by degrading service.", "Cost savings count only if SLA and complaints hold."),
]
ws = sheet("Guardrails_and_Limits", ["Primary KPI", "Must be read with", "Why", "Rule / note"], G, [30, 44, 70, 80])

start = len(G) + 4
ws.cell(row=start - 1, column=1, value="Method for measurable control limits (thresholds to be set from the Workstream 3 baseline; not set here)").font = bold
hdr = ["KPI", "Chart type", "Limit calculation", "Signal rule", "Threshold status"]
for i, h in enumerate(hdr, 1):
    c = ws.cell(row=start, column=i, value=h); c.font = hf; c.fill = hfill; c.alignment = Alignment(wrap_text=True, vertical="center")
L = [
    ("K1, K2, K3, K5, K9 (rates)", "p-chart (proportion) by day or week", "Centre = baseline proportion p-bar from stable history; limits = p-bar +/- 3 x sqrt(p-bar x (1 - p-bar) / n) for each period's n.", "One point outside limits, or 8 consecutive points on one side of the centre line.", "Not set: needs baseline (WS3)."),
    ("K4 Fraud loss rate", "Cohort chart on matured cohorts, with a hard ceiling", "Early-warning line = baseline + 2 standard deviations from bootstrap of weekly cohorts; hard ceiling = risk appetite agreed with the CRO.", "Early warning triggers review; ceiling breach triggers rollback of the latest fraud-rule or routing change.", "Not set: ceiling needs CRO agreement."),
    ("K6 SLA attainment", "p-chart plus 90th percentile duration", "As for rates, using closed cases per period as n.", "Outside limits, or 90th percentile above the SLA threshold for two periods.", "Not set: SLA threshold to be agreed."),
    ("K7, K8 (monthly values)", "Individuals and moving-range (XmR) chart", "Centre = mean; limits = mean +/- 2.66 x average moving range.", "Point outside limits, or 8 consecutive points on one side.", "Not set: needs at least 12 monthly points; synthetic data has about 3."),
    ("K10 Complaint rate", "u-chart (rate per 1,000)", "Centre = baseline complaints per 1,000; limits = u-bar +/- 3 x sqrt(u-bar / (n/1000)).", "Outside limits or 8 in a row.", "Not set: needs baseline (WS3)."),
]
for ri, r in enumerate(L, start + 1):
    for ci, v in enumerate(r, 1):
        c = ws.cell(row=ri, column=ci, value=v); c.font = bold if ci == 1 else body; c.alignment = wrap; c.border = box
for i, w in enumerate([30, 44, 70, 80, 36], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---------------- Open items ----------------
O = [
    ("O-1", "SLA thresholds (24 h reconciliation/complaint, 8 h fraud alert) are PROPOSED.", "Head of Operations, CRO", "Before WS3 SLA analysis is finalised"),
    ("O-2", "Fraud confirmation window of 90 days for cohort maturity is PROPOSED. Synthetic data has no chargeback lag.", "CRO / Head of Fraud", "Before K4 baseline is finalised"),
    ("O-3", "Fraud loss ceiling (risk appetite) is not set.", "CRO", "Before WS5 and WS8 option scoring"),
    ("O-4", "Reconciliation staff-hours need time-tracking data; until then use the cost proxy and a stated assumption.", "Head of Operations", "Before WS7 business case"),
    ("O-5", "Reviewed-outcome labels for risk declines are needed to measure the decline false-positive rate properly.", "Head of Fraud", "Flag as a data gap in WS3 and WS4"),
    ("O-6", "Alert latency target (5 to 15 minutes for pending and technical-decline alerts) is PROPOSED.", "Head of Payments, CIO", "Before WS6 architecture"),
]
sheet("Open_Items", ["ID", "Item", "Who must agree", "Needed by"], O, [8, 100, 30, 40])

wb.save(out)
print("saved", out)
