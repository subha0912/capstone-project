import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

df = pd.read_pickle("/home/claude/payments.pkl")
out = "/mnt/user-data/outputs/Capstone1_Payments_Failure_Reduction/03_Data_and_Analysis/Payments_Synthetic_Dataset.xlsx"

F = "Arial"
hdr_fill = PatternFill("solid", start_color="1F3864")
hdr_font = Font(name=F, bold=True, color="FFFFFF", size=10)
body = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
wrap = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()

# ---------------- README ----------------
ws = wb.active; ws.title = "README"
rows = [
    ("SYNTHETIC DATASET - NOT REAL CLIENT DATA", True),
    ("Project: Strive MBA Capstone 1 - Digital Payments Failure Reduction & Fraud-Aware Customer Experience", False),
    ("Prepared for: Shubhad Shraba Sabut (case pack working file)", False),
    ("Generated: 11 Nov 2026 with a seeded random generator (numpy default_rng, seed = 42). Script is reproducible.", False),
    ("", False),
    ("Why synthetic: the brief says to use supplied data if available, otherwise create a synthetic/anonymised dataset, label it, and document the generation rules. No real dataset was supplied.", False),
    ("Size: 12,000 records over 92 days (1 Jul - 30 Sep 2026), above the 500-record minimum.", False),
    ("Sheets: README | Generation_Rules | Data_Dictionary | Transactions", False),
    ("", False),
    ("IMPORTANT - how to use it honestly in the case pack:", True),
    ("1. Label every chart/table sourced from this file as 'synthetic data'.", False),
    ("2. The illustrative case facts (94.2% success, 1.8% pending >15 min, 9h median investigation, 0.07% fraud loss) were NOT hard-wired. Measure them from this data and report the gap. Do not force the data to match.", False),
    ("3. Conclusions about which processor / time band / device drives failures are properties of this synthetic data. Present them as 'what the analysis method would show on real data', not as facts about a real bank.", False),
    ("4. 1,200 staff-hours/month of reconciliation cannot be measured from transaction rows alone: it needs a scaling assumption (volume per month x minutes per case). Treat it as an assumption and show the working.", False),
]
for i, (t, b) in enumerate(rows, 1):
    c = ws.cell(row=i, column=1, value=t); c.font = Font(name=F, size=11 if b else 10, bold=b, color="C00000" if i == 1 else "000000"); c.alignment = wrap
ws.column_dimensions["A"].width = 130

# ---------------- Generation rules ----------------
ws = wb.create_sheet("Generation_Rules")
ws.append(["Area", "Rule used to generate the data", "Why it is there (analysis purpose)"])
rules = [
    ("Volume & calendar", "12,000 transactions; random day in 92-day window; hour-of-day weighted with an evening peak (18-22h).", "Time-band analysis needs realistic load shape."),
    ("Channel mix", "Mobile 50%, Web 34%, Branch_Assisted 16%.", "Channel segmentation."),
    ("Country / currency", "IN 70% (INR), AE 18% (AED), SG 12% (SGD). amount_usd_equiv uses fixed illustrative FX: INR 0.0119, AED 0.272, SGD 0.74.", "Cross-border behaviour; common currency for value-based KPIs."),
    ("Payment type", "Instant_Transfer 38%, Card 34%, Bank_Transfer 18%, Direct_Debit 10%.", "Payment-type segmentation."),
    ("Customer segment", "Retail_Mass 52%, Retail_Affluent 20%, SME 18%, New_Customer 10%. SME amounts x3.2.", "Segment impact and fraud exposure."),
    ("Processor", "PROC_A 52%, PROC_B 30%, PROC_C 18% (anonymised names).", "Processor comparison and routing."),
    ("Amounts", "Lognormal (mean 4.3, sigma 1.15) in USD-equivalent, capped 1 to 60,000.", "Skewed value distribution; value-weighted KPIs differ from count-based KPIs."),
    ("Technical failure (seeded)", "Base 2.1% x multipliers: PROC_C 1.35x; PROC_C in Evening_18-23 a further 3.2x; Android users in IN 1.3x; Card on PROC_B 1.25x; Branch_Assisted 0.7x; latency >1,500 ms 2.0x.", "Gives the analysis something to find: processor/time-band effect vs device/network effect (panel question: separating processor failure from customer/device issues)."),
    ("Pending (seeded)", "Base 1.2% x multipliers: PROC_B on non-IN country 3.0x; Bank_Transfer 1.6x; Night_00-05 1.5x. pending_minutes lognormal; >15 min is the 'long pending' definition.", "Pending-rate KPI and denominator design."),
    ("Business decline", "Base 1.8% (+1.0 pt Retail_Mass, +0.6 pt Direct_Debit): insufficient funds, limit exceeded, invalid beneficiary, customer auth failed.", "Customer-side failures that the bank should not 'fix' by routing; shows why technical decline rate is a separate KPI."),
    ("Fraud (seeded)", "confirmed fraud probability rises for New_Customer, Gaming, Travel, amount >3,000 USD, non-IN. 55% of confirmed fraud is approved (loss = amount). Alert flag rises with the same drivers; alerts have a 28% risk-decline rate vs 0.3% otherwise.", "Fraud vs customer-experience trade-off: risk declines hurt success rate; relaxing rules raises loss."),
    ("Retries & duplicates (seeded)", "Failed/pending items get 0-3 retries; 45% of retried technical failures recover to SUCCESS. 14% of pending-then-retried items are flagged duplicate_suspected.", "Panel question: how do you know a retry policy will not amplify duplicate payments?"),
    ("Complaints", "Probability by status (e.g. pending 30%, technical failure 22%, success 0.4%) +10 pts if duplicate suspected.", "Customer-impact analysis."),
    ("Reconciliation & investigation", "Flag if pending, duplicate, 55% of technical failures, or 0.4% random. Investigation exists if reconciliation OR fraud alert OR complaint. Duration lognormal, median target 9h; x0.8 for fraud alerts; x1.35 for Branch_Assisted. Start = 5-240 min after transaction.", "Investigation SLA and reconciliation-effort KPIs."),
    ("Handling cost (USD)", "0.04 base + ~5.2 if reconciled + ~3.1 if fraud alert + ~2.6 if complaint (normal draws, truncated).", "Cost per successful payment; automation business case. These unit costs are assumptions."),
    ("SEEDED DATA-QUALITY ISSUE 1", "240 failed rows have a blank failure_code (2.0% of rows).", "Test: do you notice, quantify, and treat it explicitly?"),
    ("SEEDED DATA-QUALITY ISSUE 2", "96 rows have a blank processor.", "Missing dimension key; cannot be assigned to a processor."),
    ("SEEDED DATA-QUALITY ISSUE 3", "36 rows reuse another row's Transaction_ID.", "Grain/double-counting issue: Transaction_ID is not a safe unique key."),
    ("SEEDED DATA-QUALITY ISSUE 4", "14 latency values of 25-90 seconds.", "Outliers; decide treat vs cap vs exclude."),
    ("SEEDED DATA-QUALITY ISSUE 5", "25 investigations where end time is before start time.", "Event-sequencing / late-arriving data issue; affects SLA calculation."),
    ("SEEDED DATA-QUALITY ISSUE 6", "469 Branch_Assisted rows have no latency captured (25% of branch rows).", "Systematic missingness: latency comparison across channels is biased."),
    ("Limitations", "Single quarter, no seasonality; no true device telemetry; no multi-leg event table (one row per transaction, not per event); fraud labels are instantaneous, with no chargeback delay.", "Be explicit in the case pack about what a real event-level feed would add (links to Workstream 4)."),
]
for r in rules: ws.append(list(r))

# ---------------- Data dictionary ----------------
ws2 = wb.create_sheet("Data_Dictionary")
ws2.append(["Field", "Type", "In brief's suggested list?", "Definition", "Allowed values / notes"])
dd = [
    ("Transaction_ID", "text", "Yes", "Identifier of the payment attempt.", "Expected unique; 36 duplicates seeded (see DQ issue 3)."),
    ("timestamp", "datetime", "Yes", "Time the payment was initiated.", "1 Jul - 30 Sep 2026, assumed same timezone."),
    ("channel", "category", "Yes", "Initiation channel.", "Mobile, Web, Branch_Assisted"),
    ("country", "category", "Yes", "Customer/payment country.", "IN, AE, SG"),
    ("payment_type", "category", "Yes", "Payment method.", "Instant_Transfer, Card, Bank_Transfer, Direct_Debit"),
    ("amount", "number", "Yes", "Amount in local currency.", ""),
    ("currency", "category", "Yes", "Currency of amount.", "INR, AED, SGD"),
    ("amount_usd_equiv", "number", "No (added)", "Amount converted at fixed illustrative FX.", "Use for value-weighted KPIs."),
    ("customer_segment", "category", "Yes", "Customer group.", "Retail_Mass, Retail_Affluent, SME, New_Customer"),
    ("merchant_category", "category", "Yes", "Beneficiary/merchant type.", "Retail, Utilities, Travel, Gaming, Financial_Services, Education, P2P"),
    ("payment_status", "category", "Yes", "Final observed status of the attempt.", "SUCCESS, FAILED_TECHNICAL, FAILED_BUSINESS, DECLINED_RISK, PENDING"),
    ("failure_code", "category", "Yes", "Reason code for non-success.", "Blank for success; 240 failed rows blank in error."),
    ("processor", "category", "Yes", "Payment processor (anonymised).", "PROC_A, PROC_B, PROC_C; 96 blank."),
    ("latency_ms", "number", "Yes", "End-to-end processing latency.", "Not captured for 25% of Branch_Assisted; 14 outliers."),
    ("retry_count", "integer", "Yes", "Number of retries of this payment.", "0-3"),
    ("device_type", "category", "No (added)", "Customer device/environment.", "Android, iOS, Desktop_Browser, Branch_Terminal"),
    ("time_band", "category", "No (added)", "Derived from timestamp hour.", "Night_00-05, Morning_06-11, Afternoon_12-17, Evening_18-23"),
    ("pending_minutes", "number", "No (added)", "Minutes in pending state (0 if never pending).", ">15 = long pending"),
    ("fraud_alert_flag", "0/1", "Yes", "Risk engine raised an alert.", ""),
    ("confirmed_fraud_flag", "0/1", "Yes", "Fraud later confirmed.", ""),
    ("fraud_loss_usd", "number", "No (added)", "Loss booked when confirmed fraud was approved.", "Needed to compute fraud loss rate by value."),
    ("duplicate_suspected_flag", "0/1", "No (added)", "Suspected duplicate payment caused by retry.", ""),
    ("investigation_start", "datetime", "Yes", "Start of investigation case.", "Blank if no case."),
    ("investigation_end", "datetime", "Yes", "End of investigation case.", "25 rows end before start."),
    ("reconciliation_flag", "0/1", "Yes", "Item required manual reconciliation.", ""),
    ("handling_cost_usd", "number", "Yes (handling_cost)", "Operational cost to handle the item (USD).", "Cost drivers are assumptions."),
    ("customer_complaint_flag", "0/1", "Yes", "Customer complained about this payment.", ""),
]
for r in dd: ws2.append(list(r))

# ---------------- Transactions ----------------
ws3 = wb.create_sheet("Transactions")
cols = list(df.columns)
ws3.append(cols)
for rec in df.itertuples(index=False):
    row = []
    for v in rec:
        if pd.isna(v):
            row.append(None)
        elif isinstance(v, pd.Timestamp):
            row.append(v.to_pydatetime())
        elif hasattr(v, "item"):
            row.append(v.item())
        else:
            row.append(v)
    ws3.append(row)
for r in range(2, len(df) + 2):
    for col in (2, 23, 24):
        ws3.cell(row=r, column=col).number_format = "yyyy-mm-dd hh:mm"

# ---------------- formatting ----------------
for sheet, widths in ((ws, None), (wb["Generation_Rules"], [30, 90, 60]), (ws2, [24, 10, 22, 48, 60])):
    if widths:
        for i, w in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(i)].width = w
for sheet in (wb["Generation_Rules"], ws2, ws3):
    for c in sheet[1]:
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.freeze_panes = "A2"
for sheet in (wb["Generation_Rules"], ws2):
    for row in sheet.iter_rows(min_row=2):
        for c in row:
            c.font = body; c.alignment = wrap; c.border = box
for i, name in enumerate(cols, 1):
    ws3.column_dimensions[get_column_letter(i)].width = max(14, min(26, len(name) + 4))
ws3.column_dimensions["B"].width = 18
ws3.column_dimensions["W"].width = 18
ws3.column_dimensions["X"].width = 18
ws3.auto_filter.ref = ws3.dimensions
for r in ws3.iter_rows(min_row=2):
    for c in r:
        c.font = body

wb.save(out)
print("saved", out)
