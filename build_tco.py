"""Workstream 7: 3-year TCO, benefits, sensitivities and breakeven.
All results are live formulas over the Inputs and Cost_Lines sheets.
Measured rates are carried from Payments_Analysis_Workbook.xlsx (green font, source noted);
illustrative assumptions are blue. Nothing here is a real client figure."""
import sys
from openpyxl import Workbook
from openpyxl.comments import Comment
from xlstyle import *

OUT = sys.argv[1] if len(sys.argv) > 1 else "TCO_and_Business_Case.xlsx"
wb = Workbook()

# ------------------------------------------------------------------ README
ws = wb.active; ws.title = "README"
title(ws, "Workstream 7 - 3-year TCO, business case and sensitivities (synthetic case)",
      "Draft v1, 11 Nov 2026. Prepared by Shubhad Shraba Sabut. Illustrative costs and assumptions: replace with bank quotes and data before any decision.")
lines = [
    "How to read this workbook",
    "Inputs: every assumption in one place. Blue = analyst assumption. Green = measured value carried from Payments_Analysis_Workbook.xlsx (source cell named). Yellow fill = the few assumptions the answer is most sensitive to.",
    "Volume: monthly attempts are DERIVED from the brief's own facts (1,200 reconciliation hours / measured reconciliation share / assumed minutes per case). It is not a known bank figure.",
    "Cost_Lines: one-time and recurring cost lines for the two deployment approaches from Workstream 6 (A = cloud-native managed, B = hybrid incremental) plus the interventions from Workstream 5 (Common). Each recurring line = fixed per year + variable per 1M attempts.",
    "TCO: 3-year cost by line, year and growth scenario (base, +20%, +40% transaction volume).",
    "Benefits: benefit per attempt (low / high) built from measured rates, then scaled by volume and a realisation ramp.",
    "Summary: TCO, benefits, net, NPV, breakeven volume and cost per successful payment for each option and scenario.",
    "Stage_Gates: which spend is committed now and which is released only when a gate is passed.",
    "Interpretation of '+20% / +40% growth': volume in every year is 20% / 40% above the base plan (which already grows at the base growth rate). Logged as assumption A-016.",
    "Key reading: at the volume implied by the brief's facts (~110k attempts a month) the full portfolio does not pay back on hard savings within 3 years. The case turns positive above the breakeven volume shown in Summary. Real monthly volume is therefore the first number to confirm (leadership decision LD-1).",
    "Benefits are NOT additive with each other beyond what is modelled: duplicate handling inside reconciliation is excluded from BN3, and BN1/BN4 use the same realised success gain.",
]
for i, t in enumerate(lines, 4):
    ws.cell(row=i, column=1, value=t).font = BOLD if i == 4 else BODY
    ws.cell(row=i, column=1).alignment = WRAP
ws.column_dimensions["A"].width = 140

# ------------------------------------------------------------------ Inputs
wi = wb.create_sheet("Inputs")
title(wi, "Inputs and assumptions", "Change values in column C only. Everything else recalculates.")
header(wi, 3, ["ID", "Item", "Value", "Unit", "Source / status"])
R = {}  # name -> absolute ref
rows = [
    ("SEC", "Volume"),
    ("brief_hours", "Manual reconciliation hours per month (brief, illustrative)", 1200, "hours / month", "Brief case fact 5 (A-004). Not testable from rows.", "in"),
    ("recon_share", "Share of attempts needing manual reconciliation", 0.032849, "share", "Measured: Analysis workbook, CaseFacts (393 / 11,964 in-scope attempts)", "link"),
    ("min_case", "Handling minutes per reconciliation case", 20, "minutes", "ASSUMPTION (A-017). Consistent with measured handling cost of ~USD 5.9 per case at USD 18/hour.", "key"),
    ("derived_m", "Implied monthly attempts (derived)", "=C5*60/C7/C6", "attempts / month", "Formula: hours x 60 / minutes per case / reconciliation share", "f"),
    ("base_m", "Base monthly attempts used (rounded)", "=ROUND(C8,-4)", "attempts / month", "Rounded derived volume. Replace with the bank's real volume (LD-1).", "fkey"),
    ("alt_m", "Alternative scale anchor: monthly attempts", 1000000, "attempts / month", "ASSUMPTION: typical regional-bank scale, used only to show how the answer flips (A-018).", "in"),
    ("g", "Base annual volume growth", 0.10, "per year", "ASSUMPTION: brief says volumes are growing; rate not given (A-016).", "in"),
    ("u1", "Growth scenario 1: volume above base plan", 0.20, "share", "Brief: sensitivity to +20% transaction growth", "in"),
    ("u2", "Growth scenario 2: volume above base plan", 0.40, "share", "Brief: sensitivity to +40% transaction growth", "in"),
    ("SEC", "Measured rates (synthetic data, Workstream 3)"),
    ("succ", "Payment success rate (K1)", 0.929957, "share", "Analysis workbook, OutcomeMix (11,126 / 11,964)", "link"),
    ("tech", "Technical decline rate (K2)", 0.019726, "share", "Analysis workbook, OutcomeMix (236 / 11,964)", "link"),
    ("o1_lo", "O1 upper-bound success gain, low (PROC_C evening to PROC_A level)", 0.002233, "share of attempts", "Analysis workbook, Opportunity O1 (0.22 pp)", "link"),
    ("o1_hi", "O1 upper-bound success gain, high", 0.004566, "share of attempts", "Analysis workbook, Opportunity O1 (0.46 pp)", "link"),
    ("compl", "Complaints per attempt (K10)", 0.017469, "share", "Measured: 209 / 11,964", "link"),
    ("compl_pend", "Share of complaints on PENDING payments", 0.344498, "share", "Measured: 72 / 209", "link"),
    ("compl_tech", "Share of complaints on FAILED_TECHNICAL payments", 0.248804, "share", "Measured: 52 / 209", "link"),
    ("dup", "Suspected duplicate payments per attempt", 0.002758, "share", "Measured: 33 / 11,964 (all after a pending retry, hypothesis C1)", "link"),
    ("alerts", "Fraud alerts per attempt", 0.019141, "share", "Measured: 229 / 11,964", "link"),
    ("hcost", "Handling cost per attempt (all causes)", 0.312503, "USD / attempt", "Measured: USD 3,738.79 / 11,964 (unit costs are assumptions A-007)", "link"),
    ("SEC", "Unit values (assumptions)"),
    ("v_pay", "Value of one recovered payment (net revenue + avoided retry/contact)", 1.50, "USD", "ASSUMPTION (A-019). Excludes customer-retention value (shown separately in Summary).", "key"),
    ("rate_h", "Loaded cost per operations staff hour", 18, "USD / hour", "ASSUMPTION (A-017): blended IN/AE/SG operations rate", "key"),
    ("c_compl", "Handling cost per complaint", 2.6, "USD", "Dataset unit cost (A-007). Likely understated: fully loaded contact-centre cost is higher.", "in"),
    ("c_alert", "Handling cost per fraud alert", 3.1, "USD", "Dataset unit cost (A-007)", "in"),
    ("c_dup", "Customer remediation per duplicate (refund, fee reversal, goodwill), beyond reconciliation handling", 5, "USD", "ASSUMPTION (A-020). Reconciliation handling of duplicates is already inside BN2, so not counted again.", "in"),
    ("SEC", "Effect sizes (assumptions, low / high)"),
    ("real_lo", "Realisation of O1 upper bound, low", 0.5, "share", "ASSUMPTION: capacity limits on PROC_A, partial routing", "in"),
    ("real_hi", "Realisation of O1 upper bound, high", 0.8, "share", "ASSUMPTION", "in"),
    ("rec_lo", "Reconciliation hours avoided, low (O4 + O2 + O6)", 0.30, "share", "Analysis workbook Inputs (sensitivity range)", "link"),
    ("rec_hi", "Reconciliation hours avoided, high", 0.60, "share", "Analysis workbook Inputs (sensitivity range)", "link"),
    ("dup_lo", "Duplicates removed by idempotency + no-retry-while-pending, low", 0.80, "share", "ASSUMPTION (A-021)", "in"),
    ("dup_hi", "Duplicates removed, high", 0.95, "share", "ASSUMPTION (A-021)", "in"),
    ("pend_lo", "Pending-related complaints avoided by status notifications, low", 0.20, "share", "ASSUMPTION (A-022)", "in"),
    ("pend_hi", "Pending-related complaints avoided, high", 0.40, "share", "ASSUMPTION (A-022)", "in"),
    ("al_lo", "Fraud-alert volume avoided by rule tuning, low (test may fail)", 0.0, "share", "ASSUMPTION: conditional on the O5 controlled test passing", "in"),
    ("al_hi", "Fraud-alert volume avoided by rule tuning, high", 0.30, "share", "ASSUMPTION: conditional on the O5 controlled test passing", "in"),
    ("SEC", "Timing and finance"),
    ("ramp1", "Benefit realisation, year 1", 0.40, "share of run-rate", "ASSUMPTION: roadmap delivers most fixes in months 3-12", "in"),
    ("ramp2", "Benefit realisation, year 2", 0.90, "share of run-rate", "ASSUMPTION", "in"),
    ("ramp3", "Benefit realisation, year 3", 1.00, "share of run-rate", "ASSUMPTION", "in"),
    ("disc", "Discount rate for NPV", 0.10, "per year", "ASSUMPTION: replace with bank hurdle rate", "in"),
    ("cont", "Contingency on one-time costs", 0.15, "share", "ASSUMPTION: delivery-risk buffer", "in"),
]
r = 4
for row in rows:
    if row[0] == "SEC":
        wi.cell(row=r, column=1, value=row[1]).font = SEC
        r += 1; continue
    key, item, val, unit, src, kind = row
    R[key] = f"Inputs!$C${r}"
    wi.cell(row=r, column=1, value=key)
    wi.cell(row=r, column=2, value=item)
    c = wi.cell(row=r, column=3, value=val)
    wi.cell(row=r, column=4, value=unit)
    wi.cell(row=r, column=5, value=src)
    for cc in range(1, 6):
        wi.cell(row=r, column=cc).font = BODY; wi.cell(row=r, column=cc).alignment = WRAP; wi.cell(row=r, column=cc).border = BOX
    c.font = {"in": INPUT, "key": INPUT, "link": LINK, "f": BODY, "fkey": BOLD}[kind]
    if kind in ("key", "fkey"): c.fill = YELLOW
    if unit in ("share", "per year", "share of attempts", "share of run-rate"):
        c.number_format = '0.000%' if key in ("recon_share", "o1_lo", "o1_hi", "dup", "alerts", "compl", "tech", "succ") else '0.0%'
    elif "USD" in unit:
        c.number_format = '$#,##0.00'
    else:
        c.number_format = '#,##0'
    r += 1
# check the hard-wired positions used by derived_m formula
assert R["brief_hours"] == "Inputs!$C$5" and R["recon_share"] == "Inputs!$C$6" and R["min_case"] == "Inputs!$C$7" and R["derived_m"] == "Inputs!$C$8", R
widths(wi, [12, 62, 14, 18, 80])
wi.freeze_panes = "A4"

# ------------------------------------------------------------------ Volume
wv = wb.create_sheet("Volume")
title(wv, "Annual payment attempts by scenario", "Year t volume = base monthly x 12 x (1 + base growth)^(t-1) x (1 + scenario uplift).")
header(wv, 3, ["Scenario", "Uplift vs base plan", "Year 1", "Year 2", "Year 3", "3-year total"])
scen = [("Base", "0"), ("+20% volume", R["u1"]), ("+40% volume", R["u2"])]
VROW = {}
for i, (name, up) in enumerate(scen):
    rr = 4 + i; VROW[i] = rr
    wv.cell(row=rr, column=1, value=name)
    wv.cell(row=rr, column=2, value=f"={up}").number_format = '0%'
    for t in range(3):
        col = 3 + t
        wv.cell(row=rr, column=col, value=f"={R['base_m']}*12*(1+{R['g']})^{t}*(1+$B{rr})").number_format = NUM
    wv.cell(row=rr, column=6, value=f"=SUM(C{rr}:E{rr})").number_format = NUM
body(wv, 4, 6, 1, 6)
wv["A8"] = "Growth factor sums used for breakeven (base growth)"; wv["A8"].font = SEC
wv["A9"] = "S = sum over years of (1+g)^(t-1)"; wv["B9"] = f"=1+(1+{R['g']})+(1+{R['g']})^2"
wv["A10"] = "Rr = sum over years of ramp_t x (1+g)^(t-1)"; wv["B10"] = f"={R['ramp1']}+{R['ramp2']}*(1+{R['g']})+{R['ramp3']}*(1+{R['g']})^2"
wv["A11"] = "Discounted S (for NPV)"; wv["B11"] = f"=1/(1+{R['disc']})+(1+{R['g']})/(1+{R['disc']})^2+(1+{R['g']})^2/(1+{R['disc']})^3"
wv["A12"] = "Discounted Rr (for NPV)"; wv["B12"] = f"={R['ramp1']}/(1+{R['disc']})+{R['ramp2']}*(1+{R['g']})/(1+{R['disc']})^2+{R['ramp3']}*(1+{R['g']})^2/(1+{R['disc']})^3"
for rr in range(9, 13):
    wv.cell(row=rr, column=1).font = BODY; wv.cell(row=rr, column=2).font = BODY; wv.cell(row=rr, column=2).number_format = '0.000'
widths(wv, [48, 18, 16, 16, 16, 18])

# ------------------------------------------------------------------ Cost lines
wc = wb.create_sheet("Cost_Lines")
title(wc, "Cost lines (inputs) - USD, illustrative",
      "Option A = cloud-native managed lakehouse (in-region). Option B = lean hybrid on the existing warehouse. Common = intervention build and run, needed under either option. Blue = assumption.")
header(wc, 3, ["ID", "Option", "Category (brief)", "Type", "Line item", "Fixed USD (one-time, or per year)",
               "Variable USD per 1M attempts per year", "Stage", "Basis / note"])
lines = [
    # id, opt, cat, type, item, fixed, var, stage, note
    ("A-M1", "A", "Migration", "One-time", "Historic data migration, backfill and pipeline build to lakehouse", 160000, 0, "Gated", "~8 FTE-months data engineering + vendor PS"),
    ("A-I1", "A", "Integration", "One-time", "Event feeds: payment hub, 3 processors, fraud engine, case mgmt, CRM", 140000, 0, "Committed", "Event-level feed (FR1) is needed under either option"),
    ("A-S1", "A", "Security", "One-time", "Key management, tokenisation, network isolation, pen test, PCI scoping", 70000, 0, "Committed", ""),
    ("A-T1", "A", "Training", "One-time", "Platform, KPI and runbook training (ops, risk, analysts)", 25000, 0, "Committed", ""),
    ("A-C1", "A", "Change", "One-time", "Change management, process redesign, comms", 45000, 0, "Committed", ""),
    ("A-P1", "A", "Cloud / platform", "Recurring", "Managed streaming, storage, lakehouse compute", 60000, 12000, "Gated", "Variable part scales with event volume"),
    ("A-P2", "A", "Cloud / platform", "Recurring", "BI licences", 18000, 0, "Committed", ""),
    ("A-R1", "A", "Support", "Recurring", "Platform run team: 2.0 incremental FTE", 90000, 0, "Gated", "USD 45k per FTE (assumption)"),
    ("A-S2", "A", "Security", "Recurring", "Security monitoring, key rotation, audits", 25000, 0, "Committed", ""),
    ("B-M1", "B", "Migration", "One-time", "Curated layer and KPI marts on the existing warehouse; limited backfill", 40000, 0, "Gated", "Gate G1 (month 3): build once real volume is known. Open table formats keep Option A possible"),
    ("B-I1", "B", "Integration", "One-time", "Event feeds (payment hub, 3 processors, fraud engine, case mgmt, CRM) + alert rules (O3)", 90000, 0, "Committed", "Event-level feed (FR1) and alerting (FR3) are needed under either option"),
    ("B-S1", "B", "Security", "One-time", "Tokenisation, access model, pen test", 30000, 0, "Committed", "Reuses existing on-prem controls"),
    ("B-T1", "B", "Training", "One-time", "KPI, alerting and runbook training", 10000, 0, "Committed", ""),
    ("B-C1", "B", "Change", "One-time", "Change management, process redesign, comms", 25000, 0, "Committed", ""),
    ("B-P1", "B", "Cloud / platform", "Recurring", "On-prem event bus + warehouse capacity (amortised); existing BI licences reused", 15000, 18000, "Committed", "Less elastic: capacity bought in steps, so higher variable rate than A"),
    ("B-R1", "B", "Support", "Recurring", "Run effort: 0.5 incremental FTE (rest absorbed by the existing BI team)", 22500, 0, "Committed", "USD 45k per FTE (assumption)"),
    ("B-S2", "B", "Security", "Recurring", "Security monitoring and audits", 10000, 0, "Committed", ""),
    ("X-O1", "Common", "Interventions", "One-time", "O1 Processor routing rules + PROC_C peak failover", 45000, 0, "Committed", "Workstream 5"),
    ("X-O2", "Common", "Interventions", "One-time", "O2 Safe retry: idempotency key, no retry while pending, backoff", 70000, 0, "Committed", "Workstream 5; requirement FR4"),
    ("X-O6", "Common", "Interventions", "One-time", "O6 Customer status notifications for pending > 15 min", 40000, 0, "Committed", "Workstream 5; requirement FR7"),
    ("X-O4", "Common", "Interventions", "One-time", "O4 Automated reconciliation (matching engine, exception queue)", 110000, 0, "Gated", "Gate G2: time-tracking confirms hours"),
    ("X-O5", "Common", "Interventions", "One-time", "O5 Fraud champion/challenger harness + reviewed-label workflow", 40000, 0, "Gated", "Gate G3: CRO ceiling agreed, labels flowing"),
    ("X-R1", "Common", "Support", "Recurring", "Reconciliation engine maintenance", 20000, 0, "Gated", ""),
    ("X-P1", "Common", "Cloud / platform", "Recurring", "Notification messaging (SMS / push)", 0, 150, "Committed", "~1.35% of attempts x ~USD 0.01"),
]
CL0 = 4
for i, ln in enumerate(lines):
    rr = CL0 + i
    for j, v in enumerate(ln):
        c = wc.cell(row=rr, column=1 + j, value=v)
        c.font = INPUT if j in (5, 6) else BODY
        c.alignment = WRAP; c.border = BOX
    wc.cell(row=rr, column=6).number_format = USD
    wc.cell(row=rr, column=7).number_format = USD
CL1 = CL0 + len(lines) - 1
widths(wc, [8, 9, 16, 11, 58, 16, 16, 11, 52])
wc.freeze_panes = "A4"
CLR = f"$4:${CL1}"

# ------------------------------------------------------------------ TCO
wt = wb.create_sheet("TCO")
title(wt, "3-year TCO by line, year and scenario (USD)",
      "One-time cost (with contingency) falls in Year 1. Recurring = fixed + variable x attempts / 1M. Columns grouped by scenario.")
hdr = ["ID", "Option", "Category", "Type"]
for s in ["Base", "+20%", "+40%"]:
    hdr += [f"{s} Y1", f"{s} Y2", f"{s} Y3", f"{s} 3-yr"]
header(wt, 3, hdr)
T0 = 4
for i in range(len(lines)):
    rr = T0 + i; src = CL0 + i
    wt.cell(row=rr, column=1, value=f"=Cost_Lines!A{src}")
    wt.cell(row=rr, column=2, value=f"=Cost_Lines!B{src}")
    wt.cell(row=rr, column=3, value=f"=Cost_Lines!C{src}")
    wt.cell(row=rr, column=4, value=f"=Cost_Lines!D{src}")
    for s in range(3):
        base_col = 5 + s * 4
        vr = VROW[s]
        for t in range(3):
            col = base_col + t
            vol = f"Volume!{chr(ord('C') + t)}{vr}"
            if t == 0:
                f = (f'=IF(Cost_Lines!$D{src}="One-time",Cost_Lines!$F{src}*(1+{R["cont"]}),'
                     f'Cost_Lines!$F{src}+Cost_Lines!$G{src}*{vol}/1000000)')
            else:
                f = (f'=IF(Cost_Lines!$D{src}="One-time",0,'
                     f'Cost_Lines!$F{src}+Cost_Lines!$G{src}*{vol}/1000000)')
            wt.cell(row=rr, column=col, value=f).number_format = USD
        a = wt.cell(row=rr, column=base_col).column_letter
        b = wt.cell(row=rr, column=base_col + 2).column_letter
        wt.cell(row=rr, column=base_col + 3, value=f"=SUM({a}{rr}:{b}{rr})").number_format = USD
T1 = T0 + len(lines) - 1
body(wt, T0, T1, 1, 16)
# option totals by category
rr = T1 + 2
wt.cell(row=rr, column=1, value="Totals by option (option lines + Common lines)").font = SEC
rr += 1
header(wt, rr, ["Option", "Category", "", ""] + hdr[4:])
cats = ["Migration", "Cloud / platform", "Integration", "Security", "Support", "Training", "Change", "Interventions"]
TOT = {}
for opt in ["A", "B"]:
    first = rr + 1
    for cat in cats:
        rr += 1
        wt.cell(row=rr, column=1, value=opt); wt.cell(row=rr, column=2, value=cat)
        for col in range(5, 17):
            L = wt.cell(row=3, column=col).column_letter
            f = (f'=SUMIFS({L}${T0}:{L}${T1},$B${T0}:$B${T1},"{opt}",$C${T0}:$C${T1},"{cat}")'
                 f'+SUMIFS({L}${T0}:{L}${T1},$B${T0}:$B${T1},"Common",$C${T0}:$C${T1},"{cat}")')
            wt.cell(row=rr, column=col, value=f).number_format = USD
    rr += 1
    wt.cell(row=rr, column=1, value=opt); wt.cell(row=rr, column=2, value="TOTAL")
    for col in range(5, 17):
        L = wt.cell(row=3, column=col).column_letter
        wt.cell(row=rr, column=col, value=f"=SUM({L}{first}:{L}{rr-1})").number_format = USD
        wt.cell(row=rr, column=col).font = BOLD; wt.cell(row=rr, column=col).fill = YELLOW
    TOT[opt] = rr
    body(wt, first, rr, 1, 16)
    for col in range(1, 5): wt.cell(row=rr, column=col).font = BOLD
widths(wt, [9, 16, 16, 11] + [12] * 12)
wt.freeze_panes = "E4"

# ------------------------------------------------------------------ Benefits
wbn = wb.create_sheet("Benefits")
title(wbn, "Benefits: per attempt, then by year and scenario (USD)",
      "Per-attempt value x attempts x realisation ramp. Low and high ranges. Fraud loss is a guardrail (held flat), not a benefit.")
header(wbn, 3, ["ID", "Benefit line", "Intervention(s)", "Low USD per attempt", "High USD per attempt", "How it is built", "Stage"])
tech_red_lo = f"({R['o1_lo']}*{R['real_lo']}/{R['tech']})"
tech_red_hi = f"({R['o1_hi']}*{R['real_hi']}/{R['tech']})"
bl = [
    ("BN1", "Recovered payments (fewer technical declines)", "O1, O3",
     f"={R['o1_lo']}*{R['real_lo']}*{R['v_pay']}", f"={R['o1_hi']}*{R['real_hi']}*{R['v_pay']}",
     "O1 upper-bound success gain x realisation x value per recovered payment", "Committed"),
    ("BN2", "Manual reconciliation hours avoided", "O4, O2, O6",
     f"={R['recon_share']}*{R['min_case']}/60*{R['rec_lo']}*{R['rate_h']}", f"={R['recon_share']}*{R['min_case']}/60*{R['rec_hi']}*{R['rate_h']}",
     "Reconciliation hours per attempt x share avoided x loaded hourly cost. Cash only if roles are reduced or redeployed.", "Gated"),
    ("BN3", "Duplicate-payment customer remediation avoided", "O2",
     f"={R['dup']}*{R['dup_lo']}*{R['c_dup']}", f"={R['dup']}*{R['dup_hi']}*{R['c_dup']}",
     "Duplicates per attempt x share removed x remediation cost (handling already in BN2)", "Committed"),
    ("BN4", "Complaint handling avoided", "O6, O1",
     f"={R['compl']}*{R['c_compl']}*({R['compl_pend']}*{R['pend_lo']}+{R['compl_tech']}*{tech_red_lo})",
     f"={R['compl']}*{R['c_compl']}*({R['compl_pend']}*{R['pend_hi']}+{R['compl_tech']}*{tech_red_hi})",
     "Complaints per attempt x cost x (pending share x reduction + technical share x technical-decline reduction)", "Committed"),
    ("BN5", "Fraud-alert handling avoided (conditional on test)", "O5",
     f"={R['alerts']}*{R['c_alert']}*{R['al_lo']}", f"={R['alerts']}*{R['c_alert']}*{R['al_hi']}",
     "Alerts per attempt x cost x alert volume avoided. Zero in the low case: the test may fail.", "Gated"),
]
B0 = 4
for i, b in enumerate(bl):
    rr = B0 + i
    for j, v in enumerate(b):
        c = wbn.cell(row=rr, column=1 + j, value=v)
    wbn.cell(row=rr, column=4).number_format = '$0.0000'
    wbn.cell(row=rr, column=5).number_format = '$0.0000'
B1 = B0 + len(bl) - 1
rr = B1 + 1
wbn.cell(row=rr, column=1, value="TOTAL"); wbn.cell(row=rr, column=2, value="All benefit lines")
wbn.cell(row=rr, column=4, value=f"=SUM(D{B0}:D{B1})").number_format = '$0.0000'
wbn.cell(row=rr, column=5, value=f"=SUM(E{B0}:E{B1})").number_format = '$0.0000'
BTOT = rr
rr += 1
wbn.cell(row=rr, column=1, value="COMMITTED"); wbn.cell(row=rr, column=2, value="Committed-stage lines only (BN1, BN3, BN4)")
wbn.cell(row=rr, column=4, value=f'=SUMIFS(D{B0}:D{B1},G{B0}:G{B1},"Committed")').number_format = '$0.0000'
wbn.cell(row=rr, column=5, value=f'=SUMIFS(E{B0}:E{B1},G{B0}:G{B1},"Committed")').number_format = '$0.0000'
BCOM = rr
body(wbn, B0, rr, 1, 7)
for col in range(1, 8): wbn.cell(row=BTOT, column=col).font = BOLD
wbn.cell(row=BTOT, column=4).fill = YELLOW; wbn.cell(row=BTOT, column=5).fill = YELLOW

rr += 2
wbn.cell(row=rr, column=1, value="Annual benefit by scenario (total of all lines, after realisation ramp)").font = SEC
rr += 1
header(wbn, rr, ["Scenario", "Range", "Year 1", "Year 2", "Year 3", "3-year total", "3-year NPV"])
BEN = {}
for s in range(3):
    for k, colL in (("Low", "D"), ("High", "E")):
        rr += 1
        wbn.cell(row=rr, column=1, value=scen[s][0]); wbn.cell(row=rr, column=2, value=k)
        ramps = [R['ramp1'], R['ramp2'], R['ramp3']]
        for t in range(3):
            vol = f"Volume!{chr(ord('C') + t)}{VROW[s]}"
            wbn.cell(row=rr, column=3 + t, value=f"=${colL}${BTOT}*{vol}*{ramps[t]}").number_format = USD
        wbn.cell(row=rr, column=6, value=f"=SUM(C{rr}:E{rr})").number_format = USD
        wbn.cell(row=rr, column=7, value=f"=C{rr}/(1+{R['disc']})+D{rr}/(1+{R['disc']})^2+E{rr}/(1+{R['disc']})^3").number_format = USD
        BEN[(s, k)] = rr
body(wbn, rr - 5, rr, 1, 7)

rr += 2
wbn.cell(row=rr, column=1, value="Run-rate benefit at full realisation, base Year 1 volume (USD per year)").font = SEC
rr += 1
header(wbn, rr, ["ID", "Benefit line", "Low", "High"])
RUN = {}
for i in range(len(bl)):
    rr += 1
    wbn.cell(row=rr, column=1, value=f"=A{B0+i}"); wbn.cell(row=rr, column=2, value=f"=B{B0+i}")
    wbn.cell(row=rr, column=3, value=f"=D{B0+i}*Volume!$C$4").number_format = USD
    wbn.cell(row=rr, column=4, value=f"=E{B0+i}*Volume!$C$4").number_format = USD
    RUN[bl[i][0]] = rr
rr += 1
wbn.cell(row=rr, column=1, value="TOTAL"); wbn.cell(row=rr, column=3, value=f"=SUM(C{rr-5}:C{rr-1})").number_format = USD
wbn.cell(row=rr, column=4, value=f"=SUM(D{rr-5}:D{rr-1})").number_format = USD
RUN["TOTAL"] = rr
body(wbn, rr - 5, rr, 1, 4)
rr += 2
wbn.cell(row=rr, column=1, value="Operational equivalents at base volume (for the deck)").font = SEC
rr += 1
header(wbn, rr, ["Measure", "Low", "High", "Unit"])
OPS = {}
ops = [
    ("Reconciliation hours avoided per month", f"={R['brief_hours']}*{R['rec_lo']}", f"={R['brief_hours']}*{R['rec_hi']}", "hours / month"),
    ("Success-rate gain (realised)", f"={R['o1_lo']}*{R['real_lo']}", f"={R['o1_hi']}*{R['real_hi']}", "share of attempts"),
    ("Duplicate payments avoided per month", f"={R['dup']}*{R['dup_lo']}*{R['base_m']}", f"={R['dup']}*{R['dup_hi']}*{R['base_m']}", "payments / month"),
    ("Complaints avoided per month", f"={R['compl']}*{R['base_m']}*({R['compl_pend']}*{R['pend_lo']}+{R['compl_tech']}*{tech_red_lo})",
     f"={R['compl']}*{R['base_m']}*({R['compl_pend']}*{R['pend_hi']}+{R['compl_tech']}*{tech_red_hi})", "complaints / month"),
]
for m in ops:
    rr += 1
    for j, v in enumerate(m): wbn.cell(row=rr, column=1 + j, value=v)
    fmt = PCT2 if "share" in m[3] else NUM
    wbn.cell(row=rr, column=2).number_format = fmt; wbn.cell(row=rr, column=3).number_format = fmt
    OPS[m[0]] = rr
body(wbn, rr - 3, rr, 1, 4)
widths(wbn, [12, 46, 16, 16, 16, 60, 14])

# ------------------------------------------------------------------ Summary
wsu = wb.create_sheet("Summary")
title(wsu, "Business case summary (USD, 3 years)",
      "Ranges, not point estimates. Synthetic data and illustrative costs: read the shape of the answer, not the decimals.")
header(wsu, 3, ["Option", "Scenario", "3-yr TCO", "3-yr benefit low", "3-yr benefit high", "Net low", "Net high",
                "NPV of TCO", "NPV net low", "NPV net high"])
rr = 3
SUMROW = {}
for opt in ["A", "B"]:
    for s in range(3):
        rr += 1
        totcol = {0: "H", 1: "L", 2: "P"}[s]
        wsu.cell(row=rr, column=1, value=f"Option {opt}"); wsu.cell(row=rr, column=2, value=scen[s][0])
        wsu.cell(row=rr, column=3, value=f"=TCO!{totcol}{TOT[opt]}")
        wsu.cell(row=rr, column=4, value=f"=Benefits!F{BEN[(s,'Low')]}")
        wsu.cell(row=rr, column=5, value=f"=Benefits!F{BEN[(s,'High')]}")
        wsu.cell(row=rr, column=6, value=f"=D{rr}-C{rr}")
        wsu.cell(row=rr, column=7, value=f"=E{rr}-C{rr}")
        y1, y2, y3 = [wt.cell(row=3, column=5 + s * 4 + t).column_letter for t in range(3)]
        wsu.cell(row=rr, column=8, value=f"=TCO!{y1}{TOT[opt]}/(1+{R['disc']})+TCO!{y2}{TOT[opt]}/(1+{R['disc']})^2+TCO!{y3}{TOT[opt]}/(1+{R['disc']})^3")
        wsu.cell(row=rr, column=9, value=f"=Benefits!G{BEN[(s,'Low')]}-H{rr}")
        wsu.cell(row=rr, column=10, value=f"=Benefits!G{BEN[(s,'High')]}-H{rr}")
        for col in range(3, 11): wsu.cell(row=rr, column=col).number_format = USD
        SUMROW[(opt, s)] = rr
body(wsu, 4, rr, 1, 10)

# breakeven block
rr += 2
wsu.cell(row=rr, column=1, value="Breakeven: monthly attempts at which 3-year benefits equal 3-year TCO (base growth)").font = SEC
rr += 1
header(wsu, rr, ["Option", "Fixed cost over 3 yrs (F)", "Variable cost per attempt (v)", "Benefit per attempt low", "Benefit per attempt high",
                 "Breakeven monthly attempts (high benefit)", "Breakeven monthly attempts (low benefit)", "Base monthly attempts", "Pays back at base volume?"])
BE = {}
for opt in ["A", "B"]:
    rr += 1
    wsu.cell(row=rr, column=1, value=f"Option {opt}")
    # F: one-time with contingency + 3 x fixed recurring, option + common
    fF = (f'=SUMPRODUCT(((Cost_Lines!$B$4:$B${CL1}="{opt}")+(Cost_Lines!$B$4:$B${CL1}="Common"))*(Cost_Lines!$D$4:$D${CL1}="One-time")*Cost_Lines!$F$4:$F${CL1})*(1+{R["cont"]})'
          f'+3*SUMPRODUCT(((Cost_Lines!$B$4:$B${CL1}="{opt}")+(Cost_Lines!$B$4:$B${CL1}="Common"))*(Cost_Lines!$D$4:$D${CL1}="Recurring")*Cost_Lines!$F$4:$F${CL1})')
    fv = (f'=SUMPRODUCT(((Cost_Lines!$B$4:$B${CL1}="{opt}")+(Cost_Lines!$B$4:$B${CL1}="Common"))*(Cost_Lines!$D$4:$D${CL1}="Recurring")*Cost_Lines!$G$4:$G${CL1})/1000000')
    wsu.cell(row=rr, column=2, value=fF).number_format = USD
    wsu.cell(row=rr, column=3, value=fv).number_format = '$0.0000'
    wsu.cell(row=rr, column=4, value=f"=Benefits!D{BTOT}").number_format = '$0.0000'
    wsu.cell(row=rr, column=5, value=f"=Benefits!E{BTOT}").number_format = '$0.0000'
    wsu.cell(row=rr, column=6, value=f'=IFERROR(B{rr}/(12*(E{rr}*Volume!$B$10-C{rr}*Volume!$B$9)),"never")').number_format = NUM
    wsu.cell(row=rr, column=7, value=f'=IFERROR(B{rr}/(12*(D{rr}*Volume!$B$10-C{rr}*Volume!$B$9)),"never")').number_format = NUM
    wsu.cell(row=rr, column=8, value=f"={R['base_m']}").number_format = NUM
    wsu.cell(row=rr, column=9, value=f'=IF(H{rr}>=G{rr},"Yes, even on low benefits",IF(H{rr}>=F{rr},"Only on high benefits","No: below breakeven"))')
    BE[opt] = rr
    for col in (6, 7): wsu.cell(row=rr, column=col).fill = YELLOW
body(wsu, rr - 1, rr, 1, 9)
rr += 1
wsu.cell(row=rr, column=1, value="Breakeven formula: F / (12 x (benefit per attempt x Rr - v x S)), with S and Rr from the Volume sheet. Solves 3-yr benefit = 3-yr TCO for a monthly volume M growing at the base rate.").font = SUB

# A vs B crossover
rr += 2
wsu.cell(row=rr, column=1, value="Option A vs Option B: when does cloud-native become cheaper?").font = SEC
rr += 1
wsu.cell(row=rr, column=1, value="Monthly attempts above which A's 3-yr TCO is below B's")
wsu.cell(row=rr, column=6, value=f'=IFERROR((B{BE["A"]}-B{BE["B"]})/(12*Volume!$B$9*(C{BE["B"]}-C{BE["A"]})),"never")').number_format = NUM
wsu.cell(row=rr, column=6).fill = YELLOW
CROSS = rr
body(wsu, rr, rr, 1, 6)

# alternative scale anchor
rr += 2
wsu.cell(row=rr, column=1, value="Alternative scale anchor (monthly attempts in Inputs) - base growth, 3 years").font = SEC
rr += 1
header(wsu, rr, ["Option", "Monthly attempts", "3-yr TCO", "3-yr benefit low", "3-yr benefit high", "Net low", "Net high", "Note"])
ALT = {}
for opt in ["A", "B"]:
    rr += 1
    wsu.cell(row=rr, column=1, value=f"Option {opt}")
    wsu.cell(row=rr, column=2, value=f"={R['alt_m']}").number_format = NUM
    wsu.cell(row=rr, column=3, value=f"=B{BE[opt]}+C{BE[opt]}*12*B{rr}*Volume!$B$9").number_format = USD
    wsu.cell(row=rr, column=4, value=f"=D{BE[opt]}*12*B{rr}*Volume!$B$10").number_format = USD
    wsu.cell(row=rr, column=5, value=f"=E{BE[opt]}*12*B{rr}*Volume!$B$10").number_format = USD
    wsu.cell(row=rr, column=6, value=f"=D{rr}-C{rr}").number_format = USD
    wsu.cell(row=rr, column=7, value=f"=E{rr}-C{rr}").number_format = USD
    wsu.cell(row=rr, column=8, value="At this volume the brief's 1,200 hours would imply ~2 minutes per case, so real hours (and BN2) would be higher: the case is conservative.")
    ALT[opt] = rr
body(wsu, rr - 1, rr, 1, 8)

# retention value needed
rr += 2
wsu.cell(row=rr, column=1, value="What a recovered payment would need to be worth for Option B to break even at base volume (all other benefits at high)").font = SEC
rr += 1
wsu.cell(row=rr, column=1, value="Required value per recovered payment (USD)")
rec_pay = f"({R['o1_hi']}*{R['real_hi']})"
wsu.cell(row=rr, column=6, value=f"=(C{SUMROW[('B',0)]}-(E{SUMROW[('B',0)]}-Benefits!E{B0}*Volume!$B$10*12*{R['base_m']}))/({rec_pay}*12*{R['base_m']}*Volume!$B$10)").number_format = '$#,##0.00'
wsu.cell(row=rr, column=7, value=f"=F{rr}/{R['v_pay']}").number_format = '0"x assumed value"'
wsu.cell(row=rr, column=8, value="Shows that success-rate gains alone cannot carry the spend at this volume; the case needs reconciliation savings, scale, or retention value.")
RETV = rr
body(wsu, rr, rr, 1, 8)

# cost per successful payment
rr += 2
wsu.cell(row=rr, column=1, value="Cost per successful payment (K8, fully loaded view), base scenario Year 3").font = SEC
rr += 1
header(wsu, rr, ["View", "Handling cost per attempt", "Platform + intervention run cost per attempt", "Success rate", "Cost per successful payment", "Note"])
rr += 1
wsu.cell(row=rr, column=1, value="Today (handling cost only, measured)")
wsu.cell(row=rr, column=2, value=f"={R['hcost']}")
wsu.cell(row=rr, column=3, value=0)
wsu.cell(row=rr, column=4, value=f"={R['succ']}")
wsu.cell(row=rr, column=5, value=f"=(B{rr}+C{rr})/D{rr}")
wsu.cell(row=rr, column=6, value="Matches KPI K8 in the analysis (USD 0.336)")
K8_0 = rr
for opt in ["A", "B"]:
    rr += 1
    wsu.cell(row=rr, column=1, value=f"Option {opt}, Year 3, low benefits")
    wsu.cell(row=rr, column=2, value=f"={R['hcost']}-Benefits!D{B0+1}-Benefits!D{B0+3}")
    wsu.cell(row=rr, column=3, value=f"=TCO!G{TOT[opt]}/Volume!E4")
    wsu.cell(row=rr, column=4, value=f"={R['succ']}+{R['o1_lo']}*{R['real_lo']}")
    wsu.cell(row=rr, column=5, value=f"=(B{rr}+C{rr})/D{rr}")
    wsu.cell(row=rr, column=6, value="Includes platform and intervention run cost per attempt; compare with today. Falls as volume grows.")
for r2 in range(K8_0, rr + 1):
    for col in (2, 3, 5): wsu.cell(row=r2, column=col).number_format = '$0.000'
    wsu.cell(row=r2, column=4).number_format = PCT2
body(wsu, K8_0, rr, 1, 6)
widths(wsu, [34, 16, 16, 16, 16, 18, 18, 16, 22, 16])
wsu.freeze_panes = "A4"


# ------------------------------------------------------------------ Intervention economics
we = wb.create_sheet("Intervention_Econ")
title(we, "Economics of each intervention on its own (base scenario, 3 years, USD)",
      "Hard benefits only. BN2 is attributed to O4; BN4 is split between O6 (pending complaints) and O1 (technical complaints). O3 alerting sits inside the event-feed build (B-I1) and has no direct cash benefit.")
header(we, 3, ["ID", "Intervention", "Cost lines", "3-yr cost (incl. contingency and run)", "3-yr hard benefit low", "3-yr hard benefit high",
               "Net low", "Net high", "Pays back on cash?", "Main non-cash value"])
def costf(ids):
    parts = [f'SUMIFS(TCO!$H${T0}:$H${T1},TCO!$A${T0}:$A${T1},"{i}")' for i in ids]
    return "=" + "+".join(parts)
vol3 = f"12*{R['base_m']}*Volume!$B$10"
pend_lo = f"{R['compl']}*{R['c_compl']}*{R['compl_pend']}*{R['pend_lo']}"
pend_hi = f"{R['compl']}*{R['c_compl']}*{R['compl_pend']}*{R['pend_hi']}"
techc_lo = f"{R['compl']}*{R['c_compl']}*{R['compl_tech']}*{tech_red_lo}"
techc_hi = f"{R['compl']}*{R['c_compl']}*{R['compl_tech']}*{tech_red_hi}"
ie = [
    ("O1", "Processor routing + PROC_C peak failover", ["X-O1"], f"=(Benefits!D{B0}+{techc_lo})*{vol3}", f"=(Benefits!E{B0}+{techc_hi})*{vol3}",
     "Fewer failed payments in the PROC_C evening peak (7.2% vs 2.0% technical declines)"),
    ("O2", "Safe retry (idempotency, no retry while pending)", ["X-O2"], f"=Benefits!D{B0+2}*{vol3}", f"=Benefits!E{B0+2}*{vol3}",
     "Removes double charges (21% of pending retries flagged as duplicates): conduct and trust risk"),
    ("O3", "Real-time failure and pending alerts", ["B-I1"], "=0", "=0",
     "Enabler: detection in minutes, triggers O1 failover, makes control limits operable"),
    ("O4", "Automated reconciliation", ["X-O4", "X-R1"], f"=Benefits!D{B0+1}*{vol3}", f"=Benefits!E{B0+1}*{vol3}",
     "Frees 360-720 staff-hours a month; faster case closure (SLA K6)"),
    ("O5", "Fraud-rule tuning via controlled test", ["X-O5"], f"=Benefits!D{B0+4}*{vol3}", f"=Benefits!E{B0+4}*{vol3}",
     "Fewer genuine customers declined (98.3% of alerts not fraud; 20% complain after a risk decline)"),
    ("O6", "Customer status notifications (pending > 15 min)", ["X-O6", "X-P1"], f"={pend_lo}*{vol3}", f"={pend_hi}*{vol3}",
     "Certainty for customers; fewer anxious retries that create duplicates"),
]
rr = 3
IE = {}
for code, name, ids, blo, bhi, note in ie:
    rr += 1
    we.cell(row=rr, column=1, value=code); we.cell(row=rr, column=2, value=name); we.cell(row=rr, column=3, value=", ".join(ids))
    we.cell(row=rr, column=4, value=costf(ids))
    we.cell(row=rr, column=5, value=blo); we.cell(row=rr, column=6, value=bhi)
    we.cell(row=rr, column=7, value=f"=E{rr}-D{rr}"); we.cell(row=rr, column=8, value=f"=F{rr}-D{rr}")
    we.cell(row=rr, column=9, value=f'=IF(G{rr}>=0,"Yes",IF(H{rr}>=0,"Only on high case","No"))')
    we.cell(row=rr, column=10, value=note)
    for col in range(4, 9): we.cell(row=rr, column=col).number_format = USD
    IE[code] = rr
body(we, 4, rr, 1, 10)
rr += 1
we.cell(row=rr, column=1, value="Note: O3 cost shown is the shared event-feed build (B-I1), which every other intervention also relies on. Benefits use the base-scenario ramp; they are hard savings only.").font = SUB
rr += 2
we.cell(row=rr, column=1, value="Gate G2: monthly manual reconciliation hours at which O4 pays back in 3 years").font = SEC
rr += 1
header(we, rr, ["Case", "Share of hours avoided", "O4 3-yr cost", "Breakeven hours per month (today's level, growing with volume)", "Brief's illustrative hours", "Passes?"])
G2 = {}
for k, sh in (("Low automation", R['rec_lo']), ("High automation", R['rec_hi'])):
    rr += 1
    we.cell(row=rr, column=1, value=k); we.cell(row=rr, column=2, value=f"={sh}").number_format = '0%'
    we.cell(row=rr, column=3, value=f"=D{IE['O4']}").number_format = USD
    we.cell(row=rr, column=4, value=f"=C{rr}/(12*(Volume!$B$10)*B{rr}*{R['rate_h']})").number_format = NUM
    we.cell(row=rr, column=5, value=f"={R['brief_hours']}").number_format = NUM
    we.cell(row=rr, column=6, value=f'=IF(E{rr}>=D{rr},"Yes","No")')
    we.cell(row=rr, column=4).fill = YELLOW
    G2[k] = rr
body(we, rr - 1, rr, 1, 6)
widths(we, [7, 40, 14, 18, 16, 16, 14, 14, 16, 70])

# ------------------------------------------------------------------ Stage gates
wg = wb.create_sheet("Stage_Gates")
title(wg, "Committed vs gated spend (Option B, base scenario)",
      "Committed = approved now. Gated = released only if the gate criterion is met. Keeps the bank from paying for the expensive parts before the evidence exists.")
header(wg, 3, ["Stage", "3-yr cost (Option B, base)", "3-yr benefit low", "3-yr benefit high", "Gate", "Release criterion (measurable)", "Decision owner", "Decision point"])
rr = 4
wg.cell(row=rr, column=1, value="Committed")
wg.cell(row=rr, column=2, value=f'=SUMIFS(TCO!H${T0}:H${T1},TCO!B${T0}:B${T1},"B",Cost_Lines!H${CL0}:H${CL1},"Committed")+SUMIFS(TCO!H${T0}:H${T1},TCO!B${T0}:B${T1},"Common",Cost_Lines!H${CL0}:H${CL1},"Committed")')
wg.cell(row=rr, column=3, value=f"=Benefits!D{BCOM}*12*{R['base_m']}*Volume!$B$10")
wg.cell(row=rr, column=4, value=f"=Benefits!E{BCOM}*12*{R['base_m']}*Volume!$B$10")
wg.cell(row=rr, column=5, value="G0 (now)")
wg.cell(row=rr, column=6, value="Approved now: wave-1 fixes (O1, O2, O3, O6) and the event-level data foundation. Justified by customer harm and control, not by payback.")
wg.cell(row=rr, column=7, value="COO with CFO")
wg.cell(row=rr, column=8, value="Month 0")
rr += 1
wg.cell(row=rr, column=1, value="Gated")
wg.cell(row=rr, column=2, value=f'=SUMIFS(TCO!H${T0}:H${T1},TCO!B${T0}:B${T1},"B",Cost_Lines!H${CL0}:H${CL1},"Gated")+SUMIFS(TCO!H${T0}:H${T1},TCO!B${T0}:B${T1},"Common",Cost_Lines!H${CL0}:H${CL1},"Gated")')
wg.cell(row=rr, column=3, value=f"=(Benefits!D{BTOT}-Benefits!D{BCOM})*12*{R['base_m']}*Volume!$B$10")
wg.cell(row=rr, column=4, value=f"=(Benefits!E{BTOT}-Benefits!E{BCOM})*12*{R['base_m']}*Volume!$B$10")
wg.cell(row=rr, column=5, value="G1, G2, G3")
wg.cell(row=rr, column=6, value="G1 (curated layer and KPI marts): real monthly volume measured and the platform choice confirmed. G2 (O4 automated reconciliation): 4 weeks of time-tracking confirm manual reconciliation hours at or above the O4 breakeven in Intervention_Econ (about 1,100 hours a month at 30% automation). G3 (O5 fraud tuning): CRO ceiling agreed and reviewed-outcome labels flowing.")
wg.cell(row=rr, column=7, value="CIO (G1), COO (G2), CRO (G3)")
wg.cell(row=rr, column=8, value="Month 3 (G1, G2), Month 6 (G3)")
rr += 1
wg.cell(row=rr, column=1, value="Total")
for col in (2, 3, 4):
    L = wg.cell(row=3, column=col).column_letter
    wg.cell(row=rr, column=col, value=f"=SUM({L}4:{L}5)")
for r2 in range(4, rr + 1):
    for col in (2, 3, 4): wg.cell(row=r2, column=col).number_format = USD
body(wg, 4, rr, 1, 8)
rr += 1
wg.cell(row=rr, column=1, value="Check: total equals Summary Option B base TCO")
wg.cell(row=rr, column=2, value=f'=IF(ABS(B{rr-1}-Summary!C{SUMROW[("B",0)]})<1,"OK","MISMATCH")')
rr += 2
wg.cell(row=rr, column=1, value="Option A migration gate (G4)").font = SEC
rr += 1
wg.cell(row=rr, column=1, value="Move from Option B to Option A only when monthly attempts exceed the A/B crossover in Summary, or when on-prem capacity refresh is due, whichever comes first. Open table formats in Option B keep this move cheap.")
wg.cell(row=rr, column=1).alignment = WRAP
wg.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=8)
wg.row_dimensions[rr].height = 30
widths(wg, [12, 18, 16, 16, 12, 70, 18, 22])

wb.save(OUT)
import json
json.dump({"R": R, "SUMROW": {f"{k[0]}{k[1]}": v for k, v in SUMROW.items()}, "BE": BE, "CROSS": CROSS, "ALT": ALT,
           "RETV": RETV, "IE": IE, "G2": G2, "RUN": RUN, "OPS": OPS, "BTOT": BTOT, "BCOM": BCOM, "TOT": TOT, "K8_0": K8_0, "B0": B0},
          open("tco_map.json", "w"), indent=1)
print("saved", OUT)
