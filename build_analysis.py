import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

df = pd.read_pickle("/home/claude/payments.pkl")
df = df.sort_values(["Transaction_ID", "timestamp"]).reset_index(drop=True)
N = len(df); LAST = N + 1

OUT = "/mnt/user-data/outputs/Capstone1_Payments_Failure_Reduction/03_Data_and_Analysis/Payments_Analysis_Workbook.xlsx"
F = "Arial"
NAVY = "1F3864"
hf = Font(name=F, bold=True, color="FFFFFF", size=10)
hfill = PatternFill("solid", start_color=NAVY)
body = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
blue = Font(name=F, size=10, color="0000FF")
green = Font(name=F, size=10, color="008000")
title = Font(name=F, size=13, bold=True, color=NAVY)
note = Font(name=F, size=9, italic=True, color="595959")
wrap = Alignment(wrap_text=True, vertical="top")
yellow = PatternFill("solid", start_color="FFFF00")
sect = PatternFill("solid", start_color="D9E2F3")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
PCT = "0.0%"; PCT2 = "0.00%"; INT = "#,##0"; DEC = "0.00"

wb = Workbook()

# --------------------------------------------------------------- Data sheet
wsD = wb.active; wsD.title = "Data"
orig_cols = list(df.columns)
helpers = ["Dup_ID_repeat", "In_Scope", "Pending_gt15", "Proc_clean", "Fail_Code_clean", "Inv_hours_raw", "Inv_valid",
           "Inv_hours_valid", "SLA_hours", "In_SLA", "Inv_h_Mobile", "Inv_h_Web", "Inv_h_Branch", "Latency_outlier", "Recon_cause"]
wsD.append(orig_cols + helpers)
LET = {name: get_column_letter(i + 1) for i, name in enumerate(orig_cols + helpers)}
assert LET["Transaction_ID"] == "A" and LET["customer_complaint_flag"] == "AA" and LET["Dup_ID_repeat"] == "AB"

for i, rec in enumerate(df.itertuples(index=False), start=2):
    row = []
    for v in rec:
        if v is None or (not isinstance(v, str) and pd.isna(v)) or v == "":
            row.append(None)
        elif isinstance(v, pd.Timestamp):
            row.append(v.to_pydatetime(warn=False))
        elif hasattr(v, "item"):
            row.append(v.item())
        else:
            row.append(v)
    r = i
    L = LET
    h = [
        0 if r == 2 else f"=IF({L['Transaction_ID']}{r}={L['Transaction_ID']}{r-1},1,0)",
        f"=1-{L['Dup_ID_repeat']}{r}",
        f"=IF({L['pending_minutes']}{r}>15,1,0)",
        f'=IF({L["processor"]}{r}="","Unknown",{L["processor"]}{r})',
        f'=IF({L["payment_status"]}{r}="SUCCESS","",IF({L["failure_code"]}{r}="","Unclassified",{L["failure_code"]}{r}))',
        f'=IF(AND({L["investigation_start"]}{r}<>"",{L["investigation_end"]}{r}<>""),({L["investigation_end"]}{r}-{L["investigation_start"]}{r})*24,"")',
        f'=IF({L["Inv_hours_raw"]}{r}="","",IF({L["Inv_hours_raw"]}{r}>0,1,0))',
        f'=IF(AND({L["Inv_valid"]}{r}=1,{L["In_Scope"]}{r}=1),{L["Inv_hours_raw"]}{r},"")',
        f"=IF({L['fraud_alert_flag']}{r}=1,Inputs!$B$3,Inputs!$B$4)",
        f'=IF({L["Inv_valid"]}{r}=1,IF({L["Inv_hours_raw"]}{r}<={L["SLA_hours"]}{r},1,0),"")',
        f'=IF(AND({L["Inv_hours_valid"]}{r}<>"",{L["channel"]}{r}="Mobile"),{L["Inv_hours_valid"]}{r},"")',
        f'=IF(AND({L["Inv_hours_valid"]}{r}<>"",{L["channel"]}{r}="Web"),{L["Inv_hours_valid"]}{r},"")',
        f'=IF(AND({L["Inv_hours_valid"]}{r}<>"",{L["channel"]}{r}="Branch_Assisted"),{L["Inv_hours_valid"]}{r},"")',
        f'=IF({L["latency_ms"]}{r}="",0,IF({L["latency_ms"]}{r}>Inputs!$B$5,1,0))',
        f'=IF({L["reconciliation_flag"]}{r}=1,IF({L["payment_status"]}{r}="PENDING","Pending",IF({L["duplicate_suspected_flag"]}{r}=1,"Duplicate",IF({L["payment_status"]}{r}="FAILED_TECHNICAL","Technical failure","Other"))),"")',
    ]
    wsD.append(row + h)
for c in wsD[1]:
    c.font = hf; c.fill = hfill; c.alignment = Alignment(wrap_text=True, vertical="center")
for i in range(1, len(orig_cols) + len(helpers) + 1):
    wsD.column_dimensions[get_column_letter(i)].width = 16
for col in ("B", "W", "X"):
    wsD.column_dimensions[col].width = 18
    for r in range(2, LAST + 1):
        wsD[f"{col}{r}"].number_format = "yyyy-mm-dd hh:mm"
wsD.freeze_panes = "B2"
wsD.auto_filter.ref = f"A1:{get_column_letter(len(orig_cols) + len(helpers))}{LAST}"

def R(name):
    c = LET[name]
    return f"Data!${c}$2:${c}${LAST}"

SC = R("In_Scope")
def cif(*pairs):
    parts = [f"{R(n)},{c}" for n, c in pairs] + [f"{SC},1"]
    return "COUNTIFS(" + ",".join(parts) + ")"
def sif(sumcol, *pairs):
    parts = [f"{R(n)},{c}" for n, c in pairs] + [f"{SC},1"]
    return f"SUMIFS({R(sumcol)}," + ",".join(parts) + ")"
TOTAL = f"SUM({SC})"

def hdr(ws, row, labels, start=1):
    for i, t in enumerate(labels, start):
        c = ws.cell(row=row, column=i, value=t); c.font = hf; c.fill = hfill
        c.alignment = Alignment(wrap_text=True, vertical="center")
def put(ws, addr, v, fmt=None, font=body, fill=None):
    c = ws[addr]; c.value = v; c.font = font; c.alignment = wrap; c.border = box
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    return c
def widths(ws, ws_w):
    for i, w in enumerate(ws_w, 1): ws.column_dimensions[get_column_letter(i)].width = w

# --------------------------------------------------------------- Method
wsM = wb.create_sheet("Method", 0)
M = [
    ("Workstream 3 - Analysis workbook (synthetic data)", title),
    ("Draft v1, 11 Nov 2026. All figures come from the synthetic dataset (12,000 records, 1 Jul - 30 Sep 2026). They show what the method finds, not facts about a real bank.", note),
    ("", body),
    ("How this workbook works", bold),
    ("Data: the 12,000 transactions plus 15 helper columns that are formulas (blue-header columns AB onward). Rows are sorted by Transaction_ID so repeated IDs sit next to each other.", body),
    ("De-duplication: the second and later rows with the same Transaction_ID are flagged (Dup_ID_repeat) and excluded through In_Scope. Every KPI counts In_Scope = 1 rows only. DataQuality shows the effect of this choice.", body),
    ("Inputs: thresholds, tolerance and the brief's illustrative case facts are in the Inputs sheet (blue font). Change them and everything recalculates.", body),
    ("Statistics: rates carry normal-approximation 95% intervals. Group comparisons use a two-proportion z-test (two-sided). Alpha is an input (default 5%).", body),
    ("Verdicts: each hypothesis sheet block ends with a formula-driven verdict. A verdict is 'supported in this data' or 'not supported', never a claim about a real bank.", body),
    ("", body),
    ("Sheets", bold),
    ("Inputs | CaseFacts (tests the brief's 5 illustrative facts) | OutcomeMix (outcome decomposition) | Segments (cuts by channel, processor, country, payment type, time band, device, segment, merchant) | FailureCodes | HypTests (the 10 hypotheses) | DataQuality | Opportunity (indicative, upper-bound) | Findings", body),
    ("", body),
    ("Colour code: blue font = input or assumption; black = formula; yellow fill = key assumption or verdict.", body),
]
for i, (t, f) in enumerate(M, 1):
    c = wsM.cell(row=i, column=1, value=t); c.font = f; c.alignment = wrap
wsM.column_dimensions["A"].width = 150

# --------------------------------------------------------------- Inputs
wsI = wb.create_sheet("Inputs", 1)
put(wsI, "A1", "Inputs and assumptions", font=title); wsI["A1"].border = Border()
hdr(wsI, 2, ["Item", "Value", "Note / source"])
inp = [
    (3, "SLA hours: fraud-alert cases", 8, "PROPOSED (assumption A-009), not agreed", None),
    (4, "SLA hours: reconciliation / complaint cases", 24, "PROPOSED (assumption A-009), not agreed", None),
    (5, "Latency outlier cap (ms)", 20000, "Analyst threshold; synthetic data has 25-90 s outliers", INT),
    (6, "Significance level (alpha)", 0.05, "Two-sided tests", "0.00"),
    (7, "z critical value for 95% intervals", 1.96, "Normal approximation", "0.00"),
    (8, "B2 threshold: 'large' share of business declines in non-success", 0.30, "Analyst definition of 'large'", PCT),
    (9, "D1 threshold: 'large' share of alerts not confirmed as fraud", 0.50, "Analyst definition of 'large'", PCT),
    (10, "Tolerance for median investigation hours (hours)", 0.5, "Analyst threshold for 'supported'", "0.0"),
    (11, "Tolerance for fraud loss rate (absolute)", 0.0001, "1 basis point; analyst threshold", "0.00%"),
]
for r, a, v, n, fmt in inp:
    put(wsI, f"A{r}", a); put(wsI, f"B{r}", v, fmt, font=blue, fill=yellow); put(wsI, f"C{r}", n)
put(wsI, "A13", "Illustrative case facts from the brief (NOT real client results)", font=bold, fill=sect)
put(wsI, "B13", None, fill=sect); put(wsI, "C13", None, fill=sect)
facts = [
    (14, "Payment success rate", 0.942, PCT), (15, "Pending beyond 15 minutes (share of transactions)", 0.018, PCT),
    (16, "Median investigation time (hours)", 9, "0.0"), (17, "Fraud loss (share of transaction value)", 0.0007, "0.00%"),
    (18, "Manual reconciliation (staff-hours per month)", 1200, INT),
]
for r, a, v, fmt in facts:
    put(wsI, f"A{r}", a); put(wsI, f"B{r}", v, fmt, font=blue); put(wsI, f"C{r}", "Brief, 'Illustrative case facts'")
put(wsI, "A20", "Scenario assumptions (analyst assumptions, replace with bank data)", font=bold, fill=sect)
put(wsI, "B20", None, fill=sect); put(wsI, "C20", None, fill=sect)
scen = [
    (21, "Reconciliation cases avoided by automation: low", 0.30, "ASSUMPTION for sensitivity only", PCT),
    (22, "Reconciliation cases avoided by automation: high", 0.60, "ASSUMPTION for sensitivity only", PCT),
    (23, "Bank monthly payment attempts: scenario 1", 100000, "ASSUMPTION; real volume unknown", INT),
    (24, "Bank monthly payment attempts: scenario 2", 500000, "ASSUMPTION; real volume unknown", INT),
    (25, "Bank monthly payment attempts: scenario 3", 1000000, "ASSUMPTION; real volume unknown", INT),
]
for r, a, v, n, fmt in scen:
    put(wsI, f"A{r}", a); put(wsI, f"B{r}", v, fmt, font=blue, fill=yellow); put(wsI, f"C{r}", n)
widths(wsI, [62, 16, 60])

# --------------------------------------------------------------- CaseFacts
ws = wb.create_sheet("CaseFacts", 2)
put(ws, "A1", "Testing the brief's illustrative case facts against the data", font=title); ws["A1"].border = Border()
put(ws, "A2", "Interval = 95% interval for rates; interquartile range for the median; not computed for value ratios. Verdict uses the tolerance in Inputs.", font=note); ws["A2"].border = Border()
hdr(ws, 4, ["Fact", "Case value", "Measured (in scope)", "Measured (all rows)", "Gap (measured - case)", "n", "Interval low", "Interval high", "Case value inside interval?", "Verdict"])
ZC = "Inputs!$B$7"
def rate_row(r, label, case_ref, num_scope, num_all, fmt, tol_ref=None):
    put(ws, f"A{r}", label, font=bold); put(ws, f"B{r}", f"={case_ref}", fmt, font=green)
    put(ws, f"C{r}", f"={num_scope}/{TOTAL}", fmt)
    put(ws, f"D{r}", f"={num_all}/COUNTA({R('Transaction_ID')})", fmt)
    put(ws, f"E{r}", f"=C{r}-B{r}", fmt)
    put(ws, f"F{r}", f"={TOTAL}", INT)
    put(ws, f"G{r}", f"=C{r}-{ZC}*SQRT(C{r}*(1-C{r})/F{r})", fmt)
    put(ws, f"H{r}", f"=C{r}+{ZC}*SQRT(C{r}*(1-C{r})/F{r})", fmt)
    put(ws, f"I{r}", f'=IF(AND(B{r}>=G{r},B{r}<=H{r}),"Yes","No")')
    put(ws, f"J{r}", f'=IF(I{r}="Yes","Supported: case value is consistent with the data","Challenged: data differs from the case value by "&TEXT(ABS(E{r})*100,"0.00")&" percentage points")', fill=yellow)
rate_row(5, "1. Payment success rate", "Inputs!$B$14", cif(("payment_status", '"SUCCESS"')), f"COUNTIF({R('payment_status')},\"SUCCESS\")", PCT)
rate_row(6, "2. Pending beyond 15 minutes (share of attempts)", "Inputs!$B$15", cif(("Pending_gt15", "1")), f"SUM({R('Pending_gt15')})", PCT2)
# 3 median investigation
put(ws, "A7", "3. Median investigation time (hours)", font=bold); put(ws, "B7", "=Inputs!$B$16", "0.0", font=green)
put(ws, "C7", f"=MEDIAN({R('Inv_hours_valid')})", "0.00")
put(ws, "D7", f'=MEDIAN({R("Inv_hours_raw")})', "0.00")
put(ws, "E7", "=C7-B7", "0.00"); put(ws, "F7", f"=COUNT({R('Inv_hours_valid')})", INT)
put(ws, "G7", f"=PERCENTILE({R('Inv_hours_valid')},0.25)", "0.00"); put(ws, "H7", f"=PERCENTILE({R('Inv_hours_valid')},0.75)", "0.00")
put(ws, "I7", '=IF(AND(B7>=G7,B7<=H7),"Yes (inside IQR)","No")')
put(ws, "J7", '=IF(ABS(E7)<=Inputs!$B$10,"Supported: within "&Inputs!$B$10&" hours of the case value","Challenged: data differs by "&TEXT(ABS(E7),"0.00")&" hours")', fill=yellow)
# 4 fraud loss two bases
for r, label, den_scope in ((8, "4a. Fraud loss rate (value of ALL attempts, like-for-like)", sif("amount_usd_equiv")),
                            (9, "4b. Fraud loss rate (value of APPROVED payments, KPI K4)", sif("amount_usd_equiv", ("payment_status", '"SUCCESS"')))):
    put(ws, f"A{r}", label, font=bold); put(ws, f"B{r}", "=Inputs!$B$17", PCT2 if False else "0.000%", font=green)
    put(ws, f"C{r}", f"={sif('fraud_loss_usd')}/{den_scope}", "0.000%")
    put(ws, f"D{r}", f"=SUM({R('fraud_loss_usd')})/" + (f"SUM({R('amount_usd_equiv')})" if r == 8 else f"SUMIFS({R('amount_usd_equiv')},{R('payment_status')},\"SUCCESS\")"), "0.000%")
    put(ws, f"E{r}", f"=C{r}-B{r}", "0.000%"); put(ws, f"F{r}", f"={TOTAL}", INT)
    put(ws, f"G{r}", "n/a"); put(ws, f"H{r}", "n/a"); put(ws, f"I{r}", "n/a")
    put(ws, f"J{r}", f'=IF(ABS(E{r})<=Inputs!$B$11,"Supported on the absolute tolerance (within "&TEXT(Inputs!$B$11*10000,"0")&" bp), but measured is "&TEXT(ABS(E{r})/B{r},"0%")&" "&IF(E{r}<0,"below","above")&" the case value in relative terms","Challenged: data differs by "&TEXT(ABS(E{r})*10000,"0.0")&" basis points")', fill=yellow)
# 5 reconciliation
put(ws, "A10", "5. Manual reconciliation: 1,200 staff-hours per month", font=bold); put(ws, "B10", "=Inputs!$B$18", INT, font=green)
for col in "CDEFGHI": put(ws, f"{col}10", "n/a")
put(ws, "J10", "Cannot be tested from transaction rows. See the implied-minutes table below: it needs real monthly volume and time-tracking.", fill=yellow)

put(ws, "A13", "Reconciliation workload check (assumption A-004)", font=bold, fill=sect)
for col in "BCDE": put(ws, f"{col}13", None, fill=sect)
put(ws, "A14", "Share of attempts needing manual reconciliation (in data)"); put(ws, "B14", f"={cif(('reconciliation_flag','1'))}/{TOTAL}", PCT2)
hdr(ws, 15, ["Monthly attempts (scenario)", "Cases per month", "Implied minutes per case to equal 1,200 hours", "Reading"])
for i, ref in enumerate(("Inputs!$B$23", "Inputs!$B$24", "Inputs!$B$25")):
    r = 16 + i
    put(ws, f"A{r}", f"={ref}", INT, font=green); put(ws, f"B{r}", f"=A{r}*$B$14", INT)
    put(ws, f"C{r}", f"=Inputs!$B$18*60/B{r}", "0.0")
    put(ws, f"D{r}", f'=IF(C{r}<5,"Implausibly short per case: 1,200 hours would be too low for this volume",IF(C{r}>120,"Implausibly long per case: 1,200 hours would be too high for this volume","Plausible handling time"))')
put(ws, "A20", "Reading: the 1,200 hours only makes sense for a specific monthly volume and handling time. Treat as an assumption and show the working in the case pack.", font=note); ws["A20"].border = Border()
widths(ws, [52, 14, 18, 18, 18, 10, 14, 14, 20, 70])
ws.freeze_panes = "B5"

# --------------------------------------------------------------- OutcomeMix
ws = wb.create_sheet("OutcomeMix", 3)
put(ws, "A1", "Outcome decomposition (shares of in-scope attempts)", font=title); ws["A1"].border = Border()
hdr(ws, 3, ["Outcome", "Attempts", "Share", "Interval low", "Interval high"])
outs = ["SUCCESS", "FAILED_TECHNICAL", "FAILED_BUSINESS", "DECLINED_RISK", "PENDING"]
for i, o in enumerate(outs):
    r = 4 + i
    put(ws, f"A{r}", o, font=bold); put(ws, f"B{r}", f"={cif(('payment_status', chr(34)+o+chr(34)))}", INT)
    put(ws, f"C{r}", f"=B{r}/$B$9", PCT2)
    put(ws, f"D{r}", f"=C{r}-{ZC}*SQRT(C{r}*(1-C{r})/$B$9)", PCT2); put(ws, f"E{r}", f"=C{r}+{ZC}*SQRT(C{r}*(1-C{r})/$B$9)", PCT2)
put(ws, "A9", "Total", font=bold); put(ws, "B9", "=SUM(B4:B8)", INT); put(ws, "C9", "=SUM(C4:C8)", PCT2)
put(ws, "A10", "Check: total equals in-scope row count", font=note); put(ws, "B10", f'=IF(B9={TOTAL},"OK","MISMATCH")')
put(ws, "A12", "Controllable success rate = SUCCESS / (attempts - FAILED_BUSINESS)", font=bold); put(ws, "B12", "=B4/(B9-B6)", PCT2)
put(ws, "A13", "Ceiling success rate if ALL technical declines were removed (upper bound)", font=bold); put(ws, "B13", "=(B4+B5)/B9", PCT2)
put(ws, "A14", "Ceiling if technical declines AND pending were removed (upper bound)", font=bold); put(ws, "B14", "=(B4+B5+B8)/B9", PCT2)
put(ws, "A15", "Technical share of all non-success outcomes", font=bold); put(ws, "B15", "=B5/(B9-B4)", PCT)

put(ws, "A18", "Outcome mix by channel (share of that channel's attempts)", font=bold, fill=sect)
for col in "BCDEFG": put(ws, f"{col}18", None, fill=sect)
hdr(ws, 19, ["Channel", "Attempts"] + outs)
for i, ch in enumerate(["Mobile", "Web", "Branch_Assisted"]):
    r = 20 + i
    put(ws, f"A{r}", ch, font=bold); put(ws, f"B{r}", f"={cif(('channel', f'$A{r}'))}", INT)
    for j, o in enumerate(outs):
        col = get_column_letter(3 + j)
        put(ws, f"{col}{r}", f"={cif(('channel', f'$A{r}'), ('payment_status', chr(34)+o+chr(34)))}/$B{r}", PCT2)
widths(ws, [66, 14, 14, 18, 14, 14, 14])

# --------------------------------------------------------------- Segments
ws = wb.create_sheet("Segments", 4)
put(ws, "A1", "Segment cuts (in-scope attempts)", font=title); ws["A1"].border = Border()
put(ws, "A2", "Interval on success rate is a 95% normal-approximation interval. Segments with small n should be read with care.", font=note); ws["A2"].border = Border()
hdr(ws, 4, ["Dimension", "Value", "Attempts", "Success rate", "Success CI low", "Success CI high", "Technical decline rate",
            "Business decline rate", "Risk decline rate", "Pending >15 min rate", "Complaint rate"])
dims = [
    ("Channel", "channel", ["Mobile", "Web", "Branch_Assisted"]),
    ("Processor", "Proc_clean", ["PROC_A", "PROC_B", "PROC_C", "Unknown"]),
    ("Country", "country", ["IN", "AE", "SG"]),
    ("Payment type", "payment_type", ["Instant_Transfer", "Card", "Bank_Transfer", "Direct_Debit"]),
    ("Time band", "time_band", ["Night_00-05", "Morning_06-11", "Afternoon_12-17", "Evening_18-23"]),
    ("Device", "device_type", ["Android", "iOS", "Desktop_Browser", "Branch_Terminal"]),
    ("Customer segment", "customer_segment", ["Retail_Mass", "Retail_Affluent", "SME", "New_Customer"]),
    ("Merchant category", "merchant_category", ["Retail", "Utilities", "Travel", "Gaming", "Financial_Services", "Education", "P2P"]),
]
r = 5
SEG_ROWS = {}
for dname, col, vals in dims:
    for v in vals:
        put(ws, f"A{r}", dname); put(ws, f"B{r}", v, font=bold)
        base = (col, f"$B{r}")
        put(ws, f"C{r}", f"={cif(base)}", INT)
        put(ws, f"D{r}", f"={cif(base, ('payment_status', chr(34)+'SUCCESS'+chr(34)))}/C{r}", PCT2)
        put(ws, f"E{r}", f"=D{r}-{ZC}*SQRT(D{r}*(1-D{r})/C{r})", PCT2); put(ws, f"F{r}", f"=D{r}+{ZC}*SQRT(D{r}*(1-D{r})/C{r})", PCT2)
        put(ws, f"G{r}", f"={cif(base, ('payment_status', chr(34)+'FAILED_TECHNICAL'+chr(34)))}/C{r}", PCT2)
        put(ws, f"H{r}", f"={cif(base, ('payment_status', chr(34)+'FAILED_BUSINESS'+chr(34)))}/C{r}", PCT2)
        put(ws, f"I{r}", f"={cif(base, ('payment_status', chr(34)+'DECLINED_RISK'+chr(34)))}/C{r}", PCT2)
        put(ws, f"J{r}", f"={cif(base, ('Pending_gt15', '1'))}/C{r}", PCT2)
        put(ws, f"K{r}", f"={cif(base, ('customer_complaint_flag', '1'))}/C{r}", PCT2)
        SEG_ROWS[(dname, v)] = r
        r += 1
widths(ws, [20, 20, 12, 13, 13, 13, 17, 17, 15, 17, 14])
ws.freeze_panes = "C5"

# --------------------------------------------------------------- FailureCodes
ws = wb.create_sheet("FailureCodes", 5)
put(ws, "A1", "Failure codes: size, processor mix and device mix (non-success, in scope)", font=title); ws["A1"].border = Border()
put(ws, "A2", "'Unclassified' = non-success rows with a blank failure code. Not dropped (see DataQuality).", font=note); ws["A2"].border = Border()
codes = [("PROC_TIMEOUT", "Technical"), ("ISSUER_UNAVAILABLE", "Technical"), ("NETWORK_DROP", "Technical"), ("SYSTEM_ERROR", "Technical"),
         ("AUTH_SERVICE_FAIL", "Technical"), ("INSUFFICIENT_FUNDS", "Business"), ("LIMIT_EXCEEDED", "Business"), ("INVALID_BEneficiary".upper().replace("BENEFICIARY", "BENEFICIARY"), "Business"),
         ("CUSTOMER_AUTH_FAILED", "Business"), ("RISK_RULE_DECLINE", "Risk"), ("AWAITING_PROC_CONFIRMATION", "Pending"),
         ("SETTLEMENT_DELAY", "Pending"), ("Unclassified", "Unknown")]
hdr(ws, 4, ["Failure code", "Group", "Count", "Share of non-success", "PROC_A", "PROC_B", "PROC_C", "Unknown proc.", "Android", "iOS", "Desktop", "Branch terminal"])
FC = R("Fail_Code_clean")
for i, (code, grp) in enumerate(codes):
    r = 5 + i
    put(ws, f"A{r}", code, font=bold); put(ws, f"B{r}", grp)
    put(ws, f"C{r}", f"=COUNTIFS({FC},$A{r},{SC},1)", INT)
    put(ws, f"D{r}", f"=C{r}/$C${5+len(codes)}", PCT)
    for j, p_ in enumerate(["PROC_A", "PROC_B", "PROC_C", "Unknown"]):
        put(ws, f"{get_column_letter(5+j)}{r}", f'=COUNTIFS({FC},$A{r},{R("Proc_clean")},"{p_}",{SC},1)', INT)
    for j, d_ in enumerate(["Android", "iOS", "Desktop_Browser", "Branch_Terminal"]):
        put(ws, f"{get_column_letter(9+j)}{r}", f'=COUNTIFS({FC},$A{r},{R("device_type")},"{d_}",{SC},1)', INT)
tr = 5 + len(codes)
put(ws, f"A{tr}", "Total non-success", font=bold)
for col in "CEFGHIJKL": put(ws, f"{col}{tr}", f"=SUM({col}5:{col}{tr-1})", INT)
put(ws, f"D{tr}", f"=SUM(D5:D{tr-1})", PCT)
put(ws, f"A{tr+3}", "Unclassified share of technical failures", font=bold); put(ws, f"C{tr+3}", f'=COUNTIFS({R("payment_status")},"FAILED_TECHNICAL",{FC},"Unclassified",{SC},1)/COUNTIFS({R("payment_status")},"FAILED_TECHNICAL",{SC},1)', PCT)
put(ws, f"A{tr+4}", "Unclassified share of business failures", font=bold); put(ws, f"C{tr+4}", f'=COUNTIFS({R("payment_status")},"FAILED_BUSINESS",{FC},"Unclassified",{SC},1)/COUNTIFS({R("payment_status")},"FAILED_BUSINESS",{SC},1)', PCT)
put(ws, f"A{tr+1}", "Check vs outcome mix (total non-success)", font=note); put(ws, f"C{tr+1}", f'=IF(C{tr}=OutcomeMix!B9-OutcomeMix!B4,"OK","MISMATCH")')
widths(ws, [30, 12, 10, 16, 10, 10, 10, 12, 10, 10, 10, 14])

# --------------------------------------------------------------- HypTests
ws = wb.create_sheet("HypTests", 6)
put(ws, "A1", "Hypothesis tests (10 hypotheses from the Problem Framing Pack)", font=title); ws["A1"].border = Border()
put(ws, "A2", "Each block: hypothesis, evidence tables, a two-proportion z-test where relevant, and a formula-driven verdict. Two-sided tests at the alpha in Inputs.", font=note); ws["A2"].border = Border()
ALPHA = "Inputs!$B$6"
VERD = {}
cur = [4]

def head(hid, text):
    r = cur[0]
    put(ws, f"A{r}", hid, font=bold, fill=sect); put(ws, f"B{r}", text, font=bold, fill=sect)
    for col in "CDEFG": put(ws, f"{col}{r}", None, fill=sect)
    ws.merge_cells(f"B{r}:G{r}"); ws.row_dimensions[r].height = 32
    cur[0] += 1

def ztest(g1, x1, n1, g2, x2, n2):
    r = cur[0]
    hdr(ws, r, ["Group", "Events (x)", "Sample (n)", "Rate"])
    put(ws, f"A{r+1}", g1, font=bold); put(ws, f"B{r+1}", f"={x1}", INT); put(ws, f"C{r+1}", f"={n1}", INT); put(ws, f"D{r+1}", f"=IFERROR(B{r+1}/C{r+1},0)", PCT2)
    put(ws, f"A{r+2}", g2, font=bold); put(ws, f"B{r+2}", f"={x2}", INT); put(ws, f"C{r+2}", f"={n2}", INT); put(ws, f"D{r+2}", f"=IFERROR(B{r+2}/C{r+2},0)", PCT2)
    put(ws, f"A{r+3}", "z statistic"); put(ws, f"B{r+3}", f"=IFERROR((D{r+1}-D{r+2})/SQRT(((B{r+1}+B{r+2})/(C{r+1}+C{r+2}))*(1-(B{r+1}+B{r+2})/(C{r+1}+C{r+2}))*(1/C{r+1}+1/C{r+2})),0)", DEC)
    put(ws, f"A{r+4}", "p-value (two-sided)"); put(ws, f"B{r+4}", f"=2*(1-NORMSDIST(ABS(B{r+3})))", "0.0000")
    cur[0] = r + 5
    return dict(p1=f"D{r+1}", p2=f"D{r+2}", p=f"B{r+4}", z=f"B{r+3}", n1=f"C{r+1}", n2=f"C{r+2}")

def verdict(hid, formula):
    r = cur[0]
    put(ws, f"A{r}", "VERDICT", font=bold); put(ws, f"B{r}", formula, font=bold, fill=yellow)
    ws.merge_cells(f"B{r}:G{r}"); ws.row_dimensions[r].height = 32
    VERD[hid] = f"HypTests!B{r}"
    cur[0] = r + 2

def matrix(title_txt, row_labels, row_col, col_labels, col_col, numer_pair, fmt=PCT2):
    r = cur[0]
    put(ws, f"A{r}", title_txt, font=bold)
    hdr(ws, r + 1, [""] + col_labels + ["All"])
    for i, rl in enumerate(row_labels):
        rr = r + 2 + i
        put(ws, f"A{rr}", rl, font=bold)
        for j, cl in enumerate(col_labels):
            col = get_column_letter(2 + j)
            den = cif((row_col, f"$A{rr}"), (col_col, f"{col}${r+1}"))
            num = cif((row_col, f"$A{rr}"), (col_col, f"{col}${r+1}"), numer_pair)
            put(ws, f"{col}{rr}", f"=IFERROR({num}/{den},0)", fmt)
        allc = get_column_letter(2 + len(col_labels))
        put(ws, f"{allc}{rr}", f"=IFERROR({cif((row_col, f'$A{rr}'), numer_pair)}/{cif((row_col, f'$A{rr}'))},0)", fmt)
    cur[0] = r + 2 + len(row_labels) + 1
    return r + 1  # header row index

TECH = ("payment_status", '"FAILED_TECHNICAL"')
CLS = ("Fail_Code_clean", '"<>Unclassified"')
PROCS = ["PROC_A", "PROC_B", "PROC_C"]
BANDS = ["Night_00-05", "Morning_06-11", "Afternoon_12-17", "Evening_18-23"]
PT = ["Instant_Transfer", "Card", "Bank_Transfer", "Direct_Debit"]

# ---- A1
head("A1", "At least one processor has a materially higher technical decline rate than the others, after controlling for payment type and country.")
h1 = matrix("Technical decline rate: processor x payment type", PROCS, "Proc_clean", PT, "payment_type", TECH)
a_r, b_r, c_r = h1 + 1, h1 + 2, h1 + 3
put(ws, f"A{cur[0]-1}", "Payment types where PROC_C is highest", font=bold)
put(ws, f"B{cur[0]-1}", f"=SUMPRODUCT((B{c_r}:E{c_r}>B{a_r}:E{a_r})*(B{c_r}:E{c_r}>B{b_r}:E{b_r}))", INT)
a1_pt = f"B{cur[0]-1}"
cur[0] += 1
h1b = matrix("Technical decline rate: processor x country", PROCS, "Proc_clean", ["IN", "AE", "SG"], "country", TECH)
a2_r, c2_r = h1b + 1, h1b + 3
put(ws, f"A{cur[0]-1}", "Countries where PROC_C is highest", font=bold)
put(ws, f"B{cur[0]-1}", f"=SUMPRODUCT((B{c2_r}:D{c2_r}>B{a2_r}:D{a2_r})*(B{c2_r}:D{c2_r}>B{a2_r+1}:D{a2_r+1}))", INT)
a1_ct = f"B{cur[0]-1}"
cur[0] += 1
z = ztest("PROC_C", cif(("Proc_clean", '"PROC_C"'), TECH), cif(("Proc_clean", '"PROC_C"')),
          "PROC_A and PROC_B", f"{cif(('Proc_clean', chr(34)+'PROC_A'+chr(34)), TECH)}+{cif(('Proc_clean', chr(34)+'PROC_B'+chr(34)), TECH)}",
          f"{cif(('Proc_clean', chr(34)+'PROC_A'+chr(34)))}+{cif(('Proc_clean', chr(34)+'PROC_B'+chr(34)))}")
verdict("A1", f'=IF(AND({z["p"]}<{ALPHA},{z["p1"]}>{z["p2"]},{a1_pt}>=3,{a1_ct}>=2),"SUPPORTED in this data: PROC_C has a higher technical decline rate ("&TEXT({z["p1"]},"0.0%")&" vs "&TEXT({z["p2"]},"0.0%")&"), p="&TEXT({z["p"]},"0.0000")&", highest in "&{a1_pt}&" of 4 payment types and "&{a1_ct}&" of 3 countries",IF(AND({z["p"]}<{ALPHA},{z["p1"]}>{z["p2"]},{a1_pt}>=3),"PARTLY SUPPORTED in this data: PROC_C is higher overall ("&TEXT({z["p1"]},"0.0%")&" vs "&TEXT({z["p2"]},"0.0%")&", p="&TEXT({z["p"]},"0.0000")&") and highest in "&{a1_pt}&" of 4 payment types, but highest in only "&{a1_ct}&" of 3 countries, so the gap is not uniform (see the country table)","NOT SUPPORTED: no consistent processor gap after controlling for payment type and country"))')

# ---- A2
head("A2", "The processor gap is concentrated in specific time bands (peak load), not spread evenly.")
matrix("Technical decline rate: processor x time band", PROCS, "Proc_clean", BANDS, "time_band", TECH)
z2 = ztest("PROC_C evening (18-23)", cif(("Proc_clean", '"PROC_C"'), ("time_band", '"Evening_18-23"'), TECH), cif(("Proc_clean", '"PROC_C"'), ("time_band", '"Evening_18-23"')),
           "PROC_C other time bands", f"{cif(('Proc_clean', chr(34)+'PROC_C'+chr(34)), TECH)}-{cif(('Proc_clean', chr(34)+'PROC_C'+chr(34)), ('time_band', chr(34)+'Evening_18-23'+chr(34)), TECH)}",
           f"{cif(('Proc_clean', chr(34)+'PROC_C'+chr(34)))}-{cif(('Proc_clean', chr(34)+'PROC_C'+chr(34)), ('time_band', chr(34)+'Evening_18-23'+chr(34)))}")
verdict("A2", f'=IF(AND({z2["p"]}<{ALPHA},{z2["p1"]}>{z2["p2"]}),"SUPPORTED in this data: PROC_C evening technical decline rate is "&TEXT({z2["p1"]},"0.0%")&" vs "&TEXT({z2["p2"]},"0.0%")&" in other bands (p="&TEXT({z2["p"]},"0.0000")&")","NOT SUPPORTED: no evening concentration for PROC_C")')

# ---- B1
head("B1", "Failures on Android mobile are driven by network-type failure codes rather than processor problems.")
put(ws, f"A{cur[0]}", "Method note: shares use CLASSIFIED technical failures only. Blank failure codes (see DataQuality) hit about half of technical failures, so these samples are small.", font=note); ws[f"A{cur[0]}"].border = Border(); cur[0] += 1
z3 = ztest("Android: NETWORK_DROP among classified technical failures", cif(("device_type", '"Android"'), TECH, ("Fail_Code_clean", '"NETWORK_DROP"')), cif(("device_type", '"Android"'), TECH, CLS),
           "iOS: NETWORK_DROP among classified technical failures", cif(("device_type", '"iOS"'), TECH, ("Fail_Code_clean", '"NETWORK_DROP"')), cif(("device_type", '"iOS"'), TECH, CLS))
put(ws, f"A{cur[0]}", "Control for processor: same test within PROC_A only", font=bold); cur[0] += 1
z3b = ztest("Android, PROC_A: NETWORK_DROP share", cif(("device_type", '"Android"'), ("Proc_clean", '"PROC_A"'), TECH, ("Fail_Code_clean", '"NETWORK_DROP"')), cif(("device_type", '"Android"'), ("Proc_clean", '"PROC_A"'), TECH, CLS),
            "iOS, PROC_A: NETWORK_DROP share", cif(("device_type", '"iOS"'), ("Proc_clean", '"PROC_A"'), TECH, ("Fail_Code_clean", '"NETWORK_DROP"')), cif(("device_type", '"iOS"'), ("Proc_clean", '"PROC_A"'), TECH, CLS))
put(ws, f"A{cur[0]}", "Supporting test: technical decline rate, Android (IN) vs iOS (IN)", font=bold); cur[0] += 1
z3c = ztest("Android (IN): technical decline rate", cif(("device_type", '"Android"'), ("country", '"IN"'), TECH), cif(("device_type", '"Android"'), ("country", '"IN"')),
            "iOS (IN): technical decline rate", cif(("device_type", '"iOS"'), ("country", '"IN"'), TECH), cif(("device_type", '"iOS"'), ("country", '"IN"')))
verdict("B1", f'=IF(AND({z3["p"]}<{ALPHA},{z3["p1"]}>{z3["p2"]},{z3b["p1"]}>{z3b["p2"]}),"SUPPORTED in this data: Android failures carry more network-drop codes ("&TEXT({z3["p1"]},"0.0%")&" vs "&TEXT({z3["p2"]},"0.0%")&", p="&TEXT({z3["p"]},"0.0000")&") and the direction holds within PROC_A",IF(AND({z3c["p"]}<{ALPHA},{z3c["p1"]}>{z3c["p2"]}),"PARTLY SUPPORTED / INCONCLUSIVE in this data: Android (IN) has a higher technical decline rate ("&TEXT({z3c["p1"]},"0.0%")&" vs "&TEXT({z3c["p2"]},"0.0%")&", p="&TEXT({z3c["p"]},"0.0000")&"), but the network-drop share is not significantly different (p="&TEXT({z3["p"]},"0.00")&", small samples: "&{z3["n1"]}&" vs "&{z3["n2"]}&" classified technical failures) and does not differ within PROC_A","NOT SUPPORTED: no Android effect distinguishable from iOS once processor is controlled"))')

# ---- B2
head("B2", "A large share of failures are business or customer-side and cannot be fixed by routing or retries.")
r0 = cur[0]
hdr(ws, r0, ["Non-success outcome", "Attempts", "Share of non-success"])
for i, o in enumerate(["FAILED_TECHNICAL", "FAILED_BUSINESS", "DECLINED_RISK", "PENDING"]):
    rr = r0 + 1 + i
    put(ws, f"A{rr}", o, font=bold); put(ws, f"B{rr}", f"={cif(('payment_status', chr(34)+o+chr(34)))}", INT)
    put(ws, f"C{rr}", f"=B{rr}/SUM($B${r0+1}:$B${r0+4})", PCT)
put(ws, f"A{r0+5}", "Business share threshold (Inputs)"); put(ws, f"B{r0+5}", "=Inputs!$B$8", PCT, font=green)
cur[0] = r0 + 7
verdict("B2", f'=IF(C{r0+2}>=Inputs!$B$8,"SUPPORTED in this data: business declines are "&TEXT(C{r0+2},"0.0%")&" of non-success, above the "&TEXT(Inputs!$B$8,"0%")&" threshold, so routing cannot fix them","NOT SUPPORTED: business declines are only "&TEXT(C{r0+2},"0.0%")&" of non-success (technical "&TEXT(C{r0+1},"0.0%")&")")')

# ---- C1
head("C1", "Retries after a pending payment create more duplicate-payment risk than retries after a hard failure.")
z5 = ztest("PENDING with retry >= 1: duplicate flagged", cif(("payment_status", '"PENDING"'), ("retry_count", '">=1"'), ("duplicate_suspected_flag", "1")), cif(("payment_status", '"PENDING"'), ("retry_count", '">=1"')),
           "FAILED_TECHNICAL with retry >= 1: duplicate flagged", cif(TECH, ("retry_count", '">=1"'), ("duplicate_suspected_flag", "1")), cif(TECH, ("retry_count", '">=1"')))
verdict("C1", f'=IF(AND({z5["p"]}<{ALPHA},{z5["p1"]}>{z5["p2"]}),"SUPPORTED in this data: duplicate rate after pending retries is "&TEXT({z5["p1"]},"0.0%")&" vs "&TEXT({z5["p2"]},"0.0%")&" after technical failures (p="&TEXT({z5["p"]},"0.0000")&"). Caveat: the duplicate flag is a matching-logic output","NOT SUPPORTED: no higher duplicate rate after pending retries")')

# ---- C2
head("C2", "Long pending items (over 15 minutes) cluster in particular processor, country and payment-type combinations.")
matrix("Pending >15 min rate: processor x country", PROCS, "Proc_clean", ["IN", "AE", "SG"], "country", ("Pending_gt15", "1"))
matrix("Pending >15 min rate: processor x payment type", PROCS, "Proc_clean", PT, "payment_type", ("Pending_gt15", "1"))
pend_tot = cif(("Pending_gt15", "1"))
z6 = ztest("PROC_B outside IN", cif(("Proc_clean", '"PROC_B"'), ("country", '"<>IN"'), ("Pending_gt15", "1")), cif(("Proc_clean", '"PROC_B"'), ("country", '"<>IN"')),
           "All other combinations", f"{pend_tot}-{cif(('Proc_clean', chr(34)+'PROC_B'+chr(34)), ('country', chr(34)+'<>IN'+chr(34)), ('Pending_gt15', '1'))}",
           f"{TOTAL}-{cif(('Proc_clean', chr(34)+'PROC_B'+chr(34)), ('country', chr(34)+'<>IN'+chr(34)))}")
verdict("C2", f'=IF(AND({z6["p"]}<{ALPHA},{z6["p1"]}>{z6["p2"]}),"SUPPORTED in this data: PROC_B outside IN has a long-pending rate of "&TEXT({z6["p1"]},"0.0%")&" vs "&TEXT({z6["p2"]},"0.0%")&" elsewhere (p="&TEXT({z6["p"]},"0.0000")&")","NOT SUPPORTED: long pending is spread evenly across combinations")')

# ---- D1
head("D1", "A large share of risk declines hit genuine customers, so the precision of fraud rules is low.")
r0 = cur[0]
hdr(ws, r0, ["Measure", "Count", "Rate / share", "Interval low", "Interval high"])
put(ws, f"A{r0+1}", "Fraud alerts (in scope)", font=bold); put(ws, f"B{r0+1}", f"={cif(('fraud_alert_flag', '1'))}", INT)
put(ws, f"A{r0+2}", "Alerts confirmed as fraud", font=bold); put(ws, f"B{r0+2}", f"={cif(('fraud_alert_flag', '1'), ('confirmed_fraud_flag', '1'))}", INT)
put(ws, f"C{r0+2}", f"=B{r0+2}/B{r0+1}", PCT2)
put(ws, f"A{r0+3}", "False-positive rate (alerts not confirmed)", font=bold); put(ws, f"B{r0+3}", f"=B{r0+1}-B{r0+2}", INT)
put(ws, f"C{r0+3}", f"=B{r0+3}/B{r0+1}", PCT2)
put(ws, f"D{r0+3}", f"=C{r0+3}-{ZC}*SQRT(C{r0+3}*(1-C{r0+3})/B{r0+1})", PCT2); put(ws, f"E{r0+3}", f"=C{r0+3}+{ZC}*SQRT(C{r0+3}*(1-C{r0+3})/B{r0+1})", PCT2)
put(ws, f"A{r0+4}", "Risk declines (DECLINED_RISK)", font=bold); put(ws, f"B{r0+4}", f"={cif(('payment_status', chr(34)+'DECLINED_RISK'+chr(34)))}", INT)
put(ws, f"A{r0+5}", "Risk declines not confirmed fraud (label-biased, see note)", font=bold); put(ws, f"B{r0+5}", f"={cif(('payment_status', chr(34)+'DECLINED_RISK'+chr(34)), ('confirmed_fraud_flag', '0'))}", INT)
put(ws, f"C{r0+5}", f"=B{r0+5}/B{r0+4}", PCT2)
put(ws, f"A{r0+6}", "Complaint rate after a risk decline", font=bold); put(ws, f"B{r0+6}", f"={cif(('payment_status', chr(34)+'DECLINED_RISK'+chr(34)), ('customer_complaint_flag', '1'))}", INT)
put(ws, f"C{r0+6}", f"=B{r0+6}/B{r0+4}", PCT2)
put(ws, f"A{r0+7}", "Complaint rate on successful payments", font=bold); put(ws, f"B{r0+7}", f"={cif(('payment_status', chr(34)+'SUCCESS'+chr(34)), ('customer_complaint_flag', '1'))}", INT)
put(ws, f"C{r0+7}", f"=B{r0+7}/{cif(('payment_status', chr(34)+'SUCCESS'+chr(34)))}", PCT2)
put(ws, f"A{r0+8}", "Note: declined payments rarely get a confirmed-fraud label, so the decline false-positive rate is overstated. The alert-based rate is the better evidence. Reviewed-outcome labels are needed (Open item O-5).", font=note); ws[f"A{r0+8}"].border = Border()
cur[0] = r0 + 10
verdict("D1", f'=IF(C{r0+3}>=Inputs!$B$9,"SUPPORTED in this data: "&TEXT(C{r0+3},"0.0%")&" of fraud alerts were not confirmed as fraud (above the "&TEXT(Inputs!$B$9,"0%")&" threshold); complaint rate after a risk decline is "&TEXT(C{r0+6},"0.0%")&" vs "&TEXT(C{r0+7},"0.0%")&" on successes","NOT SUPPORTED: most alerts are confirmed fraud")')

# ---- D2
head("D2", "Tuning rules for low-risk segments can recover completions while fraud loss rate stays within a control limit.")
r0 = cur[0]
hdr(ws, r0, ["Segment", "Alerts", "Alerts confirmed", "Alert precision", "Risk declines", "Loss USD", "Approved value USD", "Loss rate", "Candidate for tuning?", "Added loss low USD", "Added loss high USD", "Added loss high as % of current loss"])
segs = [("customer_segment", v) for v in ["Retail_Mass", "Retail_Affluent", "SME", "New_Customer"]] + [("merchant_category", v) for v in ["Retail", "Utilities", "Travel", "Gaming", "Financial_Services", "Education", "P2P"]]
DR = ("payment_status", '"DECLINED_RISK"')
OA_ROW = r0 + 1 + len(segs)  # overall row
for i, (col, v) in enumerate(segs):
    rr = r0 + 1 + i
    put(ws, f"A{rr}", v, font=bold)
    base = (col, f"$A{rr}")
    put(ws, f"B{rr}", f"={cif(base, ('fraud_alert_flag', '1'))}", INT)
    put(ws, f"C{rr}", f"={cif(base, ('fraud_alert_flag', '1'), ('confirmed_fraud_flag', '1'))}", INT)
    put(ws, f"D{rr}", f"=IFERROR(C{rr}/B{rr},0)", PCT2)
    put(ws, f"E{rr}", f"={cif(base, DR)}", INT)
    put(ws, f"F{rr}", f"={sif('fraud_loss_usd', base)}", "#,##0")
    put(ws, f"G{rr}", f"={sif('amount_usd_equiv', base, ('payment_status', chr(34)+'SUCCESS'+chr(34)))}", "#,##0")
    put(ws, f"H{rr}", f"=IFERROR(F{rr}/G{rr},0)", "0.000%")
    put(ws, f"I{rr}", f'=IF(AND(H{rr}<=$H${OA_ROW},D{rr}<=$D${OA_ROW},B{rr}>=30),"Yes","No")')
    avg = f"IFERROR(AVERAGEIFS({R('amount_usd_equiv')},{R(col)},$A{rr},{R('payment_status')},\"DECLINED_RISK\",{SC},1),0)"
    se = f"SQRT(D{rr}*(1-D{rr})/MAX(B{rr},1))"
    put(ws, f"J{rr}", f'=IF(I{rr}="Yes",E{rr}*MAX(0,D{rr}-{ZC}*{se})*{avg},"")', "#,##0")
    put(ws, f"K{rr}", f'=IF(I{rr}="Yes",E{rr}*MIN(1,D{rr}+{ZC}*{se})*{avg},"")', "#,##0")
    put(ws, f"L{rr}", f'=IF(I{rr}="Yes",IFERROR(K{rr}/F{rr},0),"")', PCT)
rr = OA_ROW
put(ws, f"A{rr}", "Overall", font=bold)
put(ws, f"B{rr}", f"={cif(('fraud_alert_flag', '1'))}", INT); put(ws, f"C{rr}", f"={cif(('fraud_alert_flag', '1'), ('confirmed_fraud_flag', '1'))}", INT)
put(ws, f"D{rr}", f"=C{rr}/B{rr}", PCT2); put(ws, f"E{rr}", f"={cif(DR)}", INT)
put(ws, f"F{rr}", f"={sif('fraud_loss_usd')}", "#,##0"); put(ws, f"G{rr}", f"={sif('amount_usd_equiv', ('payment_status', chr(34)+'SUCCESS'+chr(34)))}", "#,##0")
put(ws, f"H{rr}", f"=F{rr}/G{rr}", "0.000%")
cur[0] = rr + 1
put(ws, f"A{cur[0]}", "Candidate = loss rate and alert precision at or below overall, with at least 30 alerts. Added loss assumes declined items carry the segment's alert precision (ASSUMPTION; intervals are normal-approximation). Indicative screen, not a forecast; test any change on matured data (requirement FR6).", font=note)
ws[f"A{cur[0]}"].border = Border(); cur[0] += 2
cand_rng = f"I{r0+1}:I{r0+len(segs)}"
first = lambda col: f'INDEX({col}{r0+1}:{col}{r0+len(segs)},MATCH("Yes",{cand_rng},0))'
verdict("D2", f'=IF(COUNTIF({cand_rng},"Yes")=0,"NOT SUPPORTED: no segment shows room to relax rules without exceeding overall loss and precision levels",IF({first("L")}>1,"PARTLY SUPPORTED / THIN EVIDENCE: "&COUNTIF({cand_rng},"Yes")&" segment screens as a tuning candidate ("&{first("A")}&": loss rate "&TEXT({first("H")},"0.000%")&", "&{first("C")}&" confirmed of "&{first("B")}&" alerts), but the high-end added loss is "&TEXT({first("L")},"0%")&" of its current loss. Run a controlled test on matured data before any change","SUPPORTED in this data: "&COUNTIF({cand_rng},"Yes")&" segment(s) screen as tuning candidates and the added-loss range stays below current loss"))')

# ---- E1
head("E1", "Manual reconciliation effort is driven mainly by pending and duplicate items, so fixing branch C reduces effort more than automating reconciliation alone.")
r0 = cur[0]
hdr(ws, r0, ["Primary cause (single cause per case)", "Cases", "Share of cases", "Handling cost USD", "Share of cost"])
causes = ["Pending", "Duplicate", "Technical failure", "Other"]
RC = R("Recon_cause")
for i, cz in enumerate(causes):
    rr = r0 + 1 + i
    put(ws, f"A{rr}", cz, font=bold); put(ws, f"B{rr}", f'=COUNTIFS({RC},$A{rr},{SC},1)', INT)
    put(ws, f"C{rr}", f"=B{rr}/$B${r0+5}", PCT)
    put(ws, f"D{rr}", f'=SUMIFS({R("handling_cost_usd")},{RC},$A{rr},{SC},1)', "#,##0")
    put(ws, f"E{rr}", f"=D{rr}/$D${r0+5}", PCT)
put(ws, f"A{r0+5}", "Total", font=bold); put(ws, f"B{r0+5}", f"=SUM(B{r0+1}:B{r0+4})", INT); put(ws, f"C{r0+5}", f"=SUM(C{r0+1}:C{r0+4})", PCT)
put(ws, f"D{r0+5}", f"=SUM(D{r0+1}:D{r0+4})", "#,##0"); put(ws, f"E{r0+5}", f"=SUM(E{r0+1}:E{r0+4})", PCT)
put(ws, f"A{r0+6}", "Check vs reconciliation_flag count", font=note); put(ws, f"B{r0+6}", f'=IF(B{r0+5}={cif(("reconciliation_flag","1"))},"OK","MISMATCH")')
put(ws, f"A{r0+7}", "Cause hierarchy: pending, then duplicate, then technical failure, then other. Cost is the full handling cost of those cases.", font=note); ws[f"A{r0+7}"].border = Border()
cur[0] = r0 + 9
verdict("E1", f'=IF(C{r0+1}+C{r0+2}>0.5,"SUPPORTED in this data: pending and duplicate items cause "&TEXT(C{r0+1}+C{r0+2},"0.0%")&" of reconciliation cases ("&TEXT(E{r0+1}+E{r0+2},"0.0%")&" of cost); technical failures add "&TEXT(C{r0+3},"0.0%")&"","NOT SUPPORTED: other causes dominate reconciliation ("&TEXT(1-C{r0+1}-C{r0+2},"0.0%")&" not pending or duplicate)")')

# ---- E2
head("E2", "Investigation SLA breaches concentrate in branch-assisted cases and certain case types.")
r0 = cur[0]
hdr(ws, r0, ["Channel", "Valid cases", "Median hours", "90th percentile hours", "Within SLA", "SLA attainment", "Breaches"])
chs = [("Mobile", "Inv_h_Mobile"), ("Web", "Inv_h_Web"), ("Branch_Assisted", "Inv_h_Branch")]
for i, (ch, hc) in enumerate(chs):
    rr = r0 + 1 + i
    put(ws, f"A{rr}", ch, font=bold)
    put(ws, f"B{rr}", f"=COUNT({R(hc)})", INT)
    put(ws, f"C{rr}", f"=MEDIAN({R(hc)})", "0.0"); put(ws, f"D{rr}", f"=PERCENTILE({R(hc)},0.9)", "0.0")
    put(ws, f"E{rr}", f'=COUNTIFS({R("channel")},$A{rr},{R("In_SLA")},1,{SC},1)', INT)
    put(ws, f"F{rr}", f"=E{rr}/B{rr}", PCT); put(ws, f"G{rr}", f"=B{rr}-E{rr}", INT)
rr = r0 + 4
put(ws, f"A{rr}", "All channels", font=bold); put(ws, f"B{rr}", f"=SUM(B{r0+1}:B{r0+3})", INT)
put(ws, f"C{rr}", f"=MEDIAN({R('Inv_hours_valid')})", "0.0"); put(ws, f"D{rr}", f"=PERCENTILE({R('Inv_hours_valid')},0.9)", "0.0")
put(ws, f"E{rr}", f"=SUM(E{r0+1}:E{r0+3})", INT); put(ws, f"F{rr}", f"=E{rr}/B{rr}", PCT); put(ws, f"G{rr}", f"=SUM(G{r0+1}:G{r0+3})", INT)
cur[0] = r0 + 6
z7 = ztest("Branch_Assisted breach rate", f"G{r0+3}", f"B{r0+3}", "Mobile + Web breach rate", f"G{r0+1}+G{r0+2}", f"B{r0+1}+B{r0+2}")
verdict("E2", f'=IF(AND({z7["p"]}<{ALPHA},{z7["p1"]}>{z7["p2"]}),"SUPPORTED in this data: Branch_Assisted SLA breach rate is "&TEXT({z7["p1"]},"0.0%")&" vs "&TEXT({z7["p2"]},"0.0%")&" for mobile and web (p="&TEXT({z7["p"]},"0.0000")&"). Case-type split not tested (needs one primary trigger per case)","NOT SUPPORTED: SLA breach rate is not higher for branch-assisted cases")')
widths(ws, [52, 16, 16, 16, 16, 16, 16, 14, 18, 16, 16, 20])

# --------------------------------------------------------------- DataQuality
ws = wb.create_sheet("DataQuality", 7)
put(ws, "A1", "Data-quality findings and their effect", font=title); ws["A1"].border = Border()
put(ws, "A2", "Counts are detected by formula. The 'seeded' column comes from the generation rules and is shown only to check the method catches the issues.", font=note); ws["A2"].border = Border()
hdr(ws, 4, ["Issue", "Detected count", "Share of all rows", "Seeded count (generation rules)", "Treatment in this analysis", "Possible bias", "Share of the affected population", "Affected population"])
ALLR = f"COUNTA({R('Transaction_ID')})"
dq = [
    ("Repeated Transaction_ID rows", f"=SUM({R('Dup_ID_repeat')})", 36, "Later rows with the same ID excluded (In_Scope = 0).", "Which of two rows is the true record is unknown; excluding the later one may drop a valid payment."),
    ("Failed rows with blank failure_code", f'=COUNTIFS({R("payment_status")},"FAILED_TECHNICAL",{R("failure_code")},"")+COUNTIFS({R("payment_status")},"FAILED_BUSINESS",{R("failure_code")},"")', 240, "Reported as 'Unclassified' in FailureCodes, not dropped.", "Hides the cause for a large share of technical and business failures (see the affected-population column); any code-mix analysis rests on the classified half only."),
    ("Blank processor", f"=COUNTBLANK({R('processor')})", 96, "Shown as 'Unknown' in every processor cut.", "Processor comparisons exclude these rows; small effect if random."),
    ("Latency above cap (outliers)", f"=SUM({R('Latency_outlier')})", 14, "Flagged; latency is not used in KPI rates. Treat before any latency average.", "Would distort mean latency; median is safer."),
    ("Investigations with end before start", f"=COUNTIFS({R('Inv_valid')},0)", 25, "Excluded from SLA and median duration.", "Biases duration if invalid rows are mostly long or short cases."),
    ("Branch-assisted rows with no latency", f'=COUNTIFS({R("channel")},"Branch_Assisted",{R("latency_ms")},"")', 469, "Latency comparisons across channels are not made.", "Systematic missingness: channel latency comparison is biased."),
]
POP = [
    (ALLR, "All rows"),
    (f'COUNTIF({R("payment_status")},"FAILED_TECHNICAL")+COUNTIF({R("payment_status")},"FAILED_BUSINESS")', "Technical + business failed rows"),
    (ALLR, "All rows"),
    (f'COUNT({R("latency_ms")})', "Rows with latency recorded"),
    (f'COUNT({R("Inv_hours_raw")})', "Rows with an investigation"),
    (f'COUNTIF({R("channel")},"Branch_Assisted")', "Branch-assisted rows"),
]
for i, (name, f_, seeded, treat, bias) in enumerate(dq):
    r = 5 + i
    put(ws, f"A{r}", name, font=bold); put(ws, f"B{r}", f_, INT); put(ws, f"C{r}", f"=B{r}/{ALLR}", PCT2)
    put(ws, f"D{r}", seeded, INT, font=blue); put(ws, f"E{r}", treat); put(ws, f"F{r}", bias)
    pop_f, pop_name = POP[i]
    put(ws, f"G{r}", f"=B{r}/({pop_f})", PCT); put(ws, f"H{r}", pop_name)
put(ws, "A12", "Sensitivity of the headline success rate to data treatment", font=bold, fill=sect)
for col in "BCD": put(ws, f"{col}12", None, fill=sect)
hdr(ws, 13, ["Treatment", "Success rate", "Rows used", "Difference vs in-scope (pp)"])
put(ws, "A14", "In scope (duplicate-ID rows excluded) - used in this analysis", font=bold); put(ws, "B14", f"=CaseFacts!C5", PCT2); put(ws, "C14", f"={TOTAL}", INT); put(ws, "D14", "=0", DEC)
put(ws, "A15", "All rows (duplicate-ID rows included)", font=bold); put(ws, "B15", f'=COUNTIF({R("payment_status")},"SUCCESS")/{ALLR}', PCT2); put(ws, "C15", f"={ALLR}", INT)
put(ws, "D15", "=(B15-B14)*100", "0.000")
put(ws, "A16", "In scope, excluding blank-processor rows", font=bold)
put(ws, "B16", f'=COUNTIFS({SC},1,{R("payment_status")},"SUCCESS",{R("Proc_clean")},"<>Unknown")/COUNTIFS({SC},1,{R("Proc_clean")},"<>Unknown")', PCT2)
put(ws, "C16", f'=COUNTIFS({SC},1,{R("Proc_clean")},"<>Unknown")', INT); put(ws, "D16", "=(B16-B14)*100", "0.000")
put(ws, "A18", "Other limits of the data (not detectable by formula)", font=bold)
lims = ["One row per payment, not per event: pending duration and state history are summarised, so late-arriving events cannot be tested (see Workstream 4).",
        "No device or network telemetry: Android 'network' effects rely on failure codes only.",
        "No staff time-tracking: reconciliation hours cannot be measured (assumption A-004).",
        "Fraud labels are instantaneous with no chargeback lag; no reviewed-outcome label for risk declines (open items O-2, O-5).",
        "Single quarter: no seasonality, and monthly control limits cannot be set (needs at least 12 points).",
        "Synthetic data: relationships were built in, so findings show the method working, not facts about a bank."]
for i, t in enumerate(lims):
    c = ws.cell(row=19 + i, column=1, value=f"- {t}"); c.font = body
widths(ws, [52, 16, 14, 22, 60, 60, 18, 30])

# --------------------------------------------------------------- Opportunity
ws = wb.create_sheet("Opportunity", 8)
put(ws, "A1", "Opportunity sizing (indicative upper bounds, only where the data supports it)", font=title); ws["A1"].border = Border()
put(ws, "A2", "Expressed as percentage points of attempts or USD per 1,000 attempts so it scales to any real volume. Not additive: opportunities overlap. Upper bounds before retry recovery, customer abandonment and implementation limits.", font=note); ws["A2"].border = Border()
hdr(ws, 4, ["Opportunity", "Group at issue: n", "Rate at issue", "Benchmark rate", "Gap", "Gap interval low", "Gap interval high", "Success gain low (pp of attempts)", "Success gain high (pp of attempts)", "Basis"])
def opp(r, name, n_f, x_f, nb_f, xb_f, basis):
    put(ws, f"A{r}", name, font=bold)
    put(ws, f"B{r}", f"={n_f}", INT); put(ws, f"C{r}", f"={x_f}/B{r}", PCT2)
    put(ws, f"D{r}", f"={xb_f}/({nb_f})", PCT2)
    put(ws, f"E{r}", f"=C{r}-D{r}", PCT2)
    se = f"SQRT(C{r}*(1-C{r})/B{r}+D{r}*(1-D{r})/({nb_f}))"
    put(ws, f"F{r}", f"=MAX(0,E{r}-{ZC}*{se})", PCT2); put(ws, f"G{r}", f"=MAX(0,E{r}+{ZC}*{se})", PCT2)
    put(ws, f"H{r}", f"=B{r}*F{r}/{TOTAL}*100", "0.00"); put(ws, f"I{r}", f"=B{r}*G{r}/{TOTAL}*100", "0.00")
    put(ws, f"J{r}", basis)
EV = ("time_band", '"Evening_18-23"')
opp(5, "O1: bring PROC_C evening technical declines to PROC_A evening level",
    cif(("Proc_clean", '"PROC_C"'), EV), cif(("Proc_clean", '"PROC_C"'), EV, TECH),
    cif(("Proc_clean", '"PROC_A"'), EV), cif(("Proc_clean", '"PROC_A"'), EV, TECH),
    "Routing / capacity fix for PROC_C at peak. Hypotheses A1, A2.")
opp(6, "O2: bring Android (IN) technical declines to iOS (IN) level",
    cif(("device_type", '"Android"'), ("country", '"IN"')), cif(("device_type", '"Android"'), ("country", '"IN"'), TECH),
    cif(("device_type", '"iOS"'), ("country", '"IN"')), cif(("device_type", '"iOS"'), ("country", '"IN"'), TECH),
    "Network resilience and retry design on mobile. Hypothesis B1.")
put(ws, "A7", "O1 and O2 overlap (Android users also route through PROC_C): do not add them.", font=note); ws["A7"].border = Border()

hdr(ws, 9, ["Cost opportunity", "Value", "Unit", "Basis"])
put(ws, "A10", "Reconciliation cases per 1,000 attempts", font=bold); put(ws, "B10", f"={cif(('reconciliation_flag','1'))}/{TOTAL}*1000", "0.0"); put(ws, "C10", "cases per 1,000 attempts"); put(ws, "D10", "Measured")
put(ws, "A11", "Average handling cost of a reconciled case (USD)", font=bold); put(ws, "B11", f"={sif('handling_cost_usd', ('reconciliation_flag','1'))}/{cif(('reconciliation_flag','1'))}", "0.00"); put(ws, "C11", "USD per case"); put(ws, "D11", "Measured; unit costs are assumptions (A-007)")
put(ws, "A12", "Cases avoided by automation, low (assumption)", font=bold); put(ws, "B12", "=Inputs!$B$21", PCT, font=green); put(ws, "D12", "ASSUMPTION for sensitivity only")
put(ws, "A13", "Cases avoided by automation, high (assumption)", font=bold); put(ws, "B13", "=Inputs!$B$22", PCT, font=green); put(ws, "D13", "ASSUMPTION for sensitivity only")
put(ws, "A14", "Saving per 1,000 attempts, low (USD)", font=bold); put(ws, "B14", "=B10*B11*B12", "0.00"); put(ws, "C14", "USD per 1,000 attempts")
put(ws, "A15", "Saving per 1,000 attempts, high (USD)", font=bold); put(ws, "B15", "=B10*B11*B13", "0.00"); put(ws, "C15", "USD per 1,000 attempts")
put(ws, "A16", "Reconciliation cases caused by pending or duplicate items", font=bold); put(ws, "B16", f"=(COUNTIFS({RC},\"Pending\",{SC},1)+COUNTIFS({RC},\"Duplicate\",{SC},1))/{cif(('reconciliation_flag','1'))}", PCT); put(ws, "C16", "share of cases"); put(ws, "D16", "Addressable by retry and pending fixes (hypotheses C1, E1)")
widths(ws, [66, 16, 14, 16, 12, 14, 14, 18, 18, 56])

# --------------------------------------------------------------- Findings
ws = wb.create_sheet("Findings", 9)
put(ws, "A1", "Findings summary (formula-driven, synthetic data)", font=title); ws["A1"].border = Border()
put(ws, "A2", "Every sentence below recalculates from the analysis sheets. Rewrite in your own words for the case pack.", font=note); ws["A2"].border = Border()
hdr(ws, 4, ["#", "Finding"])
fnds = [
    ('="Success rate: measured "&TEXT(CaseFacts!C5,"0.0%")&" vs case "&TEXT(CaseFacts!B5,"0.0%")&". "&CaseFacts!J5'),
    ('="Pending over 15 minutes: measured "&TEXT(CaseFacts!C6,"0.00%")&" vs case "&TEXT(CaseFacts!B6,"0.0%")&". "&CaseFacts!J6'),
    ('="Median investigation: "&TEXT(CaseFacts!C7,"0.0")&" hours vs case "&TEXT(CaseFacts!B7,"0.0")&". "&CaseFacts!J7'),
    ('="Fraud loss rate: "&TEXT(CaseFacts!C8,"0.000%")&" of all-attempt value ("&TEXT(CaseFacts!C9,"0.000%")&" of approved value) vs case "&TEXT(CaseFacts!B8,"0.00%")&". "&CaseFacts!J8'),
    ('="Reconciliation hours: "&CaseFacts!J10'),
    ('="Failure pool: technical declines are "&TEXT(OutcomeMix!B15,"0.0%")&" of non-success outcomes. Removing all of them would lift success to at most "&TEXT(OutcomeMix!B13,"0.0%")&"."'),
    ('="Processor (A1): "&' + VERD["A1"]),
    ('="Time band (A2): "&' + VERD["A2"]),
    ('="Android (B1): "&' + VERD["B1"]),
    ('="Customer-side failures (B2): "&' + VERD["B2"]),
    ('="Retries and duplicates (C1): "&' + VERD["C1"]),
    ('="Pending clusters (C2): "&' + VERD["C2"]),
    ('="Fraud rule precision (D1): "&' + VERD["D1"]),
    ('="Rule tuning room (D2): "&' + VERD["D2"]),
    ('="Reconciliation drivers (E1): "&' + VERD["E1"]),
    ('="Investigation SLA (E2): "&' + VERD["E2"]),
    ('="Indicative upper-bound success gain: O1 up to "&TEXT(Opportunity!I5,"0.00")&" pp, O2 up to "&TEXT(Opportunity!I6,"0.00")&" pp of attempts (overlapping, not additive)."'),
]
for i, f_ in enumerate(fnds):
    r = 5 + i
    put(ws, f"A{r}", i + 1, font=bold); put(ws, f"B{r}", f_)
    ws.row_dimensions[r].height = 48
r = 5 + len(fnds) + 1
put(ws, f"A{r}", "", font=bold); put(ws, f"B{r}", "What would change the recommendation", font=bold, fill=sect)
chg = ["Real data showing the processor gap disappears once device, payment type and country are controlled (hypothesis A1 refuted).",
       "A fraud loss ceiling from the CRO tighter than the added-loss range for the candidate segments (D2).",
       "Evidence that retries after pending do not create duplicates when idempotency keys are used (C1).",
       "Real handling-time data showing reconciliation hours are driven by volume the interventions do not touch (E1).",
       "Reviewed-outcome labels showing risk declines are mostly genuine fraud (D1 reversed)."]
for i, t in enumerate(chg):
    put(ws, f"A{r+1+i}", i + 1); put(ws, f"B{r+1+i}", t)
widths(ws, [6, 160])

wb.save(OUT)
print("saved", OUT)
