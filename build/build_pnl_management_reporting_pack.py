"""Build the Management Reporting Pack workbook (v3, shaped to a standard P&L by Month export).

Architecture
------------
Two data sheets at the back:

- `Data`                  : shaped like a standard **Profit and Loss by Month**
                            export. Account code + Account name + Account type
                            + Report line + 12 monthly columns. User pastes
                            their accounting-software export in here, then maps each account
                            to one of five Report lines via the dropdown.
- `Internal Data Measures`: non-software inputs (working-capital days, department
                            wage shares, margin-bridge driver allocation).

Every analytical sheet aggregates from `Data` using SUMIFS keyed on the
Report line column. Charts and sparklines use Lyros brand colours.

Output:
  C:\\dev\\lyros-workbooks\\library\\pnl
        \\lyros_lib_pnl_management_reporting_pack.xlsx
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import xlsxwriter

sys.path.insert(0, str(Path(__file__).parent))
from _shared import branding as b
from _shared import data as d
from _shared import scaffold as sc
from _shared import styles as st


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\4000 - Management Reporting Pack.xlsx")


WORKBOOK_ID = "4000"
WORKBOOK_TITLE = "Management Reporting Pack"
WORKBOOK_KICKER = "Monthly board read-out for SMEs"
TARGET_USER = (
    "Owner-operated SME bookkeeper or finance lead preparing a monthly "
    "board read-out in under 90 minutes."
)
HOW_TO_USE = [
    "Open the Data sheet and paste your Profit and Loss by Month export. Map each account to a Report line.",
    "Open Internal Data Measures and update the working-capital days and department wage shares with your own figures.",
    "Every analytical sheet (Headline, P&L Monthly, Quarter Comparison, Wages, Working Capital) recalculates automatically.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "Wholesale and retail (Wholesale 50%, Retail 35%, Online 12%, Other 3%)"),
    ("REVENUE SCALE", "Circa $4M annual, growing modestly month-on-month"),
    ("GROSS MARGIN", "44 per cent on average"),
    ("HEADCOUNT MIX", "Operations 45%, Sales 22%, Finance and admin 14%, Customer service 11%, Executive 8%"),
    ("WORKING CAPITAL", "DSO 38 days, DPO 46 days, DIO 28 days"),
]

INPUTS_REQUIRED = [
    ("Profit and loss by month", "Data tab (paste P&L by Month export)"),
    ("Account to Report line mapping", "Data tab (dropdown in Report line column)"),
    ("Working capital days (DSO, DPO, DIO)", "Internal Data Measures tab → drives Working Capital tab"),
    ("Department wage allocation", "Internal Data Measures tab → drives Wages tab"),
]


# ── Layout constants ────────────────────────────────────────────────────────

NUM_MONTHS = 12

# Data sheet columns (0-based)
DATA_COL_CODE = 1   # B
DATA_COL_NAME = 2   # C
DATA_COL_TYPE = 3   # D
DATA_COL_LINE = 4   # E
DATA_COL_M1 = 5     # F (Jan)
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1  # Q (12th month)
DATA_COL_TOTAL = DATA_COL_M_LAST + 1            # R

# Data sheet rows (1-based; xlsxwriter uses 0-based so subtract 1 at write time)
DATA_ROW_HEADER = 9     # account-level table header (1-based row 9)
DATA_ROW_FIRST_ACCT = 10
DATA_ROW_LAST_ACCT = DATA_ROW_FIRST_ACCT + 16  # 17 accounts → row 26
DATA_ROW_COL_TOTAL = DATA_ROW_LAST_ACCT + 1     # row 27 column totals
DATA_ROW_RL_HEADER = DATA_ROW_COL_TOTAL + 4     # report-line summary header
DATA_ROW_RL_FIRST = DATA_ROW_RL_HEADER + 1

REPORT_LINES = ["Revenue", "Cost of sales", "Wages", "Other opex", "D&A"]

# Synthetic account list (standard three-digit account codes).
ACCOUNTS = [
    # code, name,                            type,           report_line,    share key
    ("200", "Sales - Wholesale",             "Revenue",      "Revenue",       ("revenue", 0.50)),
    ("210", "Sales - Retail",                "Revenue",      "Revenue",       ("revenue", 0.35)),
    ("220", "Sales - Online",                "Revenue",      "Revenue",       ("revenue", 0.12)),
    ("290", "Other Income",                  "Other Income", "Revenue",       ("revenue", 0.03)),
    ("310", "Cost of Goods Sold",            "Direct Costs", "Cost of sales", ("cogs",    0.55)),
    ("320", "Purchases - Materials",         "Direct Costs", "Cost of sales", ("cogs",    0.45)),
    ("477", "Wages and Salaries",            "Expense",      "Wages",         ("wages",   0.73)),
    ("478", "Superannuation",                "Expense",      "Wages",         ("wages",   0.11)),
    ("479", "Workers Compensation",          "Expense",      "Wages",         ("wages",   0.04)),
    ("480", "Annual Leave Provision",        "Expense",      "Wages",         ("wages",   0.12)),
    ("469", "Rent",                          "Overheads",    "Other opex",    ("opex",    0.30)),
    ("451", "Light Power Heating",           "Overheads",    "Other opex",    ("opex",    0.08)),
    ("433", "Insurance",                     "Overheads",    "Other opex",    ("opex",    0.12)),
    ("461", "Marketing and Advertising",     "Overheads",    "Other opex",    ("opex",    0.20)),
    ("463", "Office Expenses",               "Overheads",    "Other opex",    ("opex",    0.08)),
    ("466", "Accounting and Legal Fees",     "Overheads",    "Other opex",    ("opex",    0.22)),
    ("416", "Depreciation",                  "Depreciation", "D&A",           ("da",      1.00)),
]

# Internal Data Measures rows
IDM_ROW_WC_HEADER = 8
IDM_ROW_DSO = IDM_ROW_WC_HEADER + 1
IDM_ROW_DPO = IDM_ROW_WC_HEADER + 2
IDM_ROW_DIO = IDM_ROW_WC_HEADER + 3
IDM_ROW_CCC = IDM_ROW_WC_HEADER + 4

IDM_ROW_DEPT_HEADER = 16
IDM_ROW_DEPT_FIRST = 17

IDM_ROW_BRIDGE_HEADER = 24
IDM_ROW_BRIDGE_FIRST = 25

IDM_DEPTS = ["Operations", "Sales", "Finance and admin", "Customer service", "Executive"]
IDM_BRIDGE_DRIVERS = ["Volume", "Price", "Mix", "Cost"]


def col_letter(zero_based: int) -> str:
    s = ""
    n = zero_based
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


# ── SUMIFS reference helpers (read from Data) ──────────────────────────────

def sumifs_month(report_line: str, month_idx_0: int) -> str:
    """Sum a Report line for the given month column (0..NUM_MONTHS-1)."""
    sum_col = col_letter(DATA_COL_M1 + month_idx_0)
    crit_col = col_letter(DATA_COL_LINE)
    return (
        f"=SUMIFS('Data'!${sum_col}${DATA_ROW_FIRST_ACCT}:${sum_col}${DATA_ROW_LAST_ACCT},"
        f"'Data'!${crit_col}${DATA_ROW_FIRST_ACCT}:${crit_col}${DATA_ROW_LAST_ACCT},"
        f"\"{report_line}\")"
    )


def sumifs_fy(report_line: str) -> str:
    """Sum a Report line across the FY total column on Data."""
    sum_col = col_letter(DATA_COL_TOTAL)
    crit_col = col_letter(DATA_COL_LINE)
    return (
        f"=SUMIFS('Data'!${sum_col}${DATA_ROW_FIRST_ACCT}:${sum_col}${DATA_ROW_LAST_ACCT},"
        f"'Data'!${crit_col}${DATA_ROW_FIRST_ACCT}:${crit_col}${DATA_ROW_LAST_ACCT},"
        f"\"{report_line}\")"
    )


# ── Synthetic data generation ───────────────────────────────────────────────

def _synthetic_data() -> dict:
    rng_rev = d.make_rng("mrp3-rev")
    rng_cogs = d.make_rng("mrp3-cogs")
    rng_op = d.make_rng("mrp3-op")
    rng_wg = d.make_rng("mrp3-wg")

    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    revenue = d.revenue_series(NUM_MONTHS, rng_rev, base=320_000)
    cogs = d.cogs_series(revenue, rng_cogs, gm_pct=0.44)
    opex = d.opex_series(NUM_MONTHS, rng_op, base=82_000)
    wages = d.wages_series(revenue, rng_wg, ratio=0.21)
    da = [round(r * 0.04, 0) for r in revenue]

    rng_d1 = d.make_rng("mrp3-dso")
    rng_d2 = d.make_rng("mrp3-dpo")
    rng_d3 = d.make_rng("mrp3-dio")
    dso = [round(38 + rng_d1.uniform(-4, 6), 0) for _ in months]
    dpo = [round(46 + rng_d2.uniform(-5, 5), 0) for _ in months]
    dio = [round(28 + rng_d3.uniform(-3, 5), 0) for _ in months]

    totals = {
        "revenue": revenue, "cogs": cogs, "opex": opex,
        "wages": wages, "da": da,
    }

    # Allocate each total to its constituent accounts using the ratios in
    # ACCOUNTS. Result: a list of 17 (account_row) × 12 (month) floats.
    account_monthly: list[list[float]] = []
    for code, name, _type, _rl, (series_key, share) in ACCOUNTS:
        series = totals[series_key]
        rng_acct = d.make_rng(f"mrp3-{code}")
        row = [
            round(series[m] * (share + rng_acct.uniform(-0.015, 0.015)), 0)
            for m in range(NUM_MONTHS)
        ]
        account_monthly.append(row)

    return {
        "months": months,
        "account_monthly": account_monthly,
        "dso": dso, "dpo": dpo, "dio": dio,
        "dept_share": [0.45, 0.22, 0.14, 0.11, 0.08],
        "bridge_alloc": [0.45, 0.20, 0.15, 0.20],
    }


# ── Data sheet (P&L by Month export shape) ─────────────────────────────

def build_data_sheet(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_CODE, DATA_COL_CODE, 8)
    ws.set_column(DATA_COL_NAME, DATA_COL_NAME, 28)
    ws.set_column(DATA_COL_TYPE, DATA_COL_TYPE, 14)
    ws.set_column(DATA_COL_LINE, DATA_COL_LINE, 14)
    ws.set_column(DATA_COL_M1, DATA_COL_M_LAST, 11)
    ws.set_column(DATA_COL_TOTAL, DATA_COL_TOTAL, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Single source of truth",
        title="Drop your data here",
        last_col=LAST_COL + 1,
        explanation=(
            "Paste your Profit and Loss by Month export into the table below. "
            "Each row is one account from your chart of accounts. The Report line column "
            "controls how the rest of the workbook aggregates the numbers; keep it set to "
            "one of: Revenue, Cost of sales, Wages, Other opex, or D&A. Other sheets read "
            "from this table using SUMIFS by Report line."
        ),
    )

    # Section header for the account-level table
    st.write_section_header(
        ws, formats, row=6,
        kicker="Step 1   Paste your P&L by Month export",
        title="Profit and loss accounts",
    )

    # Table headers
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    headers = ["Code", "Account name", "Account type", "Report line"] + month_headers + ["FY total"]
    st.write_header_row(
        ws, formats, row=DATA_ROW_HEADER - 1, headers=headers,
        start_col=DATA_COL_CODE, right_align_from=5,
    )

    # Account rows
    for i, (code, name, acct_type, report_line, (_, _)) in enumerate(ACCOUNTS):
        row_0 = DATA_ROW_FIRST_ACCT - 1 + i
        zebra = i % 2 == 0
        label_fmt = formats["input_label" if not zebra else "input_label"]
        text_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(row_0, 20)
        ws.write_string(row_0, DATA_COL_CODE, code, text_fmt)
        ws.write_string(row_0, DATA_COL_NAME, name, text_fmt)
        ws.write_string(row_0, DATA_COL_TYPE, acct_type, text_fmt)
        ws.write_string(row_0, DATA_COL_LINE, report_line, formats["input_text"])
        # Monthly figures
        values = dat["account_monthly"][i]
        for m, v in enumerate(values):
            ws.write_number(row_0, DATA_COL_M1 + m, v, formats["input_value"])
        # FY total formula
        first_m = col_letter(DATA_COL_M1)
        last_m = col_letter(DATA_COL_M_LAST)
        ws.write_formula(
            row_0, DATA_COL_TOTAL,
            f"=SUM(${first_m}${DATA_ROW_FIRST_ACCT + i}:${last_m}${DATA_ROW_FIRST_ACCT + i})",
            formats["td_bold_right"],
        )

    # Data validation on Report line column: dropdown of the five report lines
    rl_first = DATA_ROW_FIRST_ACCT - 1
    rl_last = DATA_ROW_LAST_ACCT - 1
    ws.data_validation(
        rl_first, DATA_COL_LINE, rl_last, DATA_COL_LINE,
        {
            "validate": "list",
            "source": REPORT_LINES,
            "input_title": "Report line",
            "input_message": "Pick one of Revenue, Cost of sales, Wages, Other opex, D&A.",
        },
    )

    # Column totals row (sanity check)
    ct_row = DATA_ROW_COL_TOTAL - 1
    ws.set_row(ct_row, 22)
    ws.write_string(ct_row, DATA_COL_NAME, "Column totals (sanity check)", formats["td_bold_left"])
    for c in range(DATA_COL_M1, DATA_COL_TOTAL + 1):
        col = col_letter(c)
        ws.write_formula(
            ct_row, c,
            f"=SUM({col}${DATA_ROW_FIRST_ACCT}:{col}${DATA_ROW_LAST_ACCT})",
            formats["td_bold_right"],
        )

    # Report line summary block
    rl_section_row = DATA_ROW_COL_TOTAL + 1  # blank row after totals
    st.write_section_header(
        ws, formats, row=rl_section_row,
        kicker="Reference",
        title="Report line aggregates",
    )

    rl_header_row_0 = DATA_ROW_RL_HEADER - 1
    st.write_header_row(
        ws, formats, row=rl_header_row_0,
        headers=["Report line"] + month_headers + ["FY total"],
        start_col=DATA_COL_LINE, right_align_from=2,
    )

    for i, rl in enumerate(REPORT_LINES):
        row_0 = DATA_ROW_RL_FIRST - 1 + i
        zebra = i % 2 == 0
        text_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(row_0, 20)
        ws.write_string(row_0, DATA_COL_LINE, rl, text_fmt)
        for m in range(NUM_MONTHS):
            ws.write_formula(row_0, DATA_COL_M1 + m, sumifs_month(rl, m), num_fmt)
        ws.write_formula(row_0, DATA_COL_TOTAL, sumifs_fy(rl), num_fmt)

    sc.apply_page_setup(ws, sheet_title="Data")
    return ws


# ── Internal Data Measures sheet ────────────────────────────────────────────

def build_internal_measures(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("Internal Data Measures")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_CODE, DATA_COL_CODE, 30)
    ws.set_column(DATA_COL_NAME, DATA_COL_M_LAST, 11)
    ws.set_column(DATA_COL_TOTAL, DATA_COL_TOTAL, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Inputs that do not come from the P&L",
        title="Internal data measures",
        last_col=LAST_COL + 1,
        explanation=(
            "This sheet holds the non-software inputs the analytical sheets need: working-capital "
            "days outstanding and your department wage allocation. These figures vary by company "
            "and cannot be exported from your accounting software; replace the amber-bordered cells with your own. "
            "Each section below names the analytical tab that consumes the figure."
        ),
    )

    # Section 1: Working capital days
    st.write_section_header(
        ws, formats, row=6,
        kicker="Section 1   Used on: Working Capital tab",
        title="Working capital days outstanding",
    )

    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    headers = ["Metric"] + month_headers + ["Average"]
    st.write_header_row(
        ws, formats, row=IDM_ROW_WC_HEADER - 1, headers=headers,
        start_col=DATA_COL_CODE, right_align_from=2,
    )

    def write_idm_days_row(row_one: int, label: str, values: list[float]) -> None:
        row_0 = row_one - 1
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_CODE, label, formats["input_label"])
        for i, v in enumerate(values):
            ws.write_number(row_0, DATA_COL_NAME + i, v, formats["input_days"])
        first = col_letter(DATA_COL_NAME)
        last = col_letter(DATA_COL_NAME + NUM_MONTHS - 1)
        ws.write_formula(
            row_0, DATA_COL_NAME + NUM_MONTHS,
            f"=AVERAGE({first}{row_one}:{last}{row_one})",
            formats["td_days"],
        )

    write_idm_days_row(IDM_ROW_DSO, "Days Sales Outstanding (DSO)", dat["dso"])
    write_idm_days_row(IDM_ROW_DPO, "Days Payables Outstanding (DPO)", dat["dpo"])
    write_idm_days_row(IDM_ROW_DIO, "Days Inventory Outstanding (DIO)", dat["dio"])

    # CCC computed row
    ccc_row_0 = IDM_ROW_CCC - 1
    ws.set_row(ccc_row_0, 22)
    ws.write_string(ccc_row_0, DATA_COL_CODE, "Cash Conversion Cycle (CCC) = DSO + DIO - DPO", formats["td_bold_left"])
    for m in range(NUM_MONTHS):
        col = col_letter(DATA_COL_NAME + m)
        ws.write_formula(
            ccc_row_0, DATA_COL_NAME + m,
            f"={col}{IDM_ROW_DSO}+{col}{IDM_ROW_DIO}-{col}{IDM_ROW_DPO}",
            formats["td_days"],
        )
    first = col_letter(DATA_COL_NAME)
    last = col_letter(DATA_COL_NAME + NUM_MONTHS - 1)
    ws.write_formula(
        ccc_row_0, DATA_COL_NAME + NUM_MONTHS,
        f"=AVERAGE({first}{IDM_ROW_CCC}:{last}{IDM_ROW_CCC})",
        formats["td_days"],
    )

    # Section 2: Department wage allocation
    st.write_section_header(
        ws, formats, row=14,
        kicker="Section 2   Used on: Wages tab",
        title="Department wage allocation",
    )
    ws.set_row(IDM_ROW_DEPT_HEADER - 1, 26)
    ws.write_string(IDM_ROW_DEPT_HEADER - 1, DATA_COL_CODE, "Department", formats["th"])
    ws.write_string(IDM_ROW_DEPT_HEADER - 1, DATA_COL_NAME, "Share of total wages", formats["th_right"])
    for i, (dept, share) in enumerate(zip(IDM_DEPTS, dat["dept_share"])):
        row_0 = IDM_ROW_DEPT_FIRST - 1 + i
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_CODE, dept, formats["input_label"])
        ws.write_number(row_0, DATA_COL_NAME, share, formats["input_pct"])

    sc.apply_page_setup(ws, sheet_title="Internal Data Measures")
    return ws


# ── Analytical sheets ───────────────────────────────────────────────────────

def build_headline(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("Headline")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 14
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 13, 10)
    ws.set_column(14, 14, 14)
    ws.set_column(15, 15, 2)

    last_month = dat["months"][-1].strftime("%B %Y")
    sc.write_hero_band(
        ws, formats,
        kicker=f"Month ending {last_month}",
        title="Monthly board read-out",
        last_col=LAST_COL + 1,
        explanation=(
            "This sheet shows the headline metrics for the most recent month at the top, your "
            "written commentary in the middle, and a 12-month trend table at the bottom. "
            "Acronyms used below: Gross profit (GP), Earnings Before Interest, Tax, "
            "Depreciation and Amortisation (EBITDA), Net Profit After Tax (NPAT)."
        ),
    )

    # KPI cards reference the latest month and prior month aggregated from Data
    last_idx = NUM_MONTHS - 1
    prior_idx = NUM_MONTHS - 2

    def kpi(label: str, report_line: str, anchor_col: int) -> None:
        # Special case: value formulas operate on aggregated SUMIFS or formulas.
        # For Gross profit, EBITDA and NPAT we compute from constituent lines.
        if report_line == "Gross profit":
            value_f = f"={sumifs_month('Revenue', last_idx)[1:]}-{sumifs_month('Cost of sales', last_idx)[1:]}"
            prior_f = f"={sumifs_month('Revenue', prior_idx)[1:]}-{sumifs_month('Cost of sales', prior_idx)[1:]}"
        elif report_line == "EBITDA":
            value_f = (
                f"={sumifs_month('Revenue', last_idx)[1:]}-{sumifs_month('Cost of sales', last_idx)[1:]}"
                f"-{sumifs_month('Wages', last_idx)[1:]}-{sumifs_month('Other opex', last_idx)[1:]}"
            )
            prior_f = (
                f"={sumifs_month('Revenue', prior_idx)[1:]}-{sumifs_month('Cost of sales', prior_idx)[1:]}"
                f"-{sumifs_month('Wages', prior_idx)[1:]}-{sumifs_month('Other opex', prior_idx)[1:]}"
            )
        elif report_line == "Net profit":
            value_f = (
                f"={sumifs_month('Revenue', last_idx)[1:]}-{sumifs_month('Cost of sales', last_idx)[1:]}"
                f"-{sumifs_month('Wages', last_idx)[1:]}-{sumifs_month('Other opex', last_idx)[1:]}"
                f"-{sumifs_month('D&A', last_idx)[1:]}"
            )
            prior_f = (
                f"={sumifs_month('Revenue', prior_idx)[1:]}-{sumifs_month('Cost of sales', prior_idx)[1:]}"
                f"-{sumifs_month('Wages', prior_idx)[1:]}-{sumifs_month('Other opex', prior_idx)[1:]}"
                f"-{sumifs_month('D&A', prior_idx)[1:]}"
            )
        else:  # Revenue
            value_f = sumifs_month(report_line, last_idx)
            prior_f = sumifs_month(report_line, prior_idx)
        # Use ABS(prior) in the denominator so the percentage direction stays
        # intuitive even when the prior month was a loss (negative). Example:
        # prior = -10, now = -5 → +50% (loss halved), not -50%.
        change_f = (
            f"=IF(({prior_f[1:]})=0,\"n/a\","
            f"\"vs prior month  \"&TEXT((({value_f[1:]})-({prior_f[1:]}))/ABS({prior_f[1:]}),\"+0.0%;-0.0%\"))"
        )
        st.kpi_card(
            ws, formats,
            anchor_row=6, anchor_col=anchor_col, width_cols=3,
            label=label, value_formula=value_f, change_formula=change_f,
            value_kind="aud",
        )

    kpi("Revenue (latest month)", "Revenue", 1)
    kpi("Gross profit", "Gross profit", 4)
    kpi("EBITDA", "EBITDA", 7)
    kpi("Net profit", "Net profit", 10)

    # Commentary block (free text)
    section_row = 11
    st.write_section_header(ws, formats, row=section_row, kicker="Note from the finance lead", title="Commentary")
    ws.set_row(section_row + 2, 100)
    ws.merge_range(
        section_row + 2, 1, section_row + 2, LAST_COL - 1,
        "Replace this text with your three- to five-sentence board commentary.\n\n"
        "Example: Revenue lifted by 4.2 per cent on the prior month driven by stronger "
        "wholesale orders. Gross margin held flat at 45 per cent. EBITDA exceeded the "
        "monthly budget by $24k owing to lower outsourced wages. Cash position remains "
        "comfortable; the next BAS (Business Activity Statement) lodgement falls inside "
        "the upcoming quarter.",
        formats["input_text"],
    )

    # 12-month trend table
    trend_section = section_row + 5
    st.write_section_header(ws, formats, row=trend_section, kicker="Snapshot", title="12-month trend")
    header_row_idx = trend_section + 2
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    headers = ["Series"] + month_headers + ["Trend"]
    st.write_header_row(ws, formats, row=header_row_idx, headers=headers, right_align_from=2)

    trend_rows = [
        ("Revenue", "Revenue", "single"),
        ("Gross profit", "GP", "computed"),
        ("EBITDA", "EBITDA", "computed"),
        ("Net profit", "Net profit", "computed"),
    ]
    for i, (label, key, kind) in enumerate(trend_rows):
        row = header_row_idx + 1 + i
        ws.set_row(row, 22)
        zebra = (i % 2 == 0)
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.write_string(row, 1, label, label_fmt)
        for m_idx in range(NUM_MONTHS):
            col = 2 + m_idx
            if key == "Revenue":
                f_ = sumifs_month("Revenue", m_idx)
            elif key == "GP":
                f_ = f"={sumifs_month('Revenue', m_idx)[1:]}-{sumifs_month('Cost of sales', m_idx)[1:]}"
            elif key == "EBITDA":
                f_ = (
                    f"={sumifs_month('Revenue', m_idx)[1:]}-{sumifs_month('Cost of sales', m_idx)[1:]}"
                    f"-{sumifs_month('Wages', m_idx)[1:]}-{sumifs_month('Other opex', m_idx)[1:]}"
                )
            elif key == "Net profit":
                f_ = (
                    f"={sumifs_month('Revenue', m_idx)[1:]}-{sumifs_month('Cost of sales', m_idx)[1:]}"
                    f"-{sumifs_month('Wages', m_idx)[1:]}-{sumifs_month('Other opex', m_idx)[1:]}"
                    f"-{sumifs_month('D&A', m_idx)[1:]}"
                )
            ws.write_formula(row, col, f_, num_fmt)
        # Sparkline reads the row we just wrote (same Headline sheet)
        spark_cell = f"{col_letter(LAST_COL)}{row + 1}"
        spark_range = (
            f"'Headline'!{col_letter(2)}{row + 1}"
            f":{col_letter(2 + NUM_MONTHS - 1)}{row + 1}"
        )
        st.add_sparkline_line(ws, anchor_cell=spark_cell, range_str=spark_range, color=b.GREEN)

    # Per-tab tie-out checks (placed below the trend table)
    rev_trend_row_1b = header_row_idx + 2   # 1-based row of Revenue trend row
    np_trend_row_1b = header_row_idx + 5    # 1-based row of Net profit trend row
    np_derivation = (
        f"={sumifs_fy('Revenue')[1:]}-{sumifs_fy('Cost of sales')[1:]}"
        f"-{sumifs_fy('Wages')[1:]}-{sumifs_fy('Other opex')[1:]}"
        f"-{sumifs_fy('D&A')[1:]}"
    )
    st.write_checks_block(
        ws, formats,
        row=header_row_idx + len(trend_rows) + 3,
        checks=[
            {
                "name": "Revenue 12-month trend total ties to Data Revenue FY",
                "left": f"=SUM(C{rev_trend_row_1b}:N{rev_trend_row_1b})",
                "right": sumifs_fy("Revenue"),
            },
            {
                "name": "Net profit trend total ties to derived Net profit (Revenue minus all expenses)",
                "left": f"=SUM(C{np_trend_row_1b}:N{np_trend_row_1b})",
                "right": np_derivation,
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Headline")
    return ws


def build_pnl_monthly(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("P&L Monthly")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 26)
    ws.set_column(2, NUM_MONTHS + 1, 11)
    ws.set_column(NUM_MONTHS + 2, NUM_MONTHS + 2, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Last 12 months",
        title="Profit and loss by month",
        last_col=LAST_COL + 1,
        explanation=(
            "Aggregated profit and loss read from the Data sheet by Report line. "
            "Revenue and Gross profit are calculated; Gross margin and EBITDA margin "
            "are derived percentages. EBITDA stands for Earnings Before Interest, Tax, "
            "Depreciation and Amortisation."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet", title="Monthly P&L")

    header_row = section_row + 2
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Line item"] + month_headers + ["FY Total"],
        right_align_from=2,
    )

    def write_line(row: int, label: str, formulas_per_month: list[str], total_formula: str, *, zebra: bool, bold: bool = False, fmt_key: str = "td_right") -> None:
        st.write_data_row(
            ws, formats, row=row, label=label,
            formulas=formulas_per_month + [total_formula],
            zebra=zebra, bold=bold, cell_format=fmt_key,
        )

    rev_m = [sumifs_month("Revenue", m) for m in range(NUM_MONTHS)]
    cogs_m = [sumifs_month("Cost of sales", m) for m in range(NUM_MONTHS)]
    wages_m = [sumifs_month("Wages", m) for m in range(NUM_MONTHS)]
    opex_m = [sumifs_month("Other opex", m) for m in range(NUM_MONTHS)]
    da_m = [sumifs_month("D&A", m) for m in range(NUM_MONTHS)]

    gp_m = [f"={rev_m[m][1:]}-{cogs_m[m][1:]}" for m in range(NUM_MONTHS)]
    eb_m = [f"={gp_m[m][1:]}-{wages_m[m][1:]}-{opex_m[m][1:]}" for m in range(NUM_MONTHS)]
    np_m = [f"={eb_m[m][1:]}-{da_m[m][1:]}" for m in range(NUM_MONTHS)]

    gm_m = [f"=IFERROR(({gp_m[m][1:]})/({rev_m[m][1:]}),0)" for m in range(NUM_MONTHS)]
    em_m = [f"=IFERROR(({eb_m[m][1:]})/({rev_m[m][1:]}),0)" for m in range(NUM_MONTHS)]

    r = header_row + 1
    write_line(r, "Revenue", rev_m, sumifs_fy("Revenue"), zebra=False); r += 1
    # COGS shown as negative for P&L readability
    cogs_neg = [f"=-{x[1:]}" for x in cogs_m]
    cogs_total_neg = f"=-{sumifs_fy('Cost of sales')[1:]}"
    write_line(r, "Cost of sales", cogs_neg, cogs_total_neg, zebra=True); r += 1
    gp_total = f"={sumifs_fy('Revenue')[1:]}-{sumifs_fy('Cost of sales')[1:]}"
    write_line(r, "Gross profit", gp_m, gp_total, zebra=False, bold=True); r += 1
    gm_total = f"=IFERROR(({sumifs_fy('Revenue')[1:]}-{sumifs_fy('Cost of sales')[1:]})/({sumifs_fy('Revenue')[1:]}),0)"
    write_line(r, "Gross margin %", gm_m, gm_total, zebra=True, fmt_key="td_pct"); r += 1
    wages_neg = [f"=-{x[1:]}" for x in wages_m]
    wages_total_neg = f"=-{sumifs_fy('Wages')[1:]}"
    write_line(r, "Wages and salaries", wages_neg, wages_total_neg, zebra=False); r += 1
    opex_neg = [f"=-{x[1:]}" for x in opex_m]
    opex_total_neg = f"=-{sumifs_fy('Other opex')[1:]}"
    write_line(r, "Other operating expenses", opex_neg, opex_total_neg, zebra=True); r += 1
    eb_total = (
        f"={sumifs_fy('Revenue')[1:]}-{sumifs_fy('Cost of sales')[1:]}"
        f"-{sumifs_fy('Wages')[1:]}-{sumifs_fy('Other opex')[1:]}"
    )
    write_line(r, "EBITDA", eb_m, eb_total, zebra=False, bold=True); ebitda_row = r; r += 1
    em_total = f"=IFERROR(({eb_total[1:]})/({sumifs_fy('Revenue')[1:]}),0)"
    write_line(r, "EBITDA margin %", em_m, em_total, zebra=True, fmt_key="td_pct"); r += 1
    da_neg = [f"=-{x[1:]}" for x in da_m]
    da_total_neg = f"=-{sumifs_fy('D&A')[1:]}"
    write_line(r, "Depreciation and amortisation", da_neg, da_total_neg, zebra=False); r += 1
    np_total = f"={eb_total[1:]}-{sumifs_fy('D&A')[1:]}"
    st.write_total_row(ws, formats, row=r, label="Net profit", formulas=np_m + [np_total])
    npat_row = r

    rev_row_one = header_row + 2
    eb_row_one = ebitda_row + 1
    npat_row_one = r + 1
    npat_col_letter = col_letter(2 + NUM_MONTHS)

    # Per-tab tie-out checks. P&L Monthly shows expenses as negatives; all
    # checks below respect that sign convention.
    checks_anchor = r + 2  # leave one blank row after the NPAT total
    checks_end_one = st.write_checks_block(
        ws, formats,
        row=checks_anchor,
        checks=[
            {
                "name": "Revenue FY ties to Data Revenue FY",
                "left": f"=${npat_col_letter}${rev_row_one}",
                "right": sumifs_fy("Revenue"),
            },
            {
                "name": "Net profit FY equals sum of monthly Net profit",
                "left": f"=${npat_col_letter}${npat_row_one}",
                "right": f"=SUM(${col_letter(2)}${npat_row_one}:${col_letter(2 + NUM_MONTHS - 1)}${npat_row_one})",
            },
        ],
    )

    # Move the chart below the checks block
    chart_row_one = checks_end_one + 2

    cats_range = (
        f"='P&L Monthly'!${col_letter(2)}${header_row + 1}"
        f":${col_letter(2 + NUM_MONTHS - 1)}${header_row + 1}"
    )
    series = [
        {
            "name": "Revenue",
            "values": (
                f"='P&L Monthly'!${col_letter(2)}${rev_row_one}"
                f":${col_letter(2 + NUM_MONTHS - 1)}${rev_row_one}"
            ),
            "color": b.CHART_PRIMARY,
        },
        {
            "name": "EBITDA",
            "values": (
                f"='P&L Monthly'!${col_letter(2)}${eb_row_one}"
                f":${col_letter(2 + NUM_MONTHS - 1)}${eb_row_one}"
            ),
            "color": b.CHART_TERTIARY,
        },
    ]
    st.add_line_chart(
        wb, ws,
        title="Revenue and EBITDA trend",
        anchor_cell=f"B{chart_row_one}",
        series=series, cats_range=cats_range,
        width=720, height=300,
    )

    sc.apply_page_setup(ws, sheet_title="P&L Monthly")
    return ws


def build_quarter_comparison(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    """Data-driven quarter-over-quarter view of every P&L line item.

    Replaces the earlier Margin Bridge tab, which relied on a user-set driver
    allocation that just told the user back what they had told the workbook.
    This view shows the prior three months vs the current three months and
    lets the data speak.
    """
    ws = wb.add_worksheet("Quarter Comparison")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 8
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, LAST_COL, 14)
    ws.set_column(LAST_COL + 1, LAST_COL + 1, 2)

    months = dat["months"]
    prior_start_label = months[-6].strftime("%b %y")
    prior_end_label = months[-4].strftime("%b %y")
    curr_start_label = months[-3].strftime("%b %y")
    curr_end_label = months[-1].strftime("%b %y")
    prior_label = f"{prior_start_label} – {prior_end_label}"
    curr_label = f"{curr_start_label} – {curr_end_label}"

    sc.write_hero_band(
        ws, formats,
        kicker="Quarter over quarter",
        title="Quarter comparison",
        last_col=LAST_COL + 2,
        explanation=(
            f"Compares the most recent three months ({curr_label}) to the prior three months "
            f"({prior_label}) for every profit and loss line item. Use the $ change and "
            "% change columns to spot what moved and where to focus your commentary. All "
            "figures are drawn directly from the Data sheet."
        ),
    )

    section_row = 6
    st.write_section_header(
        ws, formats, row=section_row,
        kicker="Drawn from the Data sheet",
        title="Profit and loss line items, prior quarter vs current quarter",
    )

    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=[
            "Line item",
            f"Prior 3 months ({prior_label})",
            f"Current 3 months ({curr_label})",
            "Change $",
            "Change %",
        ],
        right_align_from=2,
    )

    # Helper to build the 3-month sum formula for a given report line, where
    # `start_idx` is the 0-based month index of the FIRST month in the quarter.
    def quarter_sum(report_line: str, start_idx_0: int) -> str:
        parts = "+".join(
            sumifs_month(report_line, i)[1:]
            for i in range(start_idx_0, start_idx_0 + 3)
        )
        return f"=({parts})"

    prior_start = NUM_MONTHS - 6  # months[-6:-3]
    curr_start = NUM_MONTHS - 3   # months[-3:]

    def prior_q(rl: str) -> str:
        return quarter_sum(rl, prior_start)

    def curr_q(rl: str) -> str:
        return quarter_sum(rl, curr_start)

    # Each line: (label, prior_formula, current_formula, format_key, bold, zebra, sign_for_cf)
    # sign_for_cf controls which colour direction CF uses for the % change column:
    #   "favourable" → growth is good (revenue, GP, EBITDA, NPAT)
    #   "unfavourable" → growth is bad (costs)
    #   None → no CF on this row
    sign_revenue = "favourable"
    sign_cost = "unfavourable"

    # Compute formulas
    rev_prior = prior_q("Revenue")
    rev_curr = curr_q("Revenue")
    cogs_prior = prior_q("Cost of sales")
    cogs_curr = curr_q("Cost of sales")
    wages_prior = prior_q("Wages")
    wages_curr = curr_q("Wages")
    opex_prior = prior_q("Other opex")
    opex_curr = curr_q("Other opex")
    da_prior = prior_q("D&A")
    da_curr = curr_q("D&A")

    gp_prior = f"={rev_prior[1:]}-{cogs_prior[1:]}"
    gp_curr = f"={rev_curr[1:]}-{cogs_curr[1:]}"
    eb_prior = f"={rev_prior[1:]}-{cogs_prior[1:]}-{wages_prior[1:]}-{opex_prior[1:]}"
    eb_curr = f"={rev_curr[1:]}-{cogs_curr[1:]}-{wages_curr[1:]}-{opex_curr[1:]}"
    np_prior = f"={eb_prior[1:]}-{da_prior[1:]}"
    np_curr = f"={eb_curr[1:]}-{da_curr[1:]}"

    # Layout each line in turn. P&L Monthly displays costs as negative, but on
    # this tab we keep everything positive to make the prior-vs-current side-by-
    # side cleaner. The Change % uses ABS(prior) to handle negative bases.
    lines = [
        ("Revenue", rev_prior, rev_curr, False, sign_revenue),
        ("Cost of sales", cogs_prior, cogs_curr, False, sign_cost),
        ("Gross profit", gp_prior, gp_curr, True, sign_revenue),
        ("Wages and salaries", wages_prior, wages_curr, False, sign_cost),
        ("Other operating expenses", opex_prior, opex_curr, False, sign_cost),
        ("EBITDA", eb_prior, eb_curr, True, sign_revenue),
        ("Depreciation and amortisation", da_prior, da_curr, False, sign_cost),
        ("Net profit", np_prior, np_curr, True, sign_revenue),
    ]

    first_data_row_one = header_row + 2
    r = header_row + 1
    pct_cf_ranges_favourable = []
    pct_cf_ranges_unfavourable = []
    for i, (label, p_f, c_f, bold, sign) in enumerate(lines):
        ws.set_row(r, 22)
        zebra = i % 2 == 0
        label_fmt = formats["td_bold_left"] if bold else (formats["td_zebra"] if zebra else formats["td"])
        num_fmt = formats["td_bold_right"] if bold else (formats["td_right_zebra"] if zebra else formats["td_right"])
        pct_fmt = formats["td_bold_pct"] if bold else (formats["td_pct_zebra"] if zebra else formats["td_pct"])

        ws.write_string(r, 1, label, label_fmt)
        ws.write_formula(r, 2, p_f, num_fmt)
        ws.write_formula(r, 3, c_f, num_fmt)
        # $ change = current - prior
        change_dollar = f"=D{r + 1}-C{r + 1}"
        ws.write_formula(r, 4, change_dollar, num_fmt)
        # % change = (current - prior) / ABS(prior); IFERROR handles div/0
        change_pct = f"=IFERROR((D{r + 1}-C{r + 1})/ABS(C{r + 1}),0)"
        ws.write_formula(r, 5, change_pct, pct_fmt)

        pct_cell = f"F{r + 1}"
        if sign == "favourable":
            pct_cf_ranges_favourable.append(pct_cell)
        elif sign == "unfavourable":
            pct_cf_ranges_unfavourable.append(pct_cell)

        r += 1

    # Conditional formatting on the % change column, per-cell so the sign
    # direction can vary by row (revenue/GP/EBITDA: high is green; costs: high is red).
    if pct_cf_ranges_favourable:
        rng = ",".join(pct_cf_ranges_favourable)
        st.add_three_color_scale(ws, rng, favourable_high=True)
    if pct_cf_ranges_unfavourable:
        rng = ",".join(pct_cf_ranges_unfavourable)
        st.add_three_color_scale(ws, rng, favourable_high=False)

    legend_row = r + 1
    st.write_cf_legend(
        ws, formats, row=legend_row, col=1,
        favourable_high=True,
        metric_label="Change % on revenue, GP, EBITDA, Net profit",
    )
    st.write_cf_legend(
        ws, formats, row=legend_row + 1, col=1,
        favourable_high=False,
        metric_label="Change % on costs (lower is better)",
    )

    # Per-tab tie-out checks
    last_data_row_one = r  # 1-based row of last line item ("Net profit")
    # Revenue current Q tie back to Data: sum of Sales (Revenue report line) for last 3 months
    rev_data_curr_q = f"={curr_q('Revenue')[1:]}"
    checks_end_one = st.write_checks_block(
        ws, formats,
        row=legend_row + 3,
        checks=[
            {
                "name": "Current quarter Revenue ties to Data Revenue for the last three months",
                "left": f"=$D${first_data_row_one}",
                "right": rev_data_curr_q,
            },
            {
                "name": "Net profit current quarter equals Revenue minus all expenses for the same quarter",
                "left": f"=$D${last_data_row_one}",
                "right": f"={curr_q('Revenue')[1:]}-{curr_q('Cost of sales')[1:]}"
                          f"-{curr_q('Wages')[1:]}-{curr_q('Other opex')[1:]}-{curr_q('D&A')[1:]}",
            },
        ],
    )

    # Chart: side-by-side bars of prior vs current for each line item
    chart_anchor = f"B{checks_end_one + 2}"
    cats_range = f"='Quarter Comparison'!$B${first_data_row_one}:$B${last_data_row_one}"
    series = [
        {
            "name": f"Prior 3 months",
            "values": f"='Quarter Comparison'!$C${first_data_row_one}:$C${last_data_row_one}",
            "color": b.GREEN_SPECTRUM[2],
        },
        {
            "name": f"Current 3 months",
            "values": f"='Quarter Comparison'!$D${first_data_row_one}:$D${last_data_row_one}",
            "color": b.CHART_PRIMARY,
        },
    ]
    st.add_column_chart(
        wb, ws,
        title="Prior quarter vs current quarter by line item",
        anchor_cell=chart_anchor,
        series=series, cats_range=cats_range,
        width=720, height=320,
        show_legend=True,
    )

    sc.apply_page_setup(ws, sheet_title="Quarter Comparison")
    return ws


def build_wages(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("Wages")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, NUM_MONTHS + 1, 11)
    ws.set_column(NUM_MONTHS + 2, NUM_MONTHS + 2, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Headcount and payroll exposure",
        title="Wages by department",
        last_col=LAST_COL + 1,
        explanation=(
            "Total Wages is read from the Data sheet (sum of all accounts classified as Wages). "
            "Department breakdown applies the share % set on the Internal Data Measures sheet. "
            "Payroll-to-revenue ratio is total Wages divided by Revenue per month."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet", title="Department wages by month")

    header_row = section_row + 2
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Department"] + month_headers + ["12-month total"],
        right_align_from=2,
    )

    r = header_row + 1
    for i, dept in enumerate(IDM_DEPTS):
        zebra = i % 2 == 0
        share_cell = f"'Internal Data Measures'!$C${IDM_ROW_DEPT_FIRST + i}"
        formulas = [f"=({sumifs_month('Wages', m)[1:]})*{share_cell}" for m in range(NUM_MONTHS)]
        total_f = f"=({sumifs_fy('Wages')[1:]})*{share_cell}"
        st.write_data_row(ws, formats, row=r, label=dept, formulas=formulas + [total_f], zebra=zebra, cell_format="td_right")
        r += 1

    total_formulas = [sumifs_month("Wages", m) for m in range(NUM_MONTHS)]
    total_total = sumifs_fy("Wages")
    st.write_total_row(ws, formats, row=r, label="Total wages", formulas=total_formulas + [total_total])
    r += 2

    st.write_section_header(ws, formats, row=r, kicker="Productivity", title="Payroll to revenue ratio")
    r += 2
    ratio_header_row = r
    st.write_header_row(
        ws, formats, row=ratio_header_row,
        headers=["Metric"] + month_headers + ["FY average"],
        right_align_from=2,
    )
    r += 1
    ratio_formulas = [
        f"=IFERROR(({sumifs_month('Wages', m)[1:]})/({sumifs_month('Revenue', m)[1:]}),0)"
        for m in range(NUM_MONTHS)
    ]
    ratio_total = f"=IFERROR(({sumifs_fy('Wages')[1:]})/({sumifs_fy('Revenue')[1:]}),0)"
    st.write_data_row(ws, formats, row=r, label="Payroll / revenue", formulas=ratio_formulas + [ratio_total], zebra=False, bold=True, cell_format="td_pct")

    ratio_row_one = r + 1
    cf_range = (
        f"{col_letter(2)}{ratio_row_one}"
        f":{col_letter(2 + NUM_MONTHS - 1)}{ratio_row_one}"
    )
    st.add_three_color_scale(ws, cf_range, favourable_high=False)

    legend_row = r + 2
    st.write_cf_legend(ws, formats, row=legend_row, col=1, favourable_high=False, metric_label="Payroll / revenue")

    # Per-tab tie-out checks
    fy_col = col_letter(2 + NUM_MONTHS)
    first_dept_row_1b = header_row + 2
    last_dept_row_1b = header_row + 1 + len(IDM_DEPTS)
    total_wages_row_1b = last_dept_row_1b + 1
    idm_dept_sum = (
        f"=SUM('Internal Data Measures'!$C${IDM_ROW_DEPT_FIRST}"
        f":$C${IDM_ROW_DEPT_FIRST + len(IDM_DEPTS) - 1})"
    )
    st.write_checks_block(
        ws, formats,
        row=legend_row + 3,
        checks=[
            {
                "name": "Sum of department FY totals ties to Total wages FY",
                "left": f"=SUM(${fy_col}${first_dept_row_1b}:${fy_col}${last_dept_row_1b})",
                "right": f"=${fy_col}${total_wages_row_1b}",
            },
            {
                "name": "Department wage shares sum to 100 per cent",
                "left": idm_dept_sum,
                "right": "=1",
                "is_pct": True,
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Wages")
    return ws


def build_working_capital(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("Working Capital")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 30)
    ws.set_column(2, NUM_MONTHS + 1, 11)
    ws.set_column(NUM_MONTHS + 2, NUM_MONTHS + 2, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Days outstanding and cash conversion",
        title="Working capital",
        last_col=LAST_COL + 1,
        explanation=(
            "Days Sales Outstanding (DSO) = average days to collect from customers. "
            "Days Payables Outstanding (DPO) = average days you take to pay suppliers. "
            "Days Inventory Outstanding (DIO) = average days stock sits before sale. "
            "Cash Conversion Cycle (CCC) = DSO + DIO - DPO; the lower the better."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from Internal Data Measures", title="Days outstanding by month")

    header_row = section_row + 2
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Metric"] + month_headers + ["Average"],
        right_align_from=2,
    )

    metric_rows = [
        ("Days Sales Outstanding (DSO)", IDM_ROW_DSO, False),
        ("Days Payables Outstanding (DPO)", IDM_ROW_DPO, True),
        ("Days Inventory Outstanding (DIO)", IDM_ROW_DIO, False),
    ]
    r = header_row + 1
    first_metric_row_one = r + 1
    for label, idm_row, zebra in metric_rows:
        formulas = [
            f"='Internal Data Measures'!{col_letter(DATA_COL_NAME + m)}{idm_row}"
            for m in range(NUM_MONTHS)
        ]
        first = col_letter(DATA_COL_NAME)
        last = col_letter(DATA_COL_NAME + NUM_MONTHS - 1)
        avg_f = f"=AVERAGE('Internal Data Measures'!{first}{idm_row}:{last}{idm_row})"
        st.write_data_row(ws, formats, row=r, label=label, formulas=formulas + [avg_f], zebra=zebra, cell_format="td_days")
        r += 1

    ccc_formulas = [
        f"='Internal Data Measures'!{col_letter(DATA_COL_NAME + m)}{IDM_ROW_CCC}"
        for m in range(NUM_MONTHS)
    ]
    first = col_letter(DATA_COL_NAME)
    last = col_letter(DATA_COL_NAME + NUM_MONTHS - 1)
    ccc_avg = f"=AVERAGE('Internal Data Measures'!{first}{IDM_ROW_CCC}:{last}{IDM_ROW_CCC})"
    st.write_total_row(ws, formats, row=r, label="Cash Conversion Cycle (CCC)", formulas=ccc_formulas + [ccc_avg], cell_format="total_days")
    ccc_row_one = r + 1
    r += 2

    dso_one = first_metric_row_one
    dpo_one = first_metric_row_one + 1
    dio_one = first_metric_row_one + 2
    st.add_three_color_scale(ws, f"C{dso_one}:N{dso_one}", favourable_high=False)
    st.add_three_color_scale(ws, f"C{dpo_one}:N{dpo_one}", favourable_high=True)
    st.add_three_color_scale(ws, f"C{dio_one}:N{dio_one}", favourable_high=False)

    st.write_cf_legend(ws, formats, row=r, col=1, favourable_high=False, metric_label="DSO and DIO (lower is better)")
    st.write_cf_legend(ws, formats, row=r + 1, col=1, favourable_high=True, metric_label="DPO (higher is better)")

    # Per-tab tie-out checks. Check that CCC for the latest month equals
    # DSO + DIO - DPO computed from the same monthly cells on this sheet.
    last_month_col = col_letter(2 + NUM_MONTHS - 1)  # column N
    dso_row_1b = dso_one
    dpo_row_1b = dpo_one
    dio_row_1b = dio_one
    avg_col_letter = col_letter(2 + NUM_MONTHS)  # column O
    checks_end_one = st.write_checks_block(
        ws, formats,
        row=r + 3,
        checks=[
            {
                "name": "Latest month CCC equals DSO plus DIO minus DPO",
                "left": f"=${last_month_col}${ccc_row_one}",
                "right": f"=${last_month_col}${dso_row_1b}+${last_month_col}${dio_row_1b}-${last_month_col}${dpo_row_1b}",
            },
            {
                "name": "Average CCC ties to mean of monthly CCC values",
                "left": f"=${avg_col_letter}${ccc_row_one}",
                "right": f"=AVERAGE(${col_letter(2)}${ccc_row_one}:${last_month_col}${ccc_row_one})",
            },
        ],
    )

    chart_anchor = f"B{checks_end_one + 2}"
    cats_range = f"='Working Capital'!${col_letter(2)}${header_row + 1}:${col_letter(2 + NUM_MONTHS - 1)}${header_row + 1}"
    series = [{
        "name": "Cash Conversion Cycle (days)",
        "values": f"='Working Capital'!${col_letter(2)}${ccc_row_one}:${col_letter(2 + NUM_MONTHS - 1)}${ccc_row_one}",
        "color": b.CHART_PRIMARY,
    }]
    st.add_line_chart(
        wb, ws,
        title="Cash conversion cycle by month",
        anchor_cell=chart_anchor,
        series=series, cats_range=cats_range,
        width=720, height=280,
    )

    sc.apply_page_setup(ws, sheet_title="Working Capital")
    return ws


# ── Main ────────────────────────────────────────────────────────────────────

def main() -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()

    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)

    sc.add_cover_sheet(
        wb, formats,
        workbook_title=WORKBOOK_TITLE,
        workbook_kicker=WORKBOOK_KICKER,
        workbook_id=WORKBOOK_ID,
        version=b.WORKBOOK_VERSION,
        build_date=date.today(),
        how_to_use=HOW_TO_USE,
        target_user=TARGET_USER,
        example_profile=EXAMPLE_PROFILE,
        inputs_required=INPUTS_REQUIRED,
    )
    build_headline(wb, formats, dat)
    build_pnl_monthly(wb, formats, dat)
    build_quarter_comparison(wb, formats, dat)
    build_wages(wb, formats, dat)
    build_working_capital(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    build_internal_measures(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)

    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
