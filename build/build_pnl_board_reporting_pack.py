"""Build the Board Reporting Pack workbook.

A compact, board-facing monthly pack: one-page summary with KPIs and
commentary, supporting detailed P&L, monthly cash by bank account, and a
commentary tab. Designed to be printed or screen-shared as-is.

Output:
  C:\\dev\\lyros-workbooks\\library\\pnl
        \\lyros_lib_pnl_board_reporting_pack.xlsx
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\4001 - Board Reporting Pack.xlsx")

WORKBOOK_ID = "4001"
WORKBOOK_TITLE = "Board Reporting Pack"
WORKBOOK_KICKER = "Board-ready monthly pack for founder-CEOs"
TARGET_USER = (
    "Founder-CEO, in-house Finance Controller, or fractional CFO preparing "
    "a monthly pack for an external board."
)
HOW_TO_USE = [
    "Open the Data sheet and paste your Profit and Loss by Month export.",
    "Open Internal Data Measures and update the monthly closing cash balance per bank account.",
    "The One Page sheet is the board-facing summary; the P&L Detail, Cash Position, and Commentary sheets sit behind it.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "Multi-channel SME with operating, savings, and trust accounts"),
    ("REVENUE SCALE", "Circa $4M annual"),
    ("BOARD CADENCE", "Monthly board meeting with a one-page pack and supporting tabs"),
    ("BANK ACCOUNTS", "Operating, Savings, and Trust"),
]

INPUTS_REQUIRED = [
    ("Profit and loss by month", "Data tab (paste P&L by Month export)"),
    ("Cash balance per bank account by month", "Internal Data Measures tab → drives Cash Position tab"),
    ("Board commentary", "Commentary tab (free-text input)"),
]


NUM_MONTHS = 12

DATA_COL_CODE = 1
DATA_COL_NAME = 2
DATA_COL_TYPE = 3
DATA_COL_LINE = 4
DATA_COL_M1 = 5
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1
DATA_COL_TOTAL = DATA_COL_M_LAST + 1

DATA_ROW_HEADER = 9
DATA_ROW_FIRST_ACCT = 10
DATA_ROW_LAST_ACCT = DATA_ROW_FIRST_ACCT + 16

REPORT_LINES = ["Revenue", "Cost of sales", "Wages", "Other opex", "D&A"]

ACCOUNTS = [
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

# Internal Data Measures: cash by bank account by month
BANK_ACCOUNTS = ["Operating", "Savings", "Trust"]
IDM_ROW_CASH_HEADER = 8
IDM_ROW_CASH_FIRST = 9
IDM_ROW_CASH_LAST = IDM_ROW_CASH_FIRST + len(BANK_ACCOUNTS) - 1
IDM_ROW_CASH_TOTAL = IDM_ROW_CASH_LAST + 1


def col_letter(zero_based: int) -> str:
    s = ""
    n = zero_based
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def sumifs_month(report_line: str, month_idx_0: int) -> str:
    sum_col = col_letter(DATA_COL_M1 + month_idx_0)
    crit_col = col_letter(DATA_COL_LINE)
    return (
        f"=SUMIFS('Data'!${sum_col}${DATA_ROW_FIRST_ACCT}:${sum_col}${DATA_ROW_LAST_ACCT},"
        f"'Data'!${crit_col}${DATA_ROW_FIRST_ACCT}:${crit_col}${DATA_ROW_LAST_ACCT},"
        f"\"{report_line}\")"
    )


def sumifs_fy(report_line: str) -> str:
    sum_col = col_letter(DATA_COL_TOTAL)
    crit_col = col_letter(DATA_COL_LINE)
    return (
        f"=SUMIFS('Data'!${sum_col}${DATA_ROW_FIRST_ACCT}:${sum_col}${DATA_ROW_LAST_ACCT},"
        f"'Data'!${crit_col}${DATA_ROW_FIRST_ACCT}:${crit_col}${DATA_ROW_LAST_ACCT},"
        f"\"{report_line}\")"
    )


def _synthetic_data() -> dict:
    rng_rev = d.make_rng("brp-rev")
    rng_cogs = d.make_rng("brp-cogs")
    rng_op = d.make_rng("brp-op")
    rng_wg = d.make_rng("brp-wg")

    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    revenue = d.revenue_series(NUM_MONTHS, rng_rev, base=320_000)
    cogs = d.cogs_series(revenue, rng_cogs, gm_pct=0.44)
    opex = d.opex_series(NUM_MONTHS, rng_op, base=82_000)
    wages = d.wages_series(revenue, rng_wg, ratio=0.21)
    da = [round(r * 0.04, 0) for r in revenue]

    totals = {"revenue": revenue, "cogs": cogs, "opex": opex, "wages": wages, "da": da}
    account_monthly: list[list[float]] = []
    for code, name, _t, _rl, (key, share) in ACCOUNTS:
        rng_a = d.make_rng(f"brp-{code}")
        row = [round(totals[key][m] * (share + rng_a.uniform(-0.015, 0.015)), 0) for m in range(NUM_MONTHS)]
        account_monthly.append(row)

    # Cash by bank account by month — synthetic closing balances with mild trend
    rng_op_bal = d.make_rng("brp-op-bal")
    rng_sav = d.make_rng("brp-sav")
    rng_tr = d.make_rng("brp-tr")
    cash_operating = [round(220_000 + 8_000 * i + rng_op_bal.uniform(-12_000, 12_000), 0) for i in range(NUM_MONTHS)]
    cash_savings = [round(380_000 + 4_000 * i + rng_sav.uniform(-8_000, 8_000), 0) for i in range(NUM_MONTHS)]
    cash_trust = [round(60_000 + 1_500 * i + rng_tr.uniform(-3_000, 3_000), 0) for i in range(NUM_MONTHS)]

    return {
        "months": months,
        "account_monthly": account_monthly,
        "cash_by_bank": [cash_operating, cash_savings, cash_trust],
    }


def build_data_sheet(wb, formats, dat):
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
            "Paste your Profit and Loss by Month export below. Each row is one "
            "account; the Report line column drives the analytical aggregation. The cash "
            "balance per bank account sits on the Internal Data Measures tab; cash is not "
            "exported from a P&L report."
        ),
    )

    st.write_section_header(ws, formats, row=6, kicker="Paste from your accounting software", title="Profit and loss accounts")
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    headers = ["Code", "Account name", "Account type", "Report line"] + month_headers + ["FY total"]
    st.write_header_row(ws, formats, row=DATA_ROW_HEADER - 1, headers=headers, start_col=DATA_COL_CODE, right_align_from=5)

    for i, (code, name, acct_type, report_line, _) in enumerate(ACCOUNTS):
        row_0 = DATA_ROW_FIRST_ACCT - 1 + i
        zebra = i % 2 == 0
        text_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(row_0, 20)
        ws.write_string(row_0, DATA_COL_CODE, code, text_fmt)
        ws.write_string(row_0, DATA_COL_NAME, name, text_fmt)
        ws.write_string(row_0, DATA_COL_TYPE, acct_type, text_fmt)
        ws.write_string(row_0, DATA_COL_LINE, report_line, formats["input_text"])
        for m, v in enumerate(dat["account_monthly"][i]):
            ws.write_number(row_0, DATA_COL_M1 + m, v, formats["input_value"])
        first_m = col_letter(DATA_COL_M1)
        last_m = col_letter(DATA_COL_M_LAST)
        ws.write_formula(row_0, DATA_COL_TOTAL,
                         f"=SUM(${first_m}${DATA_ROW_FIRST_ACCT + i}:${last_m}${DATA_ROW_FIRST_ACCT + i})",
                         formats["td_bold_right"])

    ws.data_validation(DATA_ROW_FIRST_ACCT - 1, DATA_COL_LINE,
                       DATA_ROW_LAST_ACCT - 1, DATA_COL_LINE,
                       {"validate": "list", "source": REPORT_LINES})

    sc.apply_page_setup(ws, sheet_title="Data")
    return ws


def build_internal_measures(wb, formats, dat):
    ws = wb.add_worksheet("Internal Data Measures")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_CODE, DATA_COL_CODE, 26)
    ws.set_column(DATA_COL_NAME, DATA_COL_M_LAST, 11)
    ws.set_column(DATA_COL_TOTAL, DATA_COL_TOTAL, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Inputs that do not come from the P&L",
        title="Internal data measures",
        last_col=LAST_COL + 1,
        explanation=(
            "Cash balance by bank account is an input the Profit and Loss report does not "
            "provide; pull it from the Balance Sheet report or your bank reconciliation. "
            "Replace the amber-bordered figures with your own."
        ),
    )

    st.write_section_header(
        ws, formats, row=6,
        kicker="Section 1   Used on: Cash Position tab",
        title="Monthly closing cash balance per bank account",
    )

    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    headers = ["Bank account"] + month_headers + ["Latest"]
    st.write_header_row(ws, formats, row=IDM_ROW_CASH_HEADER - 1, headers=headers,
                        start_col=DATA_COL_CODE, right_align_from=2)

    for i, account in enumerate(BANK_ACCOUNTS):
        row_0 = IDM_ROW_CASH_FIRST - 1 + i
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_CODE, account, formats["input_label"])
        for m, v in enumerate(dat["cash_by_bank"][i]):
            ws.write_number(row_0, DATA_COL_NAME + m, v, formats["input_value"])
        last_col_letter = col_letter(DATA_COL_NAME + NUM_MONTHS - 1)
        ws.write_formula(row_0, DATA_COL_NAME + NUM_MONTHS,
                         f"={last_col_letter}{IDM_ROW_CASH_FIRST + i}",
                         formats["td_bold_right"])

    # Total cash row
    ws.set_row(IDM_ROW_CASH_TOTAL - 1, 24)
    ws.write_string(IDM_ROW_CASH_TOTAL - 1, DATA_COL_CODE, "Total cash", formats["total_left"])
    for m in range(NUM_MONTHS):
        col = col_letter(DATA_COL_NAME + m)
        ws.write_formula(IDM_ROW_CASH_TOTAL - 1, DATA_COL_NAME + m,
                         f"=SUM({col}{IDM_ROW_CASH_FIRST}:{col}{IDM_ROW_CASH_LAST})",
                         formats["total_right"])
    last_col_letter = col_letter(DATA_COL_NAME + NUM_MONTHS - 1)
    ws.write_formula(IDM_ROW_CASH_TOTAL - 1, DATA_COL_NAME + NUM_MONTHS,
                     f"={last_col_letter}{IDM_ROW_CASH_TOTAL}",
                     formats["total_right"])

    sc.apply_page_setup(ws, sheet_title="Internal Data Measures")
    return ws


def build_one_page(wb, formats, dat):
    ws = wb.add_worksheet("One Page")
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
        title="Board reporting pack",
        last_col=LAST_COL + 1,
        explanation=(
            "Board-facing summary for the most recent month. Headline KPIs and prior-month "
            "comparison at the top, 12-month trend below, total cash position from the Internal "
            "Data Measures tab on the right. Acronyms: Gross Profit (GP), Earnings Before Interest, "
            "Tax, Depreciation and Amortisation (EBITDA), Net Profit After Tax (NPAT)."
        ),
    )

    last_idx = NUM_MONTHS - 1
    prior_idx = NUM_MONTHS - 2

    def kpi(label, anchor_col, key, value_kind="aud"):
        if key == "Revenue":
            value_f = sumifs_month("Revenue", last_idx)
            prior_f = sumifs_month("Revenue", prior_idx)
        elif key == "GP":
            value_f = f"={sumifs_month('Revenue', last_idx)[1:]}-{sumifs_month('Cost of sales', last_idx)[1:]}"
            prior_f = f"={sumifs_month('Revenue', prior_idx)[1:]}-{sumifs_month('Cost of sales', prior_idx)[1:]}"
        elif key == "EBITDA":
            value_f = (f"={sumifs_month('Revenue', last_idx)[1:]}-{sumifs_month('Cost of sales', last_idx)[1:]}"
                       f"-{sumifs_month('Wages', last_idx)[1:]}-{sumifs_month('Other opex', last_idx)[1:]}")
            prior_f = (f"={sumifs_month('Revenue', prior_idx)[1:]}-{sumifs_month('Cost of sales', prior_idx)[1:]}"
                       f"-{sumifs_month('Wages', prior_idx)[1:]}-{sumifs_month('Other opex', prior_idx)[1:]}")
        elif key == "Cash":
            # Total cash from Internal Data Measures: cell at last data column of total row
            last_col_letter = col_letter(DATA_COL_NAME + NUM_MONTHS - 1)
            prior_col_letter = col_letter(DATA_COL_NAME + NUM_MONTHS - 2)
            value_f = f"='Internal Data Measures'!{last_col_letter}{IDM_ROW_CASH_TOTAL}"
            prior_f = f"='Internal Data Measures'!{prior_col_letter}{IDM_ROW_CASH_TOTAL}"
        change_f = (
            f"=IF(({prior_f[1:]})=0,\"n/a\","
            f"\"vs prior month  \"&TEXT((({value_f[1:]})-({prior_f[1:]}))/ABS({prior_f[1:]}),\"+0.0%;-0.0%\"))"
        )
        st.kpi_card(ws, formats, anchor_row=6, anchor_col=anchor_col, width_cols=3,
                    label=label, value_formula=value_f, change_formula=change_f, value_kind=value_kind)

    kpi("Revenue", 1, "Revenue")
    kpi("Gross profit", 4, "GP")
    kpi("EBITDA", 7, "EBITDA")
    kpi("Total cash", 10, "Cash")

    # Mini P&L table
    section_row = 11
    st.write_section_header(ws, formats, row=section_row, kicker="Snapshot", title="12-month trend")
    header_row_idx = section_row + 2
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    headers = ["Series"] + month_headers + ["Trend"]
    st.write_header_row(ws, formats, row=header_row_idx, headers=headers, right_align_from=2)

    trend_rows = [
        ("Revenue", "Revenue"),
        ("Gross profit", "GP"),
        ("EBITDA", "EBITDA"),
        ("Net profit", "NPAT"),
        ("Total cash", "Cash"),
    ]

    for i, (label, key) in enumerate(trend_rows):
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
                f_ = (f"={sumifs_month('Revenue', m_idx)[1:]}-{sumifs_month('Cost of sales', m_idx)[1:]}"
                      f"-{sumifs_month('Wages', m_idx)[1:]}-{sumifs_month('Other opex', m_idx)[1:]}")
            elif key == "NPAT":
                f_ = (f"={sumifs_month('Revenue', m_idx)[1:]}-{sumifs_month('Cost of sales', m_idx)[1:]}"
                      f"-{sumifs_month('Wages', m_idx)[1:]}-{sumifs_month('Other opex', m_idx)[1:]}"
                      f"-{sumifs_month('D&A', m_idx)[1:]}")
            elif key == "Cash":
                col_l = col_letter(DATA_COL_NAME + m_idx)
                f_ = f"='Internal Data Measures'!{col_l}{IDM_ROW_CASH_TOTAL}"
            ws.write_formula(row, col, f_, num_fmt)
        spark_cell = f"{col_letter(LAST_COL)}{row + 1}"
        spark_range = (
            f"'One Page'!{col_letter(2)}{row + 1}:{col_letter(2 + NUM_MONTHS - 1)}{row + 1}"
        )
        st.add_sparkline_line(ws, anchor_cell=spark_cell, range_str=spark_range, color=b.GREEN)

    # Tie-out checks
    rev_trend_row_1b = header_row_idx + 2
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
                "name": "Total cash latest month ties to Internal Data Measures total",
                "left": f"=N{header_row_idx + 1 + 5}",
                "right": (
                    f"='Internal Data Measures'!{col_letter(DATA_COL_NAME + NUM_MONTHS - 1)}"
                    f"{IDM_ROW_CASH_TOTAL}"
                ),
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="One Page")
    return ws


def build_pnl_detail(wb, formats, dat):
    ws = wb.add_worksheet("P&L Detail")
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
        title="Detailed profit and loss",
        last_col=LAST_COL + 1,
        explanation=(
            "Full monthly P&L drawn from the Data sheet using SUMIFS by Report line. "
            "Sits behind the One Page summary for board members who want the underlying detail."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet", title="Monthly P&L")
    header_row = section_row + 2
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Line item"] + month_headers + ["FY Total"], right_align_from=2)

    rev_m = [sumifs_month("Revenue", m) for m in range(NUM_MONTHS)]
    cogs_m = [sumifs_month("Cost of sales", m) for m in range(NUM_MONTHS)]
    wages_m = [sumifs_month("Wages", m) for m in range(NUM_MONTHS)]
    opex_m = [sumifs_month("Other opex", m) for m in range(NUM_MONTHS)]
    da_m = [sumifs_month("D&A", m) for m in range(NUM_MONTHS)]
    gp_m = [f"={rev_m[m][1:]}-{cogs_m[m][1:]}" for m in range(NUM_MONTHS)]
    eb_m = [f"={gp_m[m][1:]}-{wages_m[m][1:]}-{opex_m[m][1:]}" for m in range(NUM_MONTHS)]
    np_m = [f"={eb_m[m][1:]}-{da_m[m][1:]}" for m in range(NUM_MONTHS)]

    rev_fy = sumifs_fy("Revenue")
    cogs_fy = sumifs_fy("Cost of sales")
    wages_fy = sumifs_fy("Wages")
    opex_fy = sumifs_fy("Other opex")
    da_fy = sumifs_fy("D&A")
    gp_fy = f"={rev_fy[1:]}-{cogs_fy[1:]}"
    eb_fy = f"={gp_fy[1:]}-{wages_fy[1:]}-{opex_fy[1:]}"
    np_fy = f"={eb_fy[1:]}-{da_fy[1:]}"

    r = header_row + 1
    st.write_data_row(ws, formats, row=r, label="Revenue", formulas=rev_m + [rev_fy], zebra=False, cell_format="td_right"); r += 1
    st.write_data_row(ws, formats, row=r, label="Cost of sales", formulas=[f"=-{x[1:]}" for x in cogs_m] + [f"=-{cogs_fy[1:]}"], zebra=True, cell_format="td_right"); r += 1
    st.write_data_row(ws, formats, row=r, label="Gross profit", formulas=gp_m + [gp_fy], zebra=False, bold=True, cell_format="td_right"); gp_row_1b = r + 1; r += 1
    st.write_data_row(ws, formats, row=r, label="Wages and salaries", formulas=[f"=-{x[1:]}" for x in wages_m] + [f"=-{wages_fy[1:]}"], zebra=True, cell_format="td_right"); r += 1
    st.write_data_row(ws, formats, row=r, label="Other opex", formulas=[f"=-{x[1:]}" for x in opex_m] + [f"=-{opex_fy[1:]}"], zebra=False, cell_format="td_right"); r += 1
    st.write_data_row(ws, formats, row=r, label="EBITDA", formulas=eb_m + [eb_fy], zebra=True, bold=True, cell_format="td_right"); r += 1
    st.write_data_row(ws, formats, row=r, label="D&A", formulas=[f"=-{x[1:]}" for x in da_m] + [f"=-{da_fy[1:]}"], zebra=False, cell_format="td_right"); r += 1
    st.write_total_row(ws, formats, row=r, label="Net profit", formulas=np_m + [np_fy]); npat_row_1b = r + 1

    fy_col_letter = col_letter(2 + NUM_MONTHS)
    st.write_checks_block(
        ws, formats,
        row=r + 2,
        checks=[
            {
                "name": "Revenue FY ties to Data Revenue FY",
                "left": f"=${fy_col_letter}${header_row + 2}",
                "right": rev_fy,
            },
            {
                "name": "Net profit FY equals sum of monthly Net profit",
                "left": f"=${fy_col_letter}${npat_row_1b}",
                "right": f"=SUM({col_letter(2)}{npat_row_1b}:{col_letter(2 + NUM_MONTHS - 1)}{npat_row_1b})",
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="P&L Detail")
    return ws


def build_cash_position(wb, formats, dat):
    ws = wb.add_worksheet("Cash Position")
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
        kicker="Closing balance per bank account",
        title="Cash position",
        last_col=LAST_COL + 1,
        explanation=(
            "Monthly closing cash balance for each bank account, with a Total cash row "
            "and a trend chart. Inputs are set on the Internal Data Measures tab."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from Internal Data Measures",
                            title="Closing cash balance by bank account")
    header_row = section_row + 2
    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Bank account"] + month_headers + ["Latest"], right_align_from=2)

    r = header_row + 1
    first_bank_row_1b = r + 1
    for i, account in enumerate(BANK_ACCOUNTS):
        zebra = i % 2 == 0
        idm_row = IDM_ROW_CASH_FIRST + i
        formulas = [
            f"='Internal Data Measures'!{col_letter(DATA_COL_NAME + m)}{idm_row}"
            for m in range(NUM_MONTHS)
        ]
        latest_f = f"='Internal Data Measures'!{col_letter(DATA_COL_NAME + NUM_MONTHS - 1)}{idm_row}"
        st.write_data_row(ws, formats, row=r, label=account, formulas=formulas + [latest_f],
                          zebra=zebra, cell_format="td_right")
        r += 1
    last_bank_row_1b = r

    # Total row
    total_formulas = [
        f"='Internal Data Measures'!{col_letter(DATA_COL_NAME + m)}{IDM_ROW_CASH_TOTAL}"
        for m in range(NUM_MONTHS)
    ]
    total_latest = f"='Internal Data Measures'!{col_letter(DATA_COL_NAME + NUM_MONTHS - 1)}{IDM_ROW_CASH_TOTAL}"
    st.write_total_row(ws, formats, row=r, label="Total cash", formulas=total_formulas + [total_latest])
    total_row_1b = r + 1
    r += 2

    # Tie-out checks
    checks_end_one = st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {
                "name": "Sum of bank account totals equals Total cash row",
                "left": f"=SUM({col_letter(2 + NUM_MONTHS - 1)}{first_bank_row_1b}:{col_letter(2 + NUM_MONTHS - 1)}{last_bank_row_1b})",
                "right": f"={col_letter(2 + NUM_MONTHS - 1)}{total_row_1b}",
            },
            {
                "name": "Latest column ties to last month value on the same row",
                "left": f"={col_letter(2 + NUM_MONTHS)}{total_row_1b}",
                "right": f"={col_letter(2 + NUM_MONTHS - 1)}{total_row_1b}",
            },
        ],
    )

    # Chart: total cash trend
    chart_anchor = f"B{checks_end_one + 2}"
    cats_range = f"='Cash Position'!${col_letter(2)}${header_row + 1}:${col_letter(2 + NUM_MONTHS - 1)}${header_row + 1}"
    series = [{
        "name": "Total cash",
        "values": f"='Cash Position'!${col_letter(2)}${total_row_1b}:${col_letter(2 + NUM_MONTHS - 1)}${total_row_1b}",
        "color": b.CHART_PRIMARY,
    }]
    st.add_line_chart(wb, ws, title="Total cash trend", anchor_cell=chart_anchor,
                      series=series, cats_range=cats_range, width=720, height=280)

    sc.apply_page_setup(ws, sheet_title="Cash Position")
    return ws


def build_commentary(wb, formats, dat):
    ws = wb.add_worksheet("Commentary")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 14
    ws.set_column(0, 0, 2)
    ws.set_column(1, LAST_COL - 1, 10)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Free-form board notes",
        title="Commentary",
        last_col=LAST_COL + 1,
        explanation=(
            "Use these blocks to capture the narrative behind the numbers. The Executive "
            "Summary sits on the One Page sheet next to the KPIs; this tab holds the longer "
            "form notes for board members who want detail."
        ),
    )

    blocks = [
        ("Executive summary", "Three to five sentences summarising the month in your own words."),
        ("Trading update", "What is happening in the business this month? Sales pipeline, customer wins or losses, capacity utilisation."),
        ("Financial position", "Movement in cash, debtors, creditors. Any covenant or working-capital concerns."),
        ("Key risks", "What could go wrong? Material exposures, customer concentration, regulatory matters."),
        ("Decisions sought", "What does the board need to decide this month? Investment approvals, hires, capital raises."),
    ]

    r = 6
    for i, (label, prompt) in enumerate(blocks):
        st.write_section_header(ws, formats, row=r, kicker=f"Block {i + 1}", title=label)
        ws.set_row(r + 2, 80)
        ws.merge_range(r + 2, 1, r + 2, LAST_COL - 1, prompt, formats["input_text"])
        r += 4

    sc.apply_page_setup(ws, sheet_title="Commentary")
    return ws


def main():
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
    build_one_page(wb, formats, dat)
    build_pnl_detail(wb, formats, dat)
    build_cash_position(wb, formats, dat)
    build_commentary(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    build_internal_measures(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
