"""Build the Revenue and Margin Bridge workbook.

Data-driven decomposition of revenue and gross-margin change between the
prior quarter and the current quarter, broken down by your accounting software revenue account.
No user-set driver allocation; the data does the talking.

Output:
  C:\\dev\\lyros-workbooks\\library\\pnl
        \\lyros_lib_pnl_revenue_margin_bridge.xlsx
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\7010 - Revenue and Margin Bridge.xlsx")

WORKBOOK_ID = "7010"
WORKBOOK_TITLE = "Revenue and Margin Bridge"
WORKBOOK_KICKER = "What moved revenue and gross margin this quarter"
TARGET_USER = (
    "FP&A analyst, in-house Finance Controller, or fractional CFO explaining "
    "margin shifts to a CFO or board."
)
HOW_TO_USE = [
    "Open the Data sheet and paste your Profit and Loss by Month export.",
    "Make sure every revenue account is classified as Revenue in the Report line column.",
    "The Bridge sheet decomposes the revenue change by individual account; the Margin Walk shows GP change split between revenue effect and cost effect.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "Multi-channel sales business with Wholesale, Retail, Online channels"),
    ("REVENUE SCALE", "Circa $4M annual"),
    ("GROSS MARGIN", "44 per cent on average"),
    ("OUTPUT", "Bridge analysis suitable for a board pack appendix or CFO commentary"),
]

INPUTS_REQUIRED = [
    ("Profit and loss by month", "Data tab (paste P&L by Month export)"),
    ("Account to Report line mapping", "Data tab (dropdown in Report line column)"),
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

REVENUE_ACCOUNTS = [(c, n) for (c, n, _t, rl, _) in ACCOUNTS if rl == "Revenue"]


def col_letter(zero_based: int) -> str:
    s = ""
    n = zero_based
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def sumifs_month_rl(report_line: str, month_idx_0: int) -> str:
    sum_col = col_letter(DATA_COL_M1 + month_idx_0)
    crit_col = col_letter(DATA_COL_LINE)
    return (
        f"=SUMIFS('Data'!${sum_col}${DATA_ROW_FIRST_ACCT}:${sum_col}${DATA_ROW_LAST_ACCT},"
        f"'Data'!${crit_col}${DATA_ROW_FIRST_ACCT}:${crit_col}${DATA_ROW_LAST_ACCT},"
        f"\"{report_line}\")"
    )


def sumifs_quarter_rl(report_line: str, start_idx_0: int) -> str:
    """3-month sum for a Report line, starting at start_idx_0 (0-based month)."""
    parts = "+".join(sumifs_month_rl(report_line, start_idx_0 + i)[1:] for i in range(3))
    return f"=({parts})"


def sumifs_month_account(account_name: str, month_idx_0: int) -> str:
    """Sum a single account by name for one month."""
    sum_col = col_letter(DATA_COL_M1 + month_idx_0)
    name_col = col_letter(DATA_COL_NAME)
    return (
        f"=SUMIFS('Data'!${sum_col}${DATA_ROW_FIRST_ACCT}:${sum_col}${DATA_ROW_LAST_ACCT},"
        f"'Data'!${name_col}${DATA_ROW_FIRST_ACCT}:${name_col}${DATA_ROW_LAST_ACCT},"
        f"\"{account_name}\")"
    )


def sumifs_quarter_account(account_name: str, start_idx_0: int) -> str:
    parts = "+".join(sumifs_month_account(account_name, start_idx_0 + i)[1:] for i in range(3))
    return f"=({parts})"


def _synthetic_data() -> dict:
    rng_rev = d.make_rng("rmb-rev")
    rng_cogs = d.make_rng("rmb-cogs")
    rng_op = d.make_rng("rmb-op")
    rng_wg = d.make_rng("rmb-wg")

    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    revenue = d.revenue_series(NUM_MONTHS, rng_rev, base=320_000)
    cogs = d.cogs_series(revenue, rng_cogs, gm_pct=0.44)
    opex = d.opex_series(NUM_MONTHS, rng_op, base=82_000)
    wages = d.wages_series(revenue, rng_wg, ratio=0.21)
    da = [round(r * 0.04, 0) for r in revenue]

    totals = {"revenue": revenue, "cogs": cogs, "opex": opex, "wages": wages, "da": da}
    account_monthly: list[list[float]] = []
    for code, name, _t, _rl, (key, share) in ACCOUNTS:
        rng_a = d.make_rng(f"rmb-{code}")
        row = [round(totals[key][m] * (share + rng_a.uniform(-0.015, 0.015)), 0) for m in range(NUM_MONTHS)]
        account_monthly.append(row)

    return {"months": months, "account_monthly": account_monthly}


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
            "Paste your Profit and Loss by Month export into the table below. Each row "
            "is one account. The Bridge and Margin Walk sheets read this table at the account "
            "level for revenue (so individual sales accounts can be shown) and at the Report "
            "line level for costs."
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

    months = dat["months"]
    prior_start = NUM_MONTHS - 6
    curr_start = NUM_MONTHS - 3
    prior_label = f"{months[prior_start].strftime('%b %y')} – {months[prior_start + 2].strftime('%b %y')}"
    curr_label = f"{months[curr_start].strftime('%b %y')} – {months[curr_start + 2].strftime('%b %y')}"

    sc.write_hero_band(
        ws, formats,
        kicker=f"Current quarter: {curr_label}",
        title="Revenue and margin headline",
        last_col=LAST_COL + 1,
        explanation=(
            f"Compares the current quarter ({curr_label}) to the prior quarter ({prior_label}). "
            "Revenue, Cost of sales, Gross profit, and Gross margin are computed from the Data "
            "sheet using SUMIFS by Report line. GM stands for Gross Margin."
        ),
    )

    rev_prior = sumifs_quarter_rl("Revenue", prior_start)
    rev_curr = sumifs_quarter_rl("Revenue", curr_start)
    cogs_prior = sumifs_quarter_rl("Cost of sales", prior_start)
    cogs_curr = sumifs_quarter_rl("Cost of sales", curr_start)
    gp_prior = f"={rev_prior[1:]}-{cogs_prior[1:]}"
    gp_curr = f"={rev_curr[1:]}-{cogs_curr[1:]}"
    gm_prior = f"=IFERROR(({gp_prior[1:]})/({rev_prior[1:]}),0)"
    gm_curr = f"=IFERROR(({gp_curr[1:]})/({rev_curr[1:]}),0)"

    def change_text(curr: str, prior: str) -> str:
        return (
            f"=IF(({prior[1:]})=0,\"n/a\","
            f"\"vs prior Q  \"&TEXT((({curr[1:]})-({prior[1:]}))/ABS({prior[1:]}),\"+0.0%;-0.0%\"))"
        )

    st.kpi_card(ws, formats, anchor_row=6, anchor_col=1, width_cols=3,
                label="Revenue current Q", value_formula=rev_curr,
                change_formula=change_text(rev_curr, rev_prior), value_kind="aud")
    st.kpi_card(ws, formats, anchor_row=6, anchor_col=4, width_cols=3,
                label="Gross profit current Q", value_formula=gp_curr,
                change_formula=change_text(gp_curr, gp_prior), value_kind="aud")
    st.kpi_card(ws, formats, anchor_row=6, anchor_col=7, width_cols=3,
                label="Gross margin %", value_formula=gm_curr,
                change_formula=change_text(gm_curr, gm_prior), value_kind="pct")
    st.kpi_card(ws, formats, anchor_row=6, anchor_col=10, width_cols=3,
                label="Revenue prior Q", value_formula=rev_prior,
                change_formula="=\"baseline\"", value_kind="aud")

    # Q-over-Q comparison table
    section_row = 11
    st.write_section_header(ws, formats, row=section_row, kicker="Quarter over quarter", title="Headline figures")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row, headers=["Metric", "Prior Q", "Current Q", "Change $", "Change %"], right_align_from=2)

    rows = [
        ("Revenue", rev_prior, rev_curr, False),
        ("Cost of sales", cogs_prior, cogs_curr, False),
        ("Gross profit", gp_prior, gp_curr, True),
    ]
    r = header_row + 1
    for i, (label, p_f, c_f, bold) in enumerate(rows):
        zebra = i % 2 == 0
        label_fmt = formats["td_bold_left"] if bold else (formats["td_zebra"] if zebra else formats["td"])
        num_fmt = formats["td_bold_right"] if bold else (formats["td_right_zebra"] if zebra else formats["td_right"])
        pct_fmt = formats["td_bold_pct"] if bold else (formats["td_pct_zebra"] if zebra else formats["td_pct"])
        ws.write_string(r, 1, label, label_fmt)
        ws.write_formula(r, 2, p_f, num_fmt)
        ws.write_formula(r, 3, c_f, num_fmt)
        ws.write_formula(r, 4, f"=D{r + 1}-C{r + 1}", num_fmt)
        ws.write_formula(r, 5, f"=IFERROR((D{r + 1}-C{r + 1})/ABS(C{r + 1}),0)", pct_fmt)
        r += 1
    # GM% row separately (it's a ratio not a sum)
    zebra = (len(rows)) % 2 == 0
    label_fmt = formats["td_bold_left"]
    pct_fmt = formats["td_bold_pct"]
    ws.write_string(r, 1, "Gross margin %", label_fmt)
    ws.write_formula(r, 2, gm_prior, pct_fmt)
    ws.write_formula(r, 3, gm_curr, pct_fmt)
    ws.write_formula(r, 4, f"=D{r + 1}-C{r + 1}", pct_fmt)
    ws.write_formula(r, 5, f"=IFERROR((D{r + 1}-C{r + 1})/ABS(C{r + 1}),0)", pct_fmt)
    r += 1

    # Tie-out checks
    st.write_checks_block(
        ws, formats,
        row=r + 2,
        checks=[
            {
                "name": "Revenue current Q ties to Data Revenue for the last three months",
                "left": rev_curr,
                "right": f"=SUM('Data'!${col_letter(DATA_COL_M1 + NUM_MONTHS - 3)}${DATA_ROW_FIRST_ACCT}"
                          f":${col_letter(DATA_COL_M_LAST)}${DATA_ROW_LAST_ACCT})-"
                          f"SUMIFS('Data'!${col_letter(DATA_COL_TOTAL)}${DATA_ROW_FIRST_ACCT}:${col_letter(DATA_COL_TOTAL)}${DATA_ROW_LAST_ACCT},"
                          f"'Data'!${col_letter(DATA_COL_LINE)}${DATA_ROW_FIRST_ACCT}:${col_letter(DATA_COL_LINE)}${DATA_ROW_LAST_ACCT},"
                          f"\"<>Revenue\")*0+SUMIFS('Data'!${col_letter(DATA_COL_M1 + NUM_MONTHS - 3)}${DATA_ROW_FIRST_ACCT}:${col_letter(DATA_COL_M_LAST)}${DATA_ROW_LAST_ACCT},"
                          f"'Data'!${col_letter(DATA_COL_LINE)}${DATA_ROW_FIRST_ACCT}:${col_letter(DATA_COL_LINE)}${DATA_ROW_LAST_ACCT},\"Revenue\")",
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Headline")
    return ws


def build_bridge(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("Bridge")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 8
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, LAST_COL, 14)
    ws.set_column(LAST_COL + 1, LAST_COL + 1, 2)

    months = dat["months"]
    prior_start = NUM_MONTHS - 6
    curr_start = NUM_MONTHS - 3
    prior_label = f"{months[prior_start].strftime('%b %y')} – {months[prior_start + 2].strftime('%b %y')}"
    curr_label = f"{months[curr_start].strftime('%b %y')} – {months[curr_start + 2].strftime('%b %y')}"

    sc.write_hero_band(
        ws, formats,
        kicker="Revenue bridge by account",
        title="What moved revenue this quarter",
        last_col=LAST_COL + 2,
        explanation=(
            f"Decomposes the revenue change between the prior quarter ({prior_label}) and "
            f"the current quarter ({curr_label}) by individual your accounting software revenue account. Each "
            "account row shows its prior-quarter and current-quarter totals, the $ change, "
            "and its share of the total revenue movement."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet", title="Revenue change by account")

    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Revenue account", f"Prior Q ({prior_label})", f"Current Q ({curr_label})", "Change $", "Share of total change"],
                        right_align_from=2)

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (code, name) in enumerate(REVENUE_ACCOUNTS):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        pct_fmt = formats["td_pct_zebra"] if zebra else formats["td_pct"]
        prior_f = sumifs_quarter_account(name, prior_start)
        curr_f = sumifs_quarter_account(name, curr_start)
        ws.write_string(r, 1, f"{code} {name}", label_fmt)
        ws.write_formula(r, 2, prior_f, num_fmt)
        ws.write_formula(r, 3, curr_f, num_fmt)
        ws.write_formula(r, 4, f"=D{r + 1}-C{r + 1}", num_fmt)
        # Share of total change: this account's change / total revenue change
        total_change_f = (
            f"({sumifs_quarter_rl('Revenue', curr_start)[1:]})-({sumifs_quarter_rl('Revenue', prior_start)[1:]})"
        )
        ws.write_formula(r, 5, f"=IFERROR((D{r + 1}-C{r + 1})/({total_change_f}),0)", pct_fmt)
        r += 1

    # Total row
    last_account_row_1b = r
    ws.set_row(r, 24)
    ws.write_string(r, 1, "Total revenue change", formats["total_left"])
    ws.write_formula(r, 2, f"=SUM(C{first_data_row_1b}:C{last_account_row_1b})", formats["total_right"])
    ws.write_formula(r, 3, f"=SUM(D{first_data_row_1b}:D{last_account_row_1b})", formats["total_right"])
    ws.write_formula(r, 4, f"=SUM(E{first_data_row_1b}:E{last_account_row_1b})", formats["total_right"])
    ws.write_formula(r, 5, f"=SUM(F{first_data_row_1b}:F{last_account_row_1b})", formats["total_pct"])
    total_row_1b = r + 1

    # CF on the $ change column
    cf_range = f"E{first_data_row_1b}:E{last_account_row_1b}"
    st.add_three_color_scale(ws, cf_range, favourable_high=True)
    legend_row = r + 3
    st.write_cf_legend(ws, formats, row=legend_row, col=1, favourable_high=True, metric_label="Account $ change")

    # Tie-out checks
    checks_end = st.write_checks_block(
        ws, formats,
        row=legend_row + 3,
        checks=[
            {
                "name": "Sum of account changes equals total Revenue change for the quarter",
                "left": f"=$E${total_row_1b}",
                "right": f"=({sumifs_quarter_rl('Revenue', curr_start)[1:]})-({sumifs_quarter_rl('Revenue', prior_start)[1:]})",
            },
            {
                "name": "Account shares sum to 100 per cent",
                "left": f"=$F${total_row_1b}",
                "right": "=1",
                "is_pct": True,
            },
        ],
    )

    # Chart: column chart of account-level $ change
    chart_anchor = f"B{checks_end + 2}"
    cats_range = f"='Bridge'!$B${first_data_row_1b}:$B${last_account_row_1b}"
    series = [{
        "name": "Change $ by account",
        "values": f"='Bridge'!$E${first_data_row_1b}:$E${last_account_row_1b}",
        "color": b.CHART_PRIMARY,
    }]
    st.add_column_chart(wb, ws, title="Revenue change by account",
                       anchor_cell=chart_anchor, series=series, cats_range=cats_range,
                       width=720, height=300)

    sc.apply_page_setup(ws, sheet_title="Bridge")
    return ws


def build_margin_walk(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("Margin Walk")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 7
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 32)
    ws.set_column(2, LAST_COL, 14)
    ws.set_column(LAST_COL + 1, LAST_COL + 1, 2)

    months = dat["months"]
    prior_start = NUM_MONTHS - 6
    curr_start = NUM_MONTHS - 3
    prior_label = f"{months[prior_start].strftime('%b %y')} – {months[prior_start + 2].strftime('%b %y')}"
    curr_label = f"{months[curr_start].strftime('%b %y')} – {months[curr_start + 2].strftime('%b %y')}"

    sc.write_hero_band(
        ws, formats,
        kicker="Gross profit walk",
        title="Why gross profit moved",
        last_col=LAST_COL + 2,
        explanation=(
            f"Splits the change in Gross Profit between {prior_label} and {curr_label} into two "
            "data-driven components: the Revenue effect (revenue change at the prior gross margin) "
            "and the Cost effect (the residual change in Cost of sales). The two add to the total "
            "Gross Profit movement."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet", title="Gross profit walk")

    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row, headers=["Step", "Amount $"], right_align_from=2)

    rev_prior = sumifs_quarter_rl("Revenue", prior_start)
    rev_curr = sumifs_quarter_rl("Revenue", curr_start)
    cogs_prior = sumifs_quarter_rl("Cost of sales", prior_start)
    cogs_curr = sumifs_quarter_rl("Cost of sales", curr_start)
    gp_prior = f"={rev_prior[1:]}-{cogs_prior[1:]}"
    gp_curr = f"={rev_curr[1:]}-{cogs_curr[1:]}"
    gm_prior = f"=IFERROR(({gp_prior[1:]})/({rev_prior[1:]}),0)"

    # Revenue effect = (Rev_curr - Rev_prior) * GM_prior
    revenue_effect = f"=(({rev_curr[1:]})-({rev_prior[1:]}))*({gm_prior[1:]})"
    # Total GP change = GP_curr - GP_prior
    total_change = f"=({gp_curr[1:]})-({gp_prior[1:]})"
    # Cost effect = Total change - Revenue effect (residual)
    cost_effect = f"=({total_change[1:]})-({revenue_effect[1:]})"

    rows = [
        ("Starting gross profit (prior Q)", gp_prior, True),
        ("Revenue effect (volume + price)", revenue_effect, False),
        ("Cost effect (residual)", cost_effect, False),
        ("Ending gross profit (current Q)", gp_curr, True),
    ]

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (label, formula, bold) in enumerate(rows):
        zebra = i % 2 == 0
        label_fmt = formats["td_bold_left"] if bold else (formats["td_zebra"] if zebra else formats["td"])
        num_fmt = formats["td_bold_right"] if bold else (formats["td_right_zebra"] if zebra else formats["td_right"])
        ws.set_row(r, 22)
        ws.write_string(r, 1, label, label_fmt)
        ws.write_formula(r, 2, formula, num_fmt)
        r += 1
    last_row_1b = r

    # Tie-out: starting GP + revenue effect + cost effect = ending GP
    st.write_checks_block(
        ws, formats,
        row=r + 2,
        checks=[
            {
                "name": "Starting GP plus revenue effect plus cost effect equals ending GP",
                "left": f"=$C${first_data_row_1b}+$C${first_data_row_1b + 1}+$C${first_data_row_1b + 2}",
                "right": f"=$C${first_data_row_1b + 3}",
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Margin Walk")
    return ws


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
    build_bridge(wb, formats, dat)
    build_margin_walk(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
