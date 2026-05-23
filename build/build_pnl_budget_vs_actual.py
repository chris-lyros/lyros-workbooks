"""Build the Budget vs Actual workbook.

Architecture: Data sheet holds two tables - Actuals (P&L by Month export)
and Budget (same account-level shape). Analytical sheets aggregate from both
via SUMIFS keyed on Report line and pull Actual / Budget side by side.

Output:
  C:\\dev\\lyros-workbooks\\library\\pnl
        \\lyros_lib_pnl_budget_vs_actual.xlsx
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\4010 - Budget vs Actual.xlsx")

WORKBOOK_ID = "4010"
WORKBOOK_TITLE = "Budget vs Actual"
WORKBOOK_KICKER = "Monthly variance with year-to-date bridge"
TARGET_USER = (
    "In-house Finance Controller or CFO tracking budget performance against "
    "a board-approved plan."
)
HOW_TO_USE = [
    "Open the Data sheet and paste your your accounting software Profit and Loss by Month into the Actuals table.",
    "Paste your annual budget by month into the Budget table directly below, using the same account codes.",
    "The BvA Monthly sheet shows variance for every line item; the Headline sheet calls out the top movers.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "Wholesale and retail (Wholesale 50%, Retail 35%, Online 12%, Other 3%)"),
    ("REVENUE SCALE", "Circa $4M annual"),
    ("BUDGET SHAPE", "Annual plan spread monthly with seasonal weighting"),
    ("REPORTING CADENCE", "Monthly close, monthly board review"),
]

INPUTS_REQUIRED = [
    ("Actuals: P&L by Month", "Data tab → Actuals table (paste from your accounting software)"),
    ("Budget: P&L by Month", "Data tab → Budget table (paste from your annual plan)"),
    ("Account to Report line mapping", "Data tab → Report line dropdown column"),
]


# ── Layout ─────────────────────────────────────────────────────────────────

NUM_MONTHS = 12

DATA_COL_CODE = 1
DATA_COL_NAME = 2
DATA_COL_TYPE = 3
DATA_COL_LINE = 4
DATA_COL_M1 = 5
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1  # Q
DATA_COL_TOTAL = DATA_COL_M_LAST + 1  # R

# Actuals table rows (1-based)
ACT_ROW_HEADER = 8
ACT_ROW_FIRST = 9
ACT_ROW_LAST = ACT_ROW_FIRST + 16

# Budget table rows (placed below actuals with a gap)
BUD_SECTION_ROW = ACT_ROW_LAST + 4  # 0-based section_row for Budget
BUD_ROW_HEADER = ACT_ROW_LAST + 6
BUD_ROW_FIRST = BUD_ROW_HEADER + 1
BUD_ROW_LAST = BUD_ROW_FIRST + 16

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


def col_letter(zero_based: int) -> str:
    s = ""
    n = zero_based
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


# SUMIFS helpers, parametrised by which table (Actuals or Budget) to read from.

def sumifs_month(report_line: str, month_idx_0: int, *, table: str = "actuals") -> str:
    sum_col = col_letter(DATA_COL_M1 + month_idx_0)
    crit_col = col_letter(DATA_COL_LINE)
    if table == "actuals":
        first, last = ACT_ROW_FIRST, ACT_ROW_LAST
    else:
        first, last = BUD_ROW_FIRST, BUD_ROW_LAST
    return (
        f"=SUMIFS('Data'!${sum_col}${first}:${sum_col}${last},"
        f"'Data'!${crit_col}${first}:${crit_col}${last},"
        f"\"{report_line}\")"
    )


def sumifs_fy(report_line: str, *, table: str = "actuals") -> str:
    sum_col = col_letter(DATA_COL_TOTAL)
    crit_col = col_letter(DATA_COL_LINE)
    if table == "actuals":
        first, last = ACT_ROW_FIRST, ACT_ROW_LAST
    else:
        first, last = BUD_ROW_FIRST, BUD_ROW_LAST
    return (
        f"=SUMIFS('Data'!${sum_col}${first}:${sum_col}${last},"
        f"'Data'!${crit_col}${first}:${crit_col}${last},"
        f"\"{report_line}\")"
    )


def sumifs_ytd(report_line: str, months_count: int, *, table: str = "actuals") -> str:
    """Sum the first N months of a report line (YTD)."""
    if table == "actuals":
        first, last = ACT_ROW_FIRST, ACT_ROW_LAST
    else:
        first, last = BUD_ROW_FIRST, BUD_ROW_LAST
    crit_col = col_letter(DATA_COL_LINE)
    parts = []
    for i in range(months_count):
        col = col_letter(DATA_COL_M1 + i)
        parts.append(
            f"SUMIFS('Data'!${col}${first}:${col}${last},"
            f"'Data'!${crit_col}${first}:${crit_col}${last},\"{report_line}\")"
        )
    return "=" + "+".join(parts)


# ── Synthetic data ─────────────────────────────────────────────────────────

def _synthetic_data() -> dict:
    rng_rev = d.make_rng("bva-rev")
    rng_cogs = d.make_rng("bva-cogs")
    rng_op = d.make_rng("bva-op")
    rng_wg = d.make_rng("bva-wg")

    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    revenue = d.revenue_series(NUM_MONTHS, rng_rev, base=320_000)
    cogs = d.cogs_series(revenue, rng_cogs, gm_pct=0.44)
    opex = d.opex_series(NUM_MONTHS, rng_op, base=82_000)
    wages = d.wages_series(revenue, rng_wg, ratio=0.21)
    da = [round(r * 0.04, 0) for r in revenue]

    # Budget: smoothed version of revenue (less seasonality, modest growth path)
    rng_b = d.make_rng("bva-budget")
    budget_rev_base = sum(revenue) / NUM_MONTHS  # average
    bud_rev = [round(budget_rev_base * (1.0 + 0.01 * (i - 5)), 0) for i in range(NUM_MONTHS)]
    bud_cogs = [round(r * (1 - 0.44), 0) for r in bud_rev]
    bud_opex = [round(82_000 * (1.0 + 0.005 * i), 0) for i in range(NUM_MONTHS)]
    bud_wages = [round(r * 0.20, 0) for r in bud_rev]
    bud_da = [round(r * 0.04, 0) for r in bud_rev]

    def allocate(totals: list[float], account_idx_filter, rng_salt) -> list[list[float]]:
        rows: list[list[float]] = []
        for i, (code, name, _type, _rl, (series_key, share)) in enumerate(ACCOUNTS):
            if account_idx_filter is not None and i not in account_idx_filter:
                rows.append([0] * NUM_MONTHS)
                continue
            rng_a = d.make_rng(f"{rng_salt}-{code}")
            series = totals.get(series_key, [0] * NUM_MONTHS)
            row = [round(series[m] * (share + rng_a.uniform(-0.015, 0.015)), 0) for m in range(NUM_MONTHS)]
            rows.append(row)
        return rows

    actuals_totals = {"revenue": revenue, "cogs": cogs, "opex": opex, "wages": wages, "da": da}
    budget_totals = {"revenue": bud_rev, "cogs": bud_cogs, "opex": bud_opex, "wages": bud_wages, "da": bud_da}

    actuals_monthly = allocate(actuals_totals, None, "bva-act")
    budget_monthly = allocate(budget_totals, None, "bva-bud")

    return {
        "months": months,
        "actuals_monthly": actuals_monthly,
        "budget_monthly": budget_monthly,
    }


# ── Data sheet (Actuals + Budget) ──────────────────────────────────────────

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
        title="Drop your actuals and budget here",
        last_col=LAST_COL + 1,
        explanation=(
            "Paste your Profit and Loss by Month into the Actuals table below. "
            "Paste your annual budget into the Budget table further down, using the same "
            "account codes and Report line classification. Other sheets calculate variance "
            "from these two tables via SUMIFS by Report line."
        ),
    )

    month_headers = [m.strftime("%b %y") for m in dat["months"]]

    def write_table(*, section_row_0: int, header_row_1: int, first_row_1: int, last_row_1: int,
                    section_kicker: str, section_title: str, table_data: list[list[float]]) -> None:
        st.write_section_header(ws, formats, row=section_row_0, kicker=section_kicker, title=section_title)
        headers = ["Code", "Account name", "Account type", "Report line"] + month_headers + ["FY total"]
        st.write_header_row(ws, formats, row=header_row_1 - 1, headers=headers, start_col=DATA_COL_CODE, right_align_from=5)

        for i, (code, name, acct_type, report_line, _) in enumerate(ACCOUNTS):
            row_0 = first_row_1 - 1 + i
            zebra = i % 2 == 0
            text_fmt = formats["td_zebra"] if zebra else formats["td"]
            ws.set_row(row_0, 20)
            ws.write_string(row_0, DATA_COL_CODE, code, text_fmt)
            ws.write_string(row_0, DATA_COL_NAME, name, text_fmt)
            ws.write_string(row_0, DATA_COL_TYPE, acct_type, text_fmt)
            ws.write_string(row_0, DATA_COL_LINE, report_line, formats["input_text"])
            for m, v in enumerate(table_data[i]):
                ws.write_number(row_0, DATA_COL_M1 + m, v, formats["input_value"])
            first_m = col_letter(DATA_COL_M1)
            last_m = col_letter(DATA_COL_M_LAST)
            ws.write_formula(row_0, DATA_COL_TOTAL,
                             f"=SUM(${first_m}${first_row_1 + i}:${last_m}${first_row_1 + i})",
                             formats["td_bold_right"])

        ws.data_validation(first_row_1 - 1, DATA_COL_LINE, last_row_1 - 1, DATA_COL_LINE,
                           {"validate": "list", "source": REPORT_LINES})

    # Actuals table
    write_table(
        section_row_0=6,
        header_row_1=ACT_ROW_HEADER,
        first_row_1=ACT_ROW_FIRST,
        last_row_1=ACT_ROW_LAST,
        section_kicker="Step 1   Paste your P&L by Month (Actuals)",
        section_title="Actuals",
        table_data=dat["actuals_monthly"],
    )

    # Budget table
    write_table(
        section_row_0=BUD_SECTION_ROW,
        header_row_1=BUD_ROW_HEADER,
        first_row_1=BUD_ROW_FIRST,
        last_row_1=BUD_ROW_LAST,
        section_kicker="Step 2   Paste your annual budget by month",
        section_title="Budget",
        table_data=dat["budget_monthly"],
    )

    sc.apply_page_setup(ws, sheet_title="Data")
    return ws


# ── Headline ────────────────────────────────────────────────────────────────

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
        title="Budget vs Actual summary",
        last_col=LAST_COL + 1,
        explanation=(
            "Compares the most recent month's actual figures to budget for the headline "
            "lines (Revenue, Gross Profit, EBITDA, Net Profit). Year-to-date Actual vs "
            "Budget shows the cumulative position. Acronyms: Gross Profit (GP), Earnings "
            "Before Interest, Tax, Depreciation and Amortisation (EBITDA), Net Profit "
            "After Tax (NPAT)."
        ),
    )

    last_idx = NUM_MONTHS - 1

    def kpi_av(label: str, anchor_col: int, key: str) -> None:
        # Actual value + budget + variance % packed into a card
        if key == "Revenue":
            act = sumifs_month("Revenue", last_idx, table="actuals")
            bud = sumifs_month("Revenue", last_idx, table="budget")
        elif key == "GP":
            act = f"={sumifs_month('Revenue', last_idx, table='actuals')[1:]}-{sumifs_month('Cost of sales', last_idx, table='actuals')[1:]}"
            bud = f"={sumifs_month('Revenue', last_idx, table='budget')[1:]}-{sumifs_month('Cost of sales', last_idx, table='budget')[1:]}"
        elif key == "EBITDA":
            act = (
                f"={sumifs_month('Revenue', last_idx, table='actuals')[1:]}-"
                f"{sumifs_month('Cost of sales', last_idx, table='actuals')[1:]}-"
                f"{sumifs_month('Wages', last_idx, table='actuals')[1:]}-"
                f"{sumifs_month('Other opex', last_idx, table='actuals')[1:]}"
            )
            bud = (
                f"={sumifs_month('Revenue', last_idx, table='budget')[1:]}-"
                f"{sumifs_month('Cost of sales', last_idx, table='budget')[1:]}-"
                f"{sumifs_month('Wages', last_idx, table='budget')[1:]}-"
                f"{sumifs_month('Other opex', last_idx, table='budget')[1:]}"
            )
        elif key == "NPAT":
            act = (
                f"={sumifs_month('Revenue', last_idx, table='actuals')[1:]}-"
                f"{sumifs_month('Cost of sales', last_idx, table='actuals')[1:]}-"
                f"{sumifs_month('Wages', last_idx, table='actuals')[1:]}-"
                f"{sumifs_month('Other opex', last_idx, table='actuals')[1:]}-"
                f"{sumifs_month('D&A', last_idx, table='actuals')[1:]}"
            )
            bud = (
                f"={sumifs_month('Revenue', last_idx, table='budget')[1:]}-"
                f"{sumifs_month('Cost of sales', last_idx, table='budget')[1:]}-"
                f"{sumifs_month('Wages', last_idx, table='budget')[1:]}-"
                f"{sumifs_month('Other opex', last_idx, table='budget')[1:]}-"
                f"{sumifs_month('D&A', last_idx, table='budget')[1:]}"
            )

        change_f = (
            f"=IF(({bud[1:]})=0,\"n/a\","
            f"\"vs budget  \"&TEXT((({act[1:]})-({bud[1:]}))/ABS({bud[1:]}),\"+0.0%;-0.0%\"))"
        )
        st.kpi_card(
            ws, formats,
            anchor_row=6, anchor_col=anchor_col, width_cols=3,
            label=label, value_formula=act, change_formula=change_f,
            value_kind="aud",
        )

    kpi_av("Revenue", 1, "Revenue")
    kpi_av("Gross profit", 4, "GP")
    kpi_av("EBITDA", 7, "EBITDA")
    kpi_av("Net profit", 10, "NPAT")

    # YTD Actual vs Budget table
    section_row = 11
    st.write_section_header(ws, formats, row=section_row, kicker="Year-to-date snapshot", title="Actual vs Budget, year-to-date")

    header_row_idx = section_row + 2
    ytd_months_count = NUM_MONTHS  # treat the full 12 months as YTD for sample
    headers = ["Line", "YTD Actual", "YTD Budget", "Variance $", "Variance %"]
    st.write_header_row(ws, formats, row=header_row_idx, headers=headers, right_align_from=2)

    lines = [
        ("Revenue", "Revenue"),
        ("Cost of sales", "Cost of sales"),
        ("Gross profit", None),
        ("Wages", "Wages"),
        ("Other opex", "Other opex"),
        ("EBITDA", None),
        ("D&A", "D&A"),
        ("Net profit", None),
    ]

    r = header_row_idx + 1
    first_data_row_1b = r + 1
    for i, (label, rl) in enumerate(lines):
        ws.set_row(r, 22)
        zebra = i % 2 == 0
        bold = label in ("Gross profit", "EBITDA", "Net profit")
        label_fmt = formats["td_bold_left"] if bold else (formats["td_zebra"] if zebra else formats["td"])
        num_fmt = formats["td_bold_right"] if bold else (formats["td_right_zebra"] if zebra else formats["td_right"])
        pct_fmt = formats["td_bold_pct"] if bold else (formats["td_pct_zebra"] if zebra else formats["td_pct"])
        ws.write_string(r, 1, label, label_fmt)

        if rl is not None:
            act_ytd = sumifs_ytd(rl, ytd_months_count, table="actuals")
            bud_ytd = sumifs_ytd(rl, ytd_months_count, table="budget")
        elif label == "Gross profit":
            act_ytd = f"={sumifs_ytd('Revenue', ytd_months_count, table='actuals')[1:]}-{sumifs_ytd('Cost of sales', ytd_months_count, table='actuals')[1:]}"
            bud_ytd = f"={sumifs_ytd('Revenue', ytd_months_count, table='budget')[1:]}-{sumifs_ytd('Cost of sales', ytd_months_count, table='budget')[1:]}"
        elif label == "EBITDA":
            act_ytd = (
                f"={sumifs_ytd('Revenue', ytd_months_count, table='actuals')[1:]}-"
                f"{sumifs_ytd('Cost of sales', ytd_months_count, table='actuals')[1:]}-"
                f"{sumifs_ytd('Wages', ytd_months_count, table='actuals')[1:]}-"
                f"{sumifs_ytd('Other opex', ytd_months_count, table='actuals')[1:]}"
            )
            bud_ytd = (
                f"={sumifs_ytd('Revenue', ytd_months_count, table='budget')[1:]}-"
                f"{sumifs_ytd('Cost of sales', ytd_months_count, table='budget')[1:]}-"
                f"{sumifs_ytd('Wages', ytd_months_count, table='budget')[1:]}-"
                f"{sumifs_ytd('Other opex', ytd_months_count, table='budget')[1:]}"
            )
        else:  # Net profit
            act_ytd = (
                f"={sumifs_ytd('Revenue', ytd_months_count, table='actuals')[1:]}-"
                f"{sumifs_ytd('Cost of sales', ytd_months_count, table='actuals')[1:]}-"
                f"{sumifs_ytd('Wages', ytd_months_count, table='actuals')[1:]}-"
                f"{sumifs_ytd('Other opex', ytd_months_count, table='actuals')[1:]}-"
                f"{sumifs_ytd('D&A', ytd_months_count, table='actuals')[1:]}"
            )
            bud_ytd = (
                f"={sumifs_ytd('Revenue', ytd_months_count, table='budget')[1:]}-"
                f"{sumifs_ytd('Cost of sales', ytd_months_count, table='budget')[1:]}-"
                f"{sumifs_ytd('Wages', ytd_months_count, table='budget')[1:]}-"
                f"{sumifs_ytd('Other opex', ytd_months_count, table='budget')[1:]}-"
                f"{sumifs_ytd('D&A', ytd_months_count, table='budget')[1:]}"
            )

        ws.write_formula(r, 2, act_ytd, num_fmt)
        ws.write_formula(r, 3, bud_ytd, num_fmt)
        # Variance $ = Actual - Budget
        ws.write_formula(r, 4, f"=C{r + 1}-D{r + 1}", num_fmt)
        # Variance % = (Actual - Budget) / ABS(Budget)
        ws.write_formula(r, 5, f"=IFERROR((C{r + 1}-D{r + 1})/ABS(D{r + 1}),0)", pct_fmt)
        r += 1
    last_data_row_1b = r

    # Conditional formatting on Variance % column: favourable for revenue/GP/EBITDA/NPAT,
    # unfavourable for costs. Apply per cell.
    fav_cells = [f"F{first_data_row_1b}", f"F{first_data_row_1b + 2}", f"F{first_data_row_1b + 5}", f"F{first_data_row_1b + 7}"]
    unfav_cells = [f"F{first_data_row_1b + 1}", f"F{first_data_row_1b + 3}", f"F{first_data_row_1b + 4}", f"F{first_data_row_1b + 6}"]
    st.add_three_color_scale(ws, ",".join(fav_cells), favourable_high=True)
    st.add_three_color_scale(ws, ",".join(unfav_cells), favourable_high=False)

    legend_row = r + 1
    st.write_cf_legend(ws, formats, row=legend_row, col=1, favourable_high=True,
                       metric_label="Variance % on Revenue, GP, EBITDA, Net profit")
    st.write_cf_legend(ws, formats, row=legend_row + 1, col=1, favourable_high=False,
                       metric_label="Variance % on costs (lower is better)")

    # Tie-out checks
    st.write_checks_block(
        ws, formats,
        row=legend_row + 3,
        checks=[
            {
                "name": "YTD Actual Revenue ties to Data Actuals Revenue total",
                "left": f"=$C${first_data_row_1b}",
                "right": sumifs_fy("Revenue", table="actuals"),
            },
            {
                "name": "YTD Budget Revenue ties to Data Budget Revenue total",
                "left": f"=$D${first_data_row_1b}",
                "right": sumifs_fy("Revenue", table="budget"),
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Headline")
    return ws


# ── BvA Monthly ─────────────────────────────────────────────────────────────

def build_bva_monthly(wb: xlsxwriter.Workbook, formats: dict, dat: dict):
    ws = wb.add_worksheet("BvA Monthly")
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
        title="Budget vs Actual by month",
        last_col=LAST_COL + 1,
        explanation=(
            "Three blocks stacked: Actuals (drawn from the Data Actuals table), Budget "
            "(drawn from the Data Budget table), and Variance (Actual minus Budget). Each "
            "block reads from the same Report line classification."
        ),
    )

    months = dat["months"]
    month_headers = [m.strftime("%b %y") for m in months]

    section_row = 6

    def write_block(*, start_row: int, block_label: str, table: str | None) -> int:
        """Write a 5-row block (Revenue, COGS, GP, Wages, Opex, EBITDA, D&A, NPAT) labelled.

        table = 'actuals' or 'budget' for direct SUMIFS; None for the variance block
        which references the actual and budget blocks.
        """
        st.write_section_header(ws, formats, row=start_row, kicker=block_label, title=block_label)
        header_row = start_row + 2
        st.write_header_row(ws, formats, row=header_row,
                            headers=["Line item"] + month_headers + ["FY Total"],
                            right_align_from=2)

        r = header_row + 1
        return r

    # Actuals block
    act_start = section_row
    r_act = write_block(start_row=act_start, block_label="Actuals", table="actuals")

    def write_pnl_lines(start_row_0: int, table: str) -> int:
        rev_m = [sumifs_month("Revenue", m, table=table) for m in range(NUM_MONTHS)]
        cogs_m = [sumifs_month("Cost of sales", m, table=table) for m in range(NUM_MONTHS)]
        wages_m = [sumifs_month("Wages", m, table=table) for m in range(NUM_MONTHS)]
        opex_m = [sumifs_month("Other opex", m, table=table) for m in range(NUM_MONTHS)]
        da_m = [sumifs_month("D&A", m, table=table) for m in range(NUM_MONTHS)]
        gp_m = [f"={rev_m[m][1:]}-{cogs_m[m][1:]}" for m in range(NUM_MONTHS)]
        eb_m = [f"={gp_m[m][1:]}-{wages_m[m][1:]}-{opex_m[m][1:]}" for m in range(NUM_MONTHS)]
        np_m = [f"={eb_m[m][1:]}-{da_m[m][1:]}" for m in range(NUM_MONTHS)]

        rev_fy = sumifs_fy("Revenue", table=table)
        cogs_fy = sumifs_fy("Cost of sales", table=table)
        wages_fy = sumifs_fy("Wages", table=table)
        opex_fy = sumifs_fy("Other opex", table=table)
        da_fy = sumifs_fy("D&A", table=table)
        gp_fy = f"={rev_fy[1:]}-{cogs_fy[1:]}"
        eb_fy = f"={gp_fy[1:]}-{wages_fy[1:]}-{opex_fy[1:]}"
        np_fy = f"={eb_fy[1:]}-{da_fy[1:]}"

        r = start_row_0
        st.write_data_row(ws, formats, row=r, label="Revenue", formulas=rev_m + [rev_fy], zebra=False, cell_format="td_right"); r += 1
        st.write_data_row(ws, formats, row=r, label="Cost of sales", formulas=[f"=-{x[1:]}" for x in cogs_m] + [f"=-{cogs_fy[1:]}"], zebra=True, cell_format="td_right"); r += 1
        st.write_data_row(ws, formats, row=r, label="Gross profit", formulas=gp_m + [gp_fy], zebra=False, bold=True, cell_format="td_right"); r += 1
        st.write_data_row(ws, formats, row=r, label="Wages and salaries", formulas=[f"=-{x[1:]}" for x in wages_m] + [f"=-{wages_fy[1:]}"], zebra=True, cell_format="td_right"); r += 1
        st.write_data_row(ws, formats, row=r, label="Other opex", formulas=[f"=-{x[1:]}" for x in opex_m] + [f"=-{opex_fy[1:]}"], zebra=False, cell_format="td_right"); r += 1
        st.write_data_row(ws, formats, row=r, label="EBITDA", formulas=eb_m + [eb_fy], zebra=True, bold=True, cell_format="td_right"); r += 1
        st.write_data_row(ws, formats, row=r, label="D&A", formulas=[f"=-{x[1:]}" for x in da_m] + [f"=-{da_fy[1:]}"], zebra=False, cell_format="td_right"); r += 1
        st.write_data_row(ws, formats, row=r, label="Net profit", formulas=np_m + [np_fy], zebra=True, bold=True, cell_format="td_right"); r += 1
        return r

    act_end = write_pnl_lines(r_act, "actuals")

    # Budget block
    bud_start = act_end + 1
    r_bud = write_block(start_row=bud_start, block_label="Budget", table="budget")
    bud_end = write_pnl_lines(r_bud, "budget")
    bud_first_data_row_1b = r_bud + 1

    # Variance block (Actual - Budget) referencing the two prior blocks
    var_start = bud_end + 1
    st.write_section_header(ws, formats, row=var_start, kicker="Variance", title="Actual minus Budget")
    var_header = var_start + 2
    st.write_header_row(ws, formats, row=var_header, headers=["Line item"] + month_headers + ["FY Total"], right_align_from=2)

    # Reference rows: actuals start at r_act + 1 (1-based), budget at r_bud + 1 (1-based).
    # Each line is offset by line index 0..7.
    line_names = ["Revenue", "Cost of sales", "Gross profit", "Wages and salaries",
                  "Other opex", "EBITDA", "D&A", "Net profit"]
    r_var = var_header + 1
    for i, name in enumerate(line_names):
        zebra = i % 2 == 0
        bold = name in ("Gross profit", "EBITDA", "Net profit")
        act_row_1b = r_act + 1 + i
        bud_row_1b = r_bud + 1 + i
        formulas = [
            f"={col_letter(2 + m)}{act_row_1b}-{col_letter(2 + m)}{bud_row_1b}"
            for m in range(NUM_MONTHS)
        ]
        total_f = f"={col_letter(2 + NUM_MONTHS)}{act_row_1b}-{col_letter(2 + NUM_MONTHS)}{bud_row_1b}"
        st.write_data_row(ws, formats, row=r_var, label=name, formulas=formulas + [total_f],
                          zebra=zebra, bold=bold, cell_format="td_right")
        r_var += 1

    # Tie-out checks
    st.write_checks_block(
        ws, formats,
        row=r_var + 2,
        checks=[
            {
                "name": "Actuals Revenue FY ties to Data Actuals Revenue total",
                "left": f"={col_letter(2 + NUM_MONTHS)}{r_act + 1}",
                "right": sumifs_fy("Revenue", table="actuals"),
            },
            {
                "name": "Budget Revenue FY ties to Data Budget Revenue total",
                "left": f"={col_letter(2 + NUM_MONTHS)}{r_bud + 1}",
                "right": sumifs_fy("Revenue", table="budget"),
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="BvA Monthly")
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
    build_bva_monthly(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)

    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
