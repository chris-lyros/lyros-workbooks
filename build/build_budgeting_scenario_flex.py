"""Budgeting Scenario Flex (4040).

Driver-based forecast with base, upside, and downside scenarios across
the next 12 months. The Drivers sheet holds the assumption set per
scenario; the three Forecast sheets recalculate from those drivers.
Comparison sheet shows the scenarios side-by-side.
"""

from __future__ import annotations
import sys
from datetime import date, timedelta
from pathlib import Path
import xlsxwriter

sys.path.insert(0, str(Path(__file__).parent))
from _shared import branding as b
from _shared import data as d
from _shared import scaffold as sc
from _shared import styles as st


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\4040 - Budgeting Scenario Flex.xlsx")

WORKBOOK_ID = "4040"
WORKBOOK_TITLE = "Budgeting Scenario Flex"
WORKBOOK_KICKER = "Driver-based forward 12 months across three scenarios"
TARGET_USER = (
    "Finance lead or fractional CFO modelling a hire, a property move, "
    "a price change, or a capital raise across three scenarios."
)
HOW_TO_USE = [
    "Set your starting position (last-month actuals) on the Drivers sheet.",
    "Adjust the assumption set per scenario (revenue growth, gross margin, opex growth, headcount).",
    "Each Forecast sheet recalculates; the Comparison sheet shows the three scenarios side-by-side.",
]

EXAMPLE_PROFILE = [
    ("BUSINESS PROFILE", "Trading SME with $5M annual revenue and 40 per cent gross margin"),
    ("DECISION CONTEXT", "Considering adding two delivery roles and a price increase next quarter"),
    ("SCENARIO RANGE", "Base: continue current trajectory. Upside: +10% revenue. Downside: -8% revenue."),
]

INPUTS_REQUIRED = [
    ("Starting revenue, COGS, opex, wages (last-month actuals)", "Drivers tab"),
    ("Per-scenario assumptions: revenue growth, gross margin, opex growth, wage growth", "Drivers tab"),
]

NUM_MONTHS = 12
SCENARIOS = ["Base", "Upside", "Downside"]

# Drivers sheet layout
DRV_ROW_FIRST = 9  # 1-based row where starting position begins


def col_letter(zero_based: int) -> str:
    s = ""
    n = zero_based
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def _months_forward(start: date, n: int) -> list[date]:
    """Return n month-end dates starting from the month after `start`."""
    out = []
    y, m = start.year, start.month
    for _ in range(n):
        m += 1
        if m == 13:
            m = 1
            y += 1
        out.append(d.month_end(date(y, m, 1)))
    return out


def build_drivers(wb, formats):
    ws = wb.add_worksheet("Drivers")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 5
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 32)
    ws.set_column(2, 4, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Assumptions",
        title="Starting position and scenario drivers",
        last_col=LAST_COL + 1,
        explanation=(
            "Top block: starting position from last month's actuals. Bottom block: per "
            "scenario monthly growth and margin assumptions. The Forecast sheets read "
            "from this sheet, so editing here cascades through the workbook."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Starting position", title="Last month's actuals")
    # Starting position rows
    starting = [
        ("Revenue", 420_000),
        ("Cost of sales", 252_000),
        ("Wages and on-costs", 92_000),
        ("Other operating expenses", 38_000),
    ]
    r = DRV_ROW_FIRST - 1  # convert to 0-based
    for i, (label, val) in enumerate(starting):
        zebra = i % 2 == 0
        ws.set_row(r, 22)
        ws.write_string(r, 1, label, formats["td_zebra"] if zebra else formats["td"])
        ws.write_number(r, 2, val, formats["input_value"])
        r += 1

    # Scenario assumption table
    r += 2
    st.write_section_header(ws, formats, row=r, kicker="Per scenario", title="Monthly drivers")
    header_row = r + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Driver", "Base", "Upside", "Downside"],
        right_align_from=2,
    )

    # Driver rows: monthly growth rates and ratios
    drivers = [
        ("Revenue monthly growth %", 0.010, 0.025, -0.008),
        ("Gross margin %", 0.40, 0.44, 0.36),
        ("Wages monthly growth %", 0.005, 0.010, 0.000),
        ("Other opex monthly growth %", 0.004, 0.008, 0.001),
    ]
    r = header_row + 1
    first_drv_row_1b = r + 1
    for i, (label, base, up, down) in enumerate(drivers):
        zebra = i % 2 == 0
        ws.set_row(r, 22)
        ws.write_string(r, 1, label, formats["td_zebra"] if zebra else formats["td"])
        ws.write_number(r, 2, base, formats["input_pct"])
        ws.write_number(r, 3, up, formats["input_pct"])
        ws.write_number(r, 4, down, formats["input_pct"])
        r += 1
    last_drv_row_1b = r

    sc.apply_page_setup(ws, sheet_title="Drivers")
    return first_drv_row_1b, last_drv_row_1b


def _scenario_forecast_sheet(wb, formats, scenario: str, scen_col: str):
    ws = wb.add_worksheet(scenario)
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 1 + NUM_MONTHS + 2
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, LAST_COL - 1, 12)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Scenario",
        title=f"{scenario} forecast (next 12 months)",
        last_col=LAST_COL + 1,
        explanation=(
            f"Twelve-month forecast under the {scenario} scenario, driven by the "
            f"assumptions on the Drivers sheet. Revenue compounds at the monthly "
            f"growth rate; gross margin holds at the scenario value; opex lines "
            f"compound at their own growth rate."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Twelve months", title="Forecast")
    header_row = section_row + 2

    month_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": b.WHITE, "bg_color": b.GREY_BLACK,
        "align": "right", "valign": "vcenter", "num_format": "mmm yy",
        "top": 2, "top_color": b.GREEN, "bottom": 2, "bottom_color": b.GREEN,
    })
    ws.set_row(header_row, 26)
    ws.write_string(header_row, 1, "Line", formats["th"])
    months = _months_forward(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    for i, m in enumerate(months):
        ws.write_datetime(header_row, 2 + i, m, month_fmt)
    ws.write_string(header_row, 2 + NUM_MONTHS, "Year 1 total", formats["th_right"])

    # Row plan
    # Revenue: month 1 = start * (1 + growth); month n = month n-1 * (1 + growth)
    # COGS: revenue * (1 - margin)
    # GP: revenue - COGS
    # Wages: starting wages compounded at wage growth
    # Other opex: starting opex compounded at opex growth
    # EBITDA: GP - wages - other opex
    r = header_row + 1
    rev_row_1b = r + 1
    cogs_row_1b = r + 2
    gp_row_1b = r + 3
    wages_row_1b = r + 4
    opex_row_1b = r + 5
    ebitda_row_1b = r + 6

    # Driver references (Drivers sheet)
    starting_rev_cell = f"Drivers!$C${DRV_ROW_FIRST}"
    starting_cogs_cell = f"Drivers!$C${DRV_ROW_FIRST + 1}"
    starting_wages_cell = f"Drivers!$C${DRV_ROW_FIRST + 2}"
    starting_opex_cell = f"Drivers!$C${DRV_ROW_FIRST + 3}"
    # Scenario drivers (Drivers sheet block 2)
    # The header_row of drivers block is at row 16 (1-based) approximately - resolve dynamically below
    # First driver row is DRV_ROW_FIRST + 4 (starting block) + 4 blank/section rows = approx row 17
    # We'll set fixed row 1-based positions matching build_drivers ordering:
    drv_rev_growth_row = DRV_ROW_FIRST + 4 + 4  # = 17
    drv_gm_row = drv_rev_growth_row + 1
    drv_wages_growth_row = drv_rev_growth_row + 2
    drv_opex_growth_row = drv_rev_growth_row + 3
    g_rev = f"Drivers!${scen_col}${drv_rev_growth_row}"
    g_gm = f"Drivers!${scen_col}${drv_gm_row}"
    g_wages = f"Drivers!${scen_col}${drv_wages_growth_row}"
    g_opex = f"Drivers!${scen_col}${drv_opex_growth_row}"

    # Labels
    ws.set_row(r, 22); ws.write_string(r, 1, "Revenue", formats["td_zebra"])
    ws.set_row(r + 1, 22); ws.write_string(r + 1, 1, "Cost of sales", formats["td"])
    ws.set_row(r + 2, 22); ws.write_string(r + 2, 1, "Gross profit", formats["td_zebra"])
    ws.set_row(r + 3, 22); ws.write_string(r + 3, 1, "Wages and on-costs", formats["td"])
    ws.set_row(r + 4, 22); ws.write_string(r + 4, 1, "Other operating expenses", formats["td_zebra"])
    ws.set_row(r + 5, 22); ws.write_string(r + 5, 1, "EBITDA", formats["td_bold_left"])

    for j in range(NUM_MONTHS):
        col = 2 + j
        col_l = col_letter(col)
        prev_col_l = col_letter(col - 1) if j > 0 else None

        # Revenue
        if j == 0:
            rev_f = f"={starting_rev_cell}*(1+{g_rev})"
        else:
            rev_f = f"={prev_col_l}{rev_row_1b}*(1+{g_rev})"
        ws.write_formula(rev_row_1b - 1, col, rev_f, formats["td_right_zebra"])

        # COGS = Revenue * (1 - margin)
        ws.write_formula(cogs_row_1b - 1, col,
                         f"={col_l}{rev_row_1b}*(1-{g_gm})", formats["td_right"])
        # GP
        ws.write_formula(gp_row_1b - 1, col,
                         f"={col_l}{rev_row_1b}-{col_l}{cogs_row_1b}",
                         formats["td_right_zebra"])
        # Wages
        if j == 0:
            w_f = f"={starting_wages_cell}*(1+{g_wages})"
        else:
            w_f = f"={prev_col_l}{wages_row_1b}*(1+{g_wages})"
        ws.write_formula(wages_row_1b - 1, col, w_f, formats["td_right"])
        # Opex
        if j == 0:
            o_f = f"={starting_opex_cell}*(1+{g_opex})"
        else:
            o_f = f"={prev_col_l}{opex_row_1b}*(1+{g_opex})"
        ws.write_formula(opex_row_1b - 1, col, o_f, formats["td_right_zebra"])
        # EBITDA
        ws.write_formula(ebitda_row_1b - 1, col,
                         f"={col_l}{gp_row_1b}-{col_l}{wages_row_1b}-{col_l}{opex_row_1b}",
                         formats["td_bold_right"])

    # Year-1 total column
    total_col = 2 + NUM_MONTHS
    last_month_col = col_letter(2 + NUM_MONTHS - 1)
    first_month_col = col_letter(2)
    for row_1b, fmt_key in [
        (rev_row_1b, "total_right"), (cogs_row_1b, "total_right"),
        (gp_row_1b, "total_right"), (wages_row_1b, "total_right"),
        (opex_row_1b, "total_right"), (ebitda_row_1b, "total_right"),
    ]:
        ws.write_formula(row_1b - 1, total_col,
                         f"=SUM({first_month_col}{row_1b}:{last_month_col}{row_1b})",
                         formats[fmt_key])

    # Tie-out check
    r_check = ebitda_row_1b + 2
    st.write_checks_block(
        ws, formats, row=r_check,
        checks=[
            {"name": "EBITDA equals Gross profit less wages less other opex",
             "left": f"={first_month_col}{ebitda_row_1b}",
             "right": f"={first_month_col}{gp_row_1b}-{first_month_col}{wages_row_1b}-{first_month_col}{opex_row_1b}"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title=scenario)


def build_comparison(wb, formats):
    ws = wb.add_worksheet("Comparison")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 5
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, 4, 16)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Side-by-side",
        title="Scenarios compared",
        last_col=LAST_COL + 1,
        explanation=(
            "Year-1 totals per scenario for each line of the forecast. Use this view to "
            "frame the decision: what would happen to EBITDA across plausible futures?"
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Year 1 totals", title="Three scenarios compared")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Line", "Base", "Upside", "Downside"],
                        right_align_from=2)

    # Forecast sheets' Year-1 total column is at column N (2 + 12 = 14 → col index 14 → letter N)
    total_col_letter = col_letter(2 + NUM_MONTHS)

    lines = [
        ("Revenue", 10),
        ("Cost of sales", 11),
        ("Gross profit", 12),
        ("Wages and on-costs", 13),
        ("Other operating expenses", 14),
        ("EBITDA", 15),
    ]
    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (label, sheet_row_1b) in enumerate(lines):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, label, label_fmt)
        for j, scen in enumerate(SCENARIOS):
            ws.write_formula(r, 2 + j,
                             f"='{scen}'!{total_col_letter}{sheet_row_1b}",
                             num_fmt)
        r += 1
    last_data_row_1b = r

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Upside EBITDA is greater than or equal to Base EBITDA",
             "left": "=D" + str(first_data_row_1b + 5),
             "right": "=C" + str(first_data_row_1b + 5)},
            {"name": "Downside EBITDA is less than or equal to Base EBITDA",
             "left": "=C" + str(first_data_row_1b + 5),
             "right": "=E" + str(first_data_row_1b + 5)},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Comparison")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE,
                       workbook_kicker=WORKBOOK_KICKER,
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_drivers(wb, formats)
    _scenario_forecast_sheet(wb, formats, "Base", "C")
    _scenario_forecast_sheet(wb, formats, "Upside", "D")
    _scenario_forecast_sheet(wb, formats, "Downside", "E")
    build_comparison(wb, formats)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
