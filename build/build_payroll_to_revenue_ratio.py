"""Payroll to Revenue Ratio (8100).

Payroll cost as a share of revenue over time, broken down by department.
A simple leading indicator of operating health: if the ratio creeps up
while revenue is flat, you are heading into trouble.
"""

from __future__ import annotations
import sys
from pathlib import Path
import xlsxwriter

sys.path.insert(0, str(Path(__file__).parent))
from _shared import branding as b
from _shared import data as d
from _shared import scaffold as sc
from _shared import styles as st


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\8100 - Payroll to Revenue Ratio.xlsx")

WORKBOOK_ID = "8100"
WORKBOOK_TITLE = "Payroll to Revenue Ratio"
WORKBOOK_KICKER = "Operating-health ratio over time, by department"
TARGET_USER = "Finance lead or fractional CFO monitoring payroll cost intensity month over month."
HOW_TO_USE = [
    "Open the Data sheet and paste revenue and total payroll cost by month (12 months).",
    "Update the Internal Data Measures sheet with the department wage allocation shares.",
    "The Trend and By Department sheets recalculate the ratio with sparklines and direction flags.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "Professional services SME with mixed delivery and admin headcount"),
    ("PAYROLL RATIO", "Circa 38 per cent of revenue, target range 32 to 40 per cent"),
    ("DEPT MIX", "Delivery 55%, Sales 18%, Admin 15%, Finance 8%, Executive 4%"),
]

INPUTS_REQUIRED = [
    ("Revenue and total payroll by month (12 months)", "Data tab"),
    ("Department wage allocation shares (sum to 100%)", "Internal Data Measures tab"),
]

NUM_MONTHS = 12
DATA_COL_LABEL = 1
DATA_COL_M1 = 2
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1
DATA_ROW_HEADER = 9
DATA_ROW_REV = DATA_ROW_HEADER + 1
DATA_ROW_PAYROLL = DATA_ROW_HEADER + 2


def col_letter(zero_based: int) -> str:
    s = ""
    n = zero_based
    while True:
        s = chr(ord("A") + (n % 26)) + s
        n = n // 26 - 1
        if n < 0:
            break
    return s


def _synthetic_data() -> dict:
    rng = d.make_rng("p2r")
    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    revenue = d.revenue_series(NUM_MONTHS, rng, base=420_000)
    # Payroll ratio drifts from ~36% to ~41% to show pressure
    payroll = []
    for i, r in enumerate(revenue):
        target = 0.36 + (i / NUM_MONTHS) * 0.05
        noise = rng.uniform(-0.01, 0.01)
        payroll.append(round(r * (target + noise), 0))
    return {"months": months, "revenue": revenue, "payroll": payroll}


def build_data(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = DATA_COL_M_LAST + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 28)
    ws.set_column(DATA_COL_M1, DATA_COL_M_LAST, 12)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Source data",
        title="Revenue and total payroll, by month",
        last_col=LAST_COL + 1,
        explanation=(
            "Paste twelve months of revenue and total payroll (including on-costs) into "
            "the input cells. Total payroll should include wages and salaries, "
            "superannuation, workers compensation, and any payroll tax."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Step 1", title="Paste monthly figures")
    month_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": b.WHITE, "bg_color": b.GREY_BLACK,
        "align": "right", "valign": "vcenter", "num_format": "mmm yy",
        "top": 2, "top_color": b.GREEN, "bottom": 2, "bottom_color": b.GREEN,
    })
    ws.set_row(DATA_ROW_HEADER - 1, 26)
    ws.write_string(DATA_ROW_HEADER - 1, DATA_COL_LABEL, "Metric", formats["th"])
    for i, m in enumerate(dat["months"]):
        ws.write_datetime(DATA_ROW_HEADER - 1, DATA_COL_M1 + i, m, month_fmt)

    # Revenue row
    ws.set_row(DATA_ROW_REV - 1, 24)
    ws.write_string(DATA_ROW_REV - 1, DATA_COL_LABEL, "Revenue", formats["td"])
    for i, v in enumerate(dat["revenue"]):
        ws.write_number(DATA_ROW_REV - 1, DATA_COL_M1 + i, v, formats["input_value"])

    # Payroll row
    ws.set_row(DATA_ROW_PAYROLL - 1, 24)
    ws.write_string(DATA_ROW_PAYROLL - 1, DATA_COL_LABEL, "Total payroll cost", formats["td_zebra"])
    for i, v in enumerate(dat["payroll"]):
        ws.write_number(DATA_ROW_PAYROLL - 1, DATA_COL_M1 + i, v, formats["input_value"])

    sc.apply_page_setup(ws, sheet_title="Data")


def build_internal(wb, formats):
    ws = wb.add_worksheet("Internal Data Measures")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 5
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 24)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 24)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Inputs that do not come from the P&L",
        title="Internal Data Measures",
        last_col=LAST_COL + 1,
        explanation=(
            "Department allocation of total payroll cost. These shares vary by company "
            "and cannot be exported from the P&L; replace the amber-bordered cells with "
            "your own. Used on: By Department tab."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Allocation", title="Department wage share")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Department", "Share %", "Used on"],
                        right_align_from=2)

    # Standard allocation
    dept_shares = [
        ("Delivery", 0.55),
        ("Sales", 0.18),
        ("Admin", 0.15),
        ("Finance", 0.08),
        ("Executive", 0.04),
    ]
    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (dept, share) in enumerate(dept_shares):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, dept, label_fmt)
        ws.write_number(r, 2, share, formats["input_pct"])
        ws.write_string(r, 3, "By Department", formats["td_zebra"] if zebra else formats["td"])
        r += 1
    last_data_row_1b = r

    st.write_total_row(
        ws, formats, row=r, label="Total",
        formulas=[f"=SUM(C{first_data_row_1b}:C{last_data_row_1b})", ""],
        cell_format="total_pct",
    )

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Department shares sum to 100 per cent",
             "left": f"=SUM(C{first_data_row_1b}:C{last_data_row_1b})",
             "right": "=1",
             "is_pct": True},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Internal Data Measures")


def build_trend(wb, formats, dat):
    ws = wb.add_worksheet("Trend")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = DATA_COL_M_LAST + 2
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 26)
    ws.set_column(2, DATA_COL_M_LAST + 1, 12)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Operating ratio",
        title="Payroll-to-revenue ratio over twelve months",
        last_col=LAST_COL + 1,
        explanation=(
            "Total payroll divided by revenue per month. Watch for the ratio drifting "
            "upward while revenue remains flat, which signals either over-hiring or "
            "under-pricing. The sparkline gives a quick visual read across the year."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Monthly", title="Ratio trend")
    header_row = section_row + 2
    month_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": b.WHITE, "bg_color": b.GREY_BLACK,
        "align": "right", "valign": "vcenter", "num_format": "mmm yy",
        "top": 2, "top_color": b.GREEN, "bottom": 2, "bottom_color": b.GREEN,
    })
    ws.set_row(header_row, 26)
    ws.write_string(header_row, 1, "Metric", formats["th"])
    for i, m in enumerate(dat["months"]):
        ws.write_datetime(header_row, 2 + i, m, month_fmt)
    ws.write_string(header_row, 2 + NUM_MONTHS, "Trend", formats["th_right"])

    rows_def = [
        ("Revenue", f"=Data!{col_letter(DATA_COL_M1)}{DATA_ROW_REV}:{col_letter(DATA_COL_M_LAST)}{DATA_ROW_REV}", "aud"),
        ("Total payroll", f"=Data!{col_letter(DATA_COL_M1)}{DATA_ROW_PAYROLL}:{col_letter(DATA_COL_M_LAST)}{DATA_ROW_PAYROLL}", "aud"),
        ("Payroll % of revenue", None, "pct"),
    ]
    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (label, _, kind) in enumerate(rows_def):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats[("td_pct_zebra" if kind == "pct" else "td_right_zebra") if zebra else ("td_pct" if kind == "pct" else "td_right")]
        ws.set_row(r, 22)
        ws.write_string(r, 1, label, label_fmt)
        for j in range(NUM_MONTHS):
            col = 2 + j
            data_col = col_letter(DATA_COL_M1 + j)
            if label == "Revenue":
                ws.write_formula(r, col, f"=Data!{data_col}{DATA_ROW_REV}", num_fmt)
            elif label == "Total payroll":
                ws.write_formula(r, col, f"=Data!{data_col}{DATA_ROW_PAYROLL}", num_fmt)
            else:
                ws.write_formula(r, col, f"=IFERROR(Data!{data_col}{DATA_ROW_PAYROLL}/Data!{data_col}{DATA_ROW_REV},0)", num_fmt)
        # Sparkline cell
        sl_first = col_letter(2)
        sl_last = col_letter(2 + NUM_MONTHS - 1)
        sl_range = f"Trend!${sl_first}${r + 1}:${sl_last}${r + 1}"
        sl_anchor = f"{col_letter(2 + NUM_MONTHS)}{r + 1}"
        st.add_sparkline_line(ws, anchor_cell=sl_anchor, range_str=sl_range)
        r += 1
    last_data_row_1b = r

    # Three-color scale on the ratio row only
    ratio_first_col = col_letter(2)
    ratio_last_col = col_letter(2 + NUM_MONTHS - 1)
    ratio_range = f"{ratio_first_col}{first_data_row_1b + 2}:{ratio_last_col}{first_data_row_1b + 2}"
    st.add_three_color_scale(ws, ratio_range, favourable_high=False)

    # Legend
    r_legend = r + 2
    st.write_cf_legend(ws, formats, row=r_legend, col=1, favourable_high=False,
                       metric_label="Payroll ratio (lower is better)")

    # Tie-out
    r_check = r_legend + 2
    st.write_checks_block(
        ws, formats, row=r_check,
        checks=[
            {"name": "Sum of monthly revenue equals Data sheet revenue total",
             "left": f"=SUM({ratio_first_col}{first_data_row_1b}:{ratio_last_col}{first_data_row_1b})",
             "right": f"=SUM(Data!{col_letter(DATA_COL_M1)}{DATA_ROW_REV}:{col_letter(DATA_COL_M_LAST)}{DATA_ROW_REV})"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Trend")


def build_by_dept(wb, formats, dat):
    ws = wb.add_worksheet("By Department")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = DATA_COL_M_LAST + 2
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 12)
    ws.set_column(3, DATA_COL_M_LAST + 1, 11)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Department breakdown",
        title="Payroll cost allocated by department",
        last_col=LAST_COL + 1,
        explanation=(
            "Monthly payroll cost split using the department shares on the Internal Data "
            "Measures sheet. Use this to spot which department drove a shift in the "
            "overall payroll ratio."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Allocation", title="Payroll by department")
    header_row = section_row + 2
    month_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": b.WHITE, "bg_color": b.GREY_BLACK,
        "align": "right", "valign": "vcenter", "num_format": "mmm yy",
        "top": 2, "top_color": b.GREEN, "bottom": 2, "bottom_color": b.GREEN,
    })
    ws.set_row(header_row, 26)
    ws.write_string(header_row, 1, "Department", formats["th"])
    ws.write_string(header_row, 2, "Share", formats["th_right"])
    for i, m in enumerate(dat["months"]):
        ws.write_datetime(header_row, 3 + i, m, month_fmt)

    DEPT_ROW_FIRST = 10  # 1-based row on Internal Data Measures of first dept
    num_depts = 5
    r = header_row + 1
    first_data_row_1b = r + 1
    for i in range(num_depts):
        share_row_1b = DEPT_ROW_FIRST + i
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        pct_fmt = formats["td_pct_zebra"] if zebra else formats["td_pct"]
        ws.set_row(r, 22)
        ws.write_formula(r, 1, f"='Internal Data Measures'!B{share_row_1b}", label_fmt)
        ws.write_formula(r, 2, f"='Internal Data Measures'!C{share_row_1b}", pct_fmt)
        for j in range(NUM_MONTHS):
            payroll_col = col_letter(DATA_COL_M1 + j)
            ws.write_formula(r, 3 + j,
                             f"=Data!{payroll_col}{DATA_ROW_PAYROLL}*$C{r + 1}",
                             num_fmt)
        r += 1
    last_data_row_1b = r

    st.write_total_row(
        ws, formats, row=r, label="Total payroll",
        formulas=[""] + [f"=SUM({col_letter(3 + j)}{first_data_row_1b}:{col_letter(3 + j)}{last_data_row_1b})" for j in range(NUM_MONTHS)],
    )

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Department shares sum to 100 per cent",
             "left": f"=SUM(C{first_data_row_1b}:C{last_data_row_1b})",
             "right": "=1",
             "is_pct": True},
            {"name": "Monthly allocated payroll ties to Data sheet total",
             "left": f"=SUM(D{first_data_row_1b}:O{last_data_row_1b})",
             "right": f"=SUM(Data!{col_letter(DATA_COL_M1)}{DATA_ROW_PAYROLL}:{col_letter(DATA_COL_M_LAST)}{DATA_ROW_PAYROLL})"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="By Department")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE,
                       workbook_kicker=WORKBOOK_KICKER,
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_trend(wb, formats, dat)
    build_by_dept(wb, formats, dat)
    build_data(wb, formats, dat)
    build_internal(wb, formats)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
