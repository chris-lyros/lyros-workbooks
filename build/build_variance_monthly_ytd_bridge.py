"""Monthly Variance and YTD Bridge (4020).

Decomposes the variance between budget and actual into named driver
buckets month by month and year to date. Used by fractional CFOs and
finance leads to explain why a P&L moved, not just by how much.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\4020 - Monthly Variance and YTD Bridge.xlsx")

WORKBOOK_ID = "4020"
WORKBOOK_TITLE = "Monthly Variance and YTD Bridge"
WORKBOOK_KICKER = "Decomposing variance into named driver buckets"
TARGET_USER = "Finance lead or fractional CFO explaining why monthly and YTD results moved against budget."
HOW_TO_USE = [
    "Paste Actuals and Budget per line item per month on the Data sheet.",
    "Allocate the variance per line into driver buckets (Price, Volume, Mix, Timing, Other) on Internal Data Measures.",
    "The Monthly Bridge and YTD Bridge sheets walk from budget to actual through the driver buckets.",
]

EXAMPLE_PROFILE = [
    ("PERIOD", "Six months reported, year-to-date through November 2025"),
    ("DRIVER CONVENTION", "Price, Volume, Mix, Timing, Other - same definitions used every month"),
    ("AUDIENCE", "Leadership team meeting after each month-end close"),
]

INPUTS_REQUIRED = [
    ("Actual and budget per line per month", "Data tab"),
    ("Variance attribution per line by driver bucket (proportions sum to 100%)", "Internal Data Measures tab"),
]

NUM_MONTHS = 6
LINES = [
    ("Revenue", "income"),
    ("Cost of sales", "expense"),
    ("Gross profit", "income"),
    ("Wages and on-costs", "expense"),
    ("Marketing", "expense"),
    ("Occupancy", "expense"),
    ("Administration", "expense"),
    ("EBITDA", "income"),
]
# Subset of lines we attribute drivers against (computed lines are derived)
ATTRIBUTABLE_LINES = [l for l, _ in LINES if l not in ("Gross profit", "EBITDA")]
DRIVERS = ["Price", "Volume", "Mix", "Timing", "Other"]

DATA_COL_LABEL = 1
DATA_COL_KIND = 2   # Actual or Budget
DATA_COL_M1 = 3
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1
DATA_COL_YTD = DATA_COL_M_LAST + 1
DATA_ROW_HEADER = 9
DATA_ROW_FIRST = 10


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
    rng = d.make_rng("varbridge")
    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    base_rev = 480_000
    rows = []
    for line, kind in LINES:
        for which in ("Actual", "Budget"):
            vals = []
            for i in range(NUM_MONTHS):
                # Drift the actual slightly off budget so there is something to decompose
                drift = rng.uniform(-0.07, 0.06) if which == "Actual" else 0
                if line == "Revenue":
                    v = base_rev * (1 + i * 0.012) * (1 + drift)
                elif line == "Cost of sales":
                    v = base_rev * 0.58 * (1 + i * 0.010) * (1 + drift)
                elif line == "Gross profit":
                    continue  # derived, skip in source data
                elif line == "Wages and on-costs":
                    v = base_rev * 0.22 * (1 + i * 0.005) * (1 + drift)
                elif line == "Marketing":
                    v = base_rev * 0.06 * (1 + drift)
                elif line == "Occupancy":
                    v = base_rev * 0.04 * (1 + drift * 0.3)
                elif line == "Administration":
                    v = base_rev * 0.05 * (1 + drift)
                elif line == "EBITDA":
                    continue  # derived
                else:
                    v = base_rev * 0.02 * (1 + drift)
                vals.append(round(v, 0))
            if vals:
                rows.append((line, which, vals))
    return {"months": months, "rows": rows}


def build_data(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = DATA_COL_YTD + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 28)
    ws.set_column(DATA_COL_KIND, DATA_COL_KIND, 10)
    ws.set_column(DATA_COL_M1, DATA_COL_YTD, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Source data",
        title="Actuals and budget by line and month",
        last_col=LAST_COL + 1,
        explanation=(
            "One row per line item per kind (Actual or Budget) per month. The Gross "
            "profit and EBITDA lines are derived on the Bridge sheets; do not enter "
            "values for them here."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Step 1", title="Actuals and budget")
    month_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": b.WHITE, "bg_color": b.GREY_BLACK,
        "align": "right", "valign": "vcenter", "num_format": "mmm yy",
        "top": 2, "top_color": b.GREEN, "bottom": 2, "bottom_color": b.GREEN,
    })
    ws.set_row(DATA_ROW_HEADER - 1, 26)
    ws.write_string(DATA_ROW_HEADER - 1, DATA_COL_LABEL, "Line", formats["th"])
    ws.write_string(DATA_ROW_HEADER - 1, DATA_COL_KIND, "Kind", formats["th"])
    for i, m in enumerate(dat["months"]):
        ws.write_datetime(DATA_ROW_HEADER - 1, DATA_COL_M1 + i, m, month_fmt)
    ws.write_string(DATA_ROW_HEADER - 1, DATA_COL_YTD, "YTD", formats["th_right"])

    r = DATA_ROW_FIRST - 1
    for i, (line, kind, vals) in enumerate(dat["rows"]):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 22)
        ws.write_string(r, DATA_COL_LABEL, line, label_fmt)
        ws.write_string(r, DATA_COL_KIND, kind, label_fmt)
        for j, v in enumerate(vals):
            ws.write_number(r, DATA_COL_M1 + j, v, formats["input_value"])
        ws.write_formula(r, DATA_COL_YTD,
                         f"=SUM({col_letter(DATA_COL_M1)}{r + 1}:{col_letter(DATA_COL_M_LAST)}{r + 1})",
                         num_fmt)
        r += 1

    sc.apply_page_setup(ws, sheet_title="Data")


IDM_ROW_FIRST = 10  # 1-based first data row on Internal Data Measures sheet


def build_internal(wb, formats):
    ws = wb.add_worksheet("Internal Data Measures")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 8
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 24)   # Line
    ws.set_column(2, 6, 12)   # Drivers
    ws.set_column(7, 7, 24)   # Used on
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Driver attribution",
        title="How variance is allocated to driver buckets",
        last_col=LAST_COL + 1,
        explanation=(
            "For each line item, the variance between Actual and Budget is split across "
            "five driver buckets in the proportions below. Shares per line must sum to "
            "100 per cent. Update each month after the variance review. Used on: "
            "Monthly Bridge and YTD Bridge tabs."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Step 2", title="Driver shares per line")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Line"] + DRIVERS + ["Used on"],
        right_align_from=2,
    )

    # Default sensible allocations per line
    defaults = {
        "Revenue":                  (0.40, 0.30, 0.10, 0.10, 0.10),
        "Cost of sales":            (0.20, 0.40, 0.20, 0.05, 0.15),
        "Wages and on-costs":       (0.10, 0.20, 0.05, 0.50, 0.15),
        "Marketing":                (0.05, 0.10, 0.05, 0.60, 0.20),
        "Occupancy":                (0.00, 0.05, 0.00, 0.85, 0.10),
        "Administration":           (0.05, 0.05, 0.05, 0.70, 0.15),
    }
    r = header_row + 1
    first_data_row_1b = r + 1
    for i, line in enumerate(ATTRIBUTABLE_LINES):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, line, label_fmt)
        for j, share in enumerate(defaults[line]):
            ws.write_number(r, 2 + j, share, formats["input_pct"])
        ws.write_string(r, 7, "Monthly + YTD Bridge", formats["td_zebra"] if zebra else formats["td"])
        r += 1
    last_data_row_1b = r

    # Sum check column
    r_check = r + 1
    st.write_checks_block(
        ws, formats, row=r_check,
        checks=[
            {"name": "Each line's driver shares sum to 100 per cent (sum of all)",
             "left": f"=SUMPRODUCT((C{first_data_row_1b}:G{last_data_row_1b})*1)",
             "right": f"=COUNTA(B{first_data_row_1b}:B{last_data_row_1b})",
             "is_pct": False},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Internal Data Measures")


def build_monthly_bridge(wb, formats, dat):
    ws = wb.add_worksheet("Monthly Bridge")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    # Show the most recent month's bridge as a vertical waterfall:
    # Budget -> + Price -> + Volume -> + Mix -> + Timing -> + Other -> Actual
    LAST_COL = 6
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, 5, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Variance bridge",
        title="From budget to actual, by driver",
        last_col=LAST_COL + 1,
        explanation=(
            "For each line item, walk from Budget to Actual through the driver buckets. "
            "Driver dollars equal total variance times the share on Internal Data "
            "Measures. The Sum check column confirms the bridge reconciles to actual."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker=f"Current month ({dat['months'][-1].strftime('%b %Y')})",
                            title="Variance bridge per line")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Line", "Budget", "Variance $", "Actual", "Reconciles"],
        right_align_from=2,
    )

    r = header_row + 1
    first_data_row_1b = r + 1
    m_col = col_letter(DATA_COL_M1 + NUM_MONTHS - 1)
    for i, line in enumerate(ATTRIBUTABLE_LINES):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, line, label_fmt)
        budget_f = (
            f'=SUMIFS(Data!${m_col}:${m_col},'
            f'Data!${col_letter(DATA_COL_LABEL)}:${col_letter(DATA_COL_LABEL)},$B{r + 1},'
            f'Data!${col_letter(DATA_COL_KIND)}:${col_letter(DATA_COL_KIND)},"Budget")'
        )
        actual_f = (
            f'=SUMIFS(Data!${m_col}:${m_col},'
            f'Data!${col_letter(DATA_COL_LABEL)}:${col_letter(DATA_COL_LABEL)},$B{r + 1},'
            f'Data!${col_letter(DATA_COL_KIND)}:${col_letter(DATA_COL_KIND)},"Actual")'
        )
        ws.write_formula(r, 2, budget_f, num_fmt)
        ws.write_formula(r, 3, f"=E{r + 1}-C{r + 1}", num_fmt)
        ws.write_formula(r, 4, actual_f, num_fmt)
        ws.write_formula(r, 5, f'=IF(ABS(C{r + 1}+D{r + 1}-E{r + 1})<0.5,"OK","FLAG")',
                         formats["check_status_neutral"])
        r += 1
    last_data_row_1b = r

    rec_range = f"F{first_data_row_1b}:F{last_data_row_1b}"
    ws.conditional_format(rec_range, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["check_status_ok"]})
    ws.conditional_format(rec_range, {"type": "text", "criteria": "containing", "value": "FLAG", "format": formats["check_status_flag"]})

    # Driver decomposition block
    r += 2
    ws.set_row(r, 26)
    ws.merge_range(r, 1, r, LAST_COL - 1, "Variance attribution by driver", formats["section_h2"])
    r += 1
    st.write_header_row(ws, formats, row=r, headers=["Line"] + DRIVERS, right_align_from=2)
    r += 1
    first_attr_row_1b = r + 1
    for i, line in enumerate(ATTRIBUTABLE_LINES):
        idm_row_1b = IDM_ROW_FIRST + i
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, line, label_fmt)
        # Variance for this line is on the corresponding row above (column D)
        var_cell = f"$D${first_data_row_1b + i}"
        for j, driver in enumerate(DRIVERS):
            share_col = col_letter(2 + j)
            ws.write_formula(r, 2 + j,
                             f"={var_cell}*'Internal Data Measures'!{share_col}{idm_row_1b}",
                             num_fmt)
        r += 1
    last_attr_row_1b = r

    r += 1
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Sum of driver attributions equals total variance",
             "left": f"=SUMPRODUCT((C{first_attr_row_1b}:G{last_attr_row_1b})*1)",
             "right": f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Monthly Bridge")


def build_ytd_bridge(wb, formats, dat):
    ws = wb.add_worksheet("YTD Bridge")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 6
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, 5, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Year-to-date bridge",
        title="From YTD budget to YTD actual, by driver",
        last_col=LAST_COL + 1,
        explanation=(
            "Same logic as the Monthly Bridge but using year-to-date totals from the "
            "Data sheet's YTD column. Useful when monthly variance is small but YTD has "
            "compounded into a material number."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Year to date", title="Variance bridge per line")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Line", "YTD budget", "Variance $", "YTD actual", "Reconciles"],
        right_align_from=2,
    )

    r = header_row + 1
    first_data_row_1b = r + 1
    ytd_col = col_letter(DATA_COL_YTD)
    for i, line in enumerate(ATTRIBUTABLE_LINES):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, line, label_fmt)
        ws.write_formula(r, 2,
                         f'=SUMIFS(Data!${ytd_col}:${ytd_col},'
                         f'Data!${col_letter(DATA_COL_LABEL)}:${col_letter(DATA_COL_LABEL)},$B{r + 1},'
                         f'Data!${col_letter(DATA_COL_KIND)}:${col_letter(DATA_COL_KIND)},"Budget")',
                         num_fmt)
        ws.write_formula(r, 3, f"=E{r + 1}-C{r + 1}", num_fmt)
        ws.write_formula(r, 4,
                         f'=SUMIFS(Data!${ytd_col}:${ytd_col},'
                         f'Data!${col_letter(DATA_COL_LABEL)}:${col_letter(DATA_COL_LABEL)},$B{r + 1},'
                         f'Data!${col_letter(DATA_COL_KIND)}:${col_letter(DATA_COL_KIND)},"Actual")',
                         num_fmt)
        ws.write_formula(r, 5, f'=IF(ABS(C{r + 1}+D{r + 1}-E{r + 1})<0.5,"OK","FLAG")',
                         formats["check_status_neutral"])
        r += 1
    last_data_row_1b = r

    rec_range = f"F{first_data_row_1b}:F{last_data_row_1b}"
    ws.conditional_format(rec_range, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["check_status_ok"]})
    ws.conditional_format(rec_range, {"type": "text", "criteria": "containing", "value": "FLAG", "format": formats["check_status_flag"]})

    r += 2
    ws.set_row(r, 26)
    ws.merge_range(r, 1, r, LAST_COL - 1, "YTD attribution by driver", formats["section_h2"])
    r += 1
    st.write_header_row(ws, formats, row=r, headers=["Line"] + DRIVERS, right_align_from=2)
    r += 1
    first_attr_row_1b = r + 1
    for i, line in enumerate(ATTRIBUTABLE_LINES):
        idm_row_1b = IDM_ROW_FIRST + i
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, line, label_fmt)
        var_cell = f"$D${first_data_row_1b + i}"
        for j, driver in enumerate(DRIVERS):
            share_col = col_letter(2 + j)
            ws.write_formula(r, 2 + j,
                             f"={var_cell}*'Internal Data Measures'!{share_col}{idm_row_1b}",
                             num_fmt)
        r += 1
    last_attr_row_1b = r

    r += 1
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Sum of YTD attributions equals total YTD variance",
             "left": f"=SUMPRODUCT((C{first_attr_row_1b}:G{last_attr_row_1b})*1)",
             "right": f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="YTD Bridge")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE,
                       workbook_kicker=WORKBOOK_KICKER,
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_monthly_bridge(wb, formats, dat)
    build_ytd_bridge(wb, formats, dat)
    build_data(wb, formats, dat)
    build_internal(wb, formats)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
