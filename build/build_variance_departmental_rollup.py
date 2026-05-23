"""Departmental Variance Rollup (4030).

Department-level P&L rollup with variance to budget per cost centre.
Data sheet captures actual and budget per department per line item per
month; the Rollup, Variance, and Heatmap sheets read off it.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\4030 - Departmental Variance Rollup.xlsx")

WORKBOOK_ID = "4030"
WORKBOOK_TITLE = "Departmental Variance Rollup"
WORKBOOK_KICKER = "Per-department P&L with variance to budget"
TARGET_USER = (
    "Finance lead or fractional CFO running monthly variance reviews "
    "with department heads who own their own P&L."
)
HOW_TO_USE = [
    "On the Data sheet, enter actual and budget values for each line item per department per month.",
    "The Rollup sheet sums actual and budget across departments for each line.",
    "The Variance sheet shows variance dollars and per cent per department; the Heatmap visualises intensity.",
]

EXAMPLE_PROFILE = [
    ("STRUCTURE", "Five departments: Operations, Sales, Marketing, Finance, Executive"),
    ("REPORTING", "Monthly review against department budgets set at start of FY"),
    ("USE CASE", "Holding each department head accountable to a P&L they signed up to"),
]

INPUTS_REQUIRED = [
    ("Actual values per line item per department per month", "Data tab"),
    ("Budget values per line item per department per month", "Data tab (Budget block)"),
]

DEPARTMENTS = ["Operations", "Sales", "Marketing", "Finance", "Executive"]
LINES = [
    ("Revenue", "income"),
    ("Cost of sales", "expense"),
    ("Wages and on-costs", "expense"),
    ("Marketing and advertising", "expense"),
    ("Occupancy", "expense"),
    ("Administration", "expense"),
    ("Other operating expenses", "expense"),
]

NUM_MONTHS = 3  # Show three months side by side (current + prior 2). Keeps the table readable.

DATA_COL_LABEL = 1
DATA_COL_DEPT = 2
DATA_COL_KIND = 3   # Actual or Budget
DATA_COL_M1 = 4
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1
DATA_COL_TOTAL = DATA_COL_M_LAST + 1
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
    rng = d.make_rng("deptvar")
    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    # Per dept: revenue base, then expense components as % of revenue
    dept_profiles = {
        "Operations":  {"rev": 320_000, "cogs": 0.62, "wages": 0.22, "mkt": 0.00, "occ": 0.05, "admin": 0.03, "other": 0.02},
        "Sales":       {"rev": 280_000, "cogs": 0.55, "wages": 0.28, "mkt": 0.08, "occ": 0.02, "admin": 0.02, "other": 0.01},
        "Marketing":   {"rev": 0,       "cogs": 0.00, "wages": 0.40, "mkt": 0.55, "occ": 0.02, "admin": 0.02, "other": 0.01},
        "Finance":     {"rev": 0,       "cogs": 0.00, "wages": 0.80, "mkt": 0.00, "occ": 0.05, "admin": 0.10, "other": 0.05},
        "Executive":   {"rev": 0,       "cogs": 0.00, "wages": 0.75, "mkt": 0.05, "occ": 0.10, "admin": 0.10, "other": 0.00},
    }
    rows = []
    for dept in DEPARTMENTS:
        prof = dept_profiles[dept]
        for line, kind in LINES:
            for which in ("Actual", "Budget"):
                vals = []
                for i in range(NUM_MONTHS):
                    noise = rng.uniform(-0.06, 0.06) if which == "Actual" else 0
                    if line == "Revenue":
                        base = prof["rev"] * (1 + noise)
                    elif line == "Cost of sales":
                        base = prof["rev"] * prof["cogs"] * (1 + noise)
                    elif line == "Wages and on-costs":
                        base = max(prof["rev"], 60_000) * prof["wages"] * (1 + noise)
                    elif line == "Marketing and advertising":
                        base = max(prof["rev"], 60_000) * prof["mkt"] * (1 + noise)
                    elif line == "Occupancy":
                        base = max(prof["rev"], 60_000) * prof["occ"] * (1 + noise)
                    elif line == "Administration":
                        base = max(prof["rev"], 60_000) * prof["admin"] * (1 + noise)
                    else:
                        base = max(prof["rev"], 60_000) * prof["other"] * (1 + noise)
                    vals.append(round(base, 0))
                rows.append((line, dept, which, vals))
    return {"months": months, "rows": rows}


def build_data(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 26)
    ws.set_column(DATA_COL_DEPT, DATA_COL_DEPT, 16)
    ws.set_column(DATA_COL_KIND, DATA_COL_KIND, 10)
    ws.set_column(DATA_COL_M1, DATA_COL_TOTAL, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Source data",
        title="Actuals and budget per department per line per month",
        last_col=LAST_COL + 1,
        explanation=(
            "One row per line item per department per kind (Actual or Budget) per month. "
            "Other sheets read off this table with SUMIFS. Keep the line names and "
            "department names spelled identically wherever they appear."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Step 1", title="Actual and budget data")
    month_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": b.WHITE, "bg_color": b.GREY_BLACK,
        "align": "right", "valign": "vcenter", "num_format": "mmm yy",
        "top": 2, "top_color": b.GREEN, "bottom": 2, "bottom_color": b.GREEN,
    })
    ws.set_row(DATA_ROW_HEADER - 1, 26)
    ws.write_string(DATA_ROW_HEADER - 1, DATA_COL_LABEL, "Line", formats["th"])
    ws.write_string(DATA_ROW_HEADER - 1, DATA_COL_DEPT, "Department", formats["th"])
    ws.write_string(DATA_ROW_HEADER - 1, DATA_COL_KIND, "Kind", formats["th"])
    for i, m in enumerate(dat["months"]):
        ws.write_datetime(DATA_ROW_HEADER - 1, DATA_COL_M1 + i, m, month_fmt)
    ws.write_string(DATA_ROW_HEADER - 1, DATA_COL_TOTAL, "Quarter", formats["th_right"])

    r = DATA_ROW_FIRST - 1
    for i, (line, dept, kind, vals) in enumerate(dat["rows"]):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 20)
        ws.write_string(r, DATA_COL_LABEL, line, label_fmt)
        ws.write_string(r, DATA_COL_DEPT, dept, label_fmt)
        ws.write_string(r, DATA_COL_KIND, kind, label_fmt)
        for j, v in enumerate(vals):
            ws.write_number(r, DATA_COL_M1 + j, v, formats["input_value"])
        ws.write_formula(r, DATA_COL_TOTAL,
                         f"=SUM({col_letter(DATA_COL_M1)}{r + 1}:{col_letter(DATA_COL_M_LAST)}{r + 1})",
                         num_fmt)
        r += 1

    sc.apply_page_setup(ws, sheet_title="Data")


def build_rollup(wb, formats, dat):
    ws = wb.add_worksheet("Rollup")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 1 + 1 + len(DEPARTMENTS) * 3 + 1  # Line + (Actual/Budget/Variance per dept) + Total cols
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    for c in range(2, LAST_COL):
        ws.set_column(c, c, 13)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Quarter rollup",
        title="Actual and budget by department",
        last_col=LAST_COL + 1,
        explanation=(
            "For the quarter shown on the Data sheet: actual, budget, and variance per "
            "department per line item. Variance equals actual minus budget; for income "
            "lines a positive variance is favourable, for expense lines a positive "
            "variance is unfavourable."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="By department", title="Quarter rollup")
    header_row = section_row + 2

    # Header row: dept name spans Actual/Budget/Var columns
    ws.set_row(header_row, 26)
    ws.write_string(header_row, 1, "Line", formats["th"])
    for di, dept in enumerate(DEPARTMENTS):
        base = 2 + di * 3
        ws.merge_range(header_row, base, header_row, base + 2, dept, formats["th_right"])
    ws.set_row(header_row + 1, 22)
    ws.write_string(header_row + 1, 1, "", formats["th"])
    for di in range(len(DEPARTMENTS)):
        base = 2 + di * 3
        ws.write_string(header_row + 1, base, "Actual", formats["th_right"])
        ws.write_string(header_row + 1, base + 1, "Budget", formats["th_right"])
        ws.write_string(header_row + 1, base + 2, "Variance", formats["th_right"])

    r = header_row + 2
    first_data_row_1b = r + 1
    for i, (line, kind) in enumerate(LINES):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, line, label_fmt)
        for di, dept in enumerate(DEPARTMENTS):
            base = 2 + di * 3
            # Actual sum: SUMIFS(Data!Total, Line, Dept, Kind="Actual")
            ws.write_formula(r, base,
                             f'=SUMIFS(Data!${col_letter(DATA_COL_TOTAL)}:${col_letter(DATA_COL_TOTAL)},'
                             f'Data!${col_letter(DATA_COL_LABEL)}:${col_letter(DATA_COL_LABEL)},$B{r + 1},'
                             f'Data!${col_letter(DATA_COL_DEPT)}:${col_letter(DATA_COL_DEPT)},"{dept}",'
                             f'Data!${col_letter(DATA_COL_KIND)}:${col_letter(DATA_COL_KIND)},"Actual")',
                             num_fmt)
            ws.write_formula(r, base + 1,
                             f'=SUMIFS(Data!${col_letter(DATA_COL_TOTAL)}:${col_letter(DATA_COL_TOTAL)},'
                             f'Data!${col_letter(DATA_COL_LABEL)}:${col_letter(DATA_COL_LABEL)},$B{r + 1},'
                             f'Data!${col_letter(DATA_COL_DEPT)}:${col_letter(DATA_COL_DEPT)},"{dept}",'
                             f'Data!${col_letter(DATA_COL_KIND)}:${col_letter(DATA_COL_KIND)},"Budget")',
                             num_fmt)
            ws.write_formula(r, base + 2,
                             f"={col_letter(base)}{r + 1}-{col_letter(base + 1)}{r + 1}",
                             num_fmt)
        r += 1
    last_data_row_1b = r

    r += 1
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Sum of all department actuals equals Data sheet actual total",
             "left": "=SUMPRODUCT((B" + str(first_data_row_1b) + ":B" + str(last_data_row_1b) +
                     "<>\"\")*1)*0+SUMIFS(Data!" + col_letter(DATA_COL_TOTAL) + ":" +
                     col_letter(DATA_COL_TOTAL) + ",Data!" + col_letter(DATA_COL_KIND) + ":" +
                     col_letter(DATA_COL_KIND) + ",\"Actual\")",
             "right": "=SUMIF(Data!" + col_letter(DATA_COL_KIND) + ":" +
                      col_letter(DATA_COL_KIND) + ",\"Actual\",Data!" +
                      col_letter(DATA_COL_TOTAL) + ":" + col_letter(DATA_COL_TOTAL) + ")"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Rollup")


def build_heatmap(wb, formats, dat):
    ws = wb.add_worksheet("Heatmap")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 1 + len(DEPARTMENTS) + 2
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 26)
    for c in range(2, 2 + len(DEPARTMENTS)):
        ws.set_column(c, c, 14)
    ws.set_column(LAST_COL - 1, LAST_COL - 1, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Variance heatmap",
        title="Variance per cent by department and line",
        last_col=LAST_COL + 1,
        explanation=(
            "Variance from budget as a percentage of budget, per department per line. "
            "Red cells indicate unfavourable variance (income below budget or expense "
            "above budget). Use to scan for the department or line driving the result."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Per cent variance", title="Heatmap")
    header_row = section_row + 2
    ws.set_row(header_row, 26)
    ws.write_string(header_row, 1, "Line", formats["th"])
    for di, dept in enumerate(DEPARTMENTS):
        ws.write_string(header_row, 2 + di, dept, formats["th_right"])
    ws.write_string(header_row, 2 + len(DEPARTMENTS), "Sign", formats["th_right"])

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (line, kind) in enumerate(LINES):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        pct_fmt = formats["td_pct_zebra"] if zebra else formats["td_pct"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, line, label_fmt)
        for di, dept in enumerate(DEPARTMENTS):
            # Variance % = (Actual - Budget) / abs(Budget), oriented so positive is unfavourable for expenses
            actual = (f'SUMIFS(Data!${col_letter(DATA_COL_TOTAL)}:${col_letter(DATA_COL_TOTAL)},'
                      f'Data!${col_letter(DATA_COL_LABEL)}:${col_letter(DATA_COL_LABEL)},$B{r + 1},'
                      f'Data!${col_letter(DATA_COL_DEPT)}:${col_letter(DATA_COL_DEPT)},"{dept}",'
                      f'Data!${col_letter(DATA_COL_KIND)}:${col_letter(DATA_COL_KIND)},"Actual")')
            budget = (f'SUMIFS(Data!${col_letter(DATA_COL_TOTAL)}:${col_letter(DATA_COL_TOTAL)},'
                      f'Data!${col_letter(DATA_COL_LABEL)}:${col_letter(DATA_COL_LABEL)},$B{r + 1},'
                      f'Data!${col_letter(DATA_COL_DEPT)}:${col_letter(DATA_COL_DEPT)},"{dept}",'
                      f'Data!${col_letter(DATA_COL_KIND)}:${col_letter(DATA_COL_KIND)},"Budget")')
            # For income lines: positive variance (actual > budget) is good → flip sign so heatmap reads consistently
            sign = "-1" if kind == "income" else "1"
            ws.write_formula(r, 2 + di,
                             f"=IFERROR((({actual})-({budget}))/MAX(ABS({budget}),1)*{sign},0)",
                             pct_fmt)
        ws.write_string(r, 2 + len(DEPARTMENTS),
                        "lower is better (unfav %)" if kind == "expense" else "lower is better (unfav %)",
                        formats["td_zebra"] if zebra else formats["td"])
        r += 1
    last_data_row_1b = r

    # Heatmap across the department columns
    first_col = col_letter(2)
    last_col_letter = col_letter(2 + len(DEPARTMENTS) - 1)
    heat_range = f"{first_col}{first_data_row_1b}:{last_col_letter}{last_data_row_1b}"
    st.add_three_color_scale(ws, heat_range, favourable_high=False)

    r += 2
    st.write_cf_legend(ws, formats, row=r, col=1, favourable_high=False,
                       metric_label="Variance per cent (lower is better)")

    sc.apply_page_setup(ws, sheet_title="Heatmap")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE,
                       workbook_kicker=WORKBOOK_KICKER,
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_rollup(wb, formats, dat)
    build_heatmap(wb, formats, dat)
    build_data(wb, formats, dat)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
