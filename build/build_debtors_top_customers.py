"""Top Customers by Revenue workbook.

Customer-level revenue extract from your accounting software with 12-month total, year-on-year
change, share of total revenue, and concentration ratio.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\7000 - Top Customers by Revenue.xlsx")

WORKBOOK_ID = "7000"
WORKBOOK_TITLE = "Top Customers by Revenue"
WORKBOOK_KICKER = "Customer concentration and growth"
TARGET_USER = "Founder, sales lead, or CFO reviewing customer mix and concentration risk."
HOW_TO_USE = [
    "Open the Data sheet and paste your customer revenue extract (Sales by Customer Summary).",
    "Enter the prior-year revenue per customer in the second column to compute year-on-year change.",
    "The Ranking sheet sorts by current-year revenue and shows top-customer concentration.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "B2B SME with 15-20 active customers"),
    ("ANNUAL REVENUE", "Circa $4M"),
    ("TOP-5 CONCENTRATION", "Typically 45-60% of total revenue"),
]

INPUTS_REQUIRED = [
    ("Current year revenue per customer", "Data tab"),
    ("Prior year revenue per customer", "Data tab"),
]

DATA_COL_LABEL = 1
DATA_COL_CURR = 2
DATA_COL_PRIOR = 3
DATA_ROW_HEADER = 8
DATA_ROW_FIRST = 9
DATA_ROW_LAST = DATA_ROW_FIRST + 14
DATA_ROW_TOTAL = DATA_ROW_LAST + 1


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
    rng = d.make_rng("topcust")
    roster = d.roster()[:15]
    rows = []
    for co in roster:
        scale = rng.uniform(0.3, 3.0)
        curr = round(180_000 * scale + rng.uniform(-20_000, 20_000), 0)
        prior = round(curr * rng.uniform(0.78, 1.18), 0)
        rows.append((co.name, curr, prior))
    return {"rows": rows}


def build_data_sheet(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)
    LAST_COL = DATA_COL_PRIOR + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 32)
    ws.set_column(DATA_COL_CURR, DATA_COL_PRIOR, 18)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Single source of truth",
                       title="Drop your customer revenue here",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "Paste your Sales by Customer Summary for the current 12 months "
                           "and the prior 12 months. The Ranking sheet sorts by current-year revenue "
                           "and computes concentration ratio (top-5 share, top-10 share)."
                       ))

    st.write_section_header(ws, formats, row=6, kicker="Paste from your accounting software", title="Customer revenue")
    st.write_header_row(ws, formats, row=DATA_ROW_HEADER - 1,
                        headers=["Customer", "Current 12 months", "Prior 12 months"],
                        start_col=DATA_COL_LABEL, right_align_from=2)
    for i, (name, curr, prior) in enumerate(dat["rows"]):
        row_0 = DATA_ROW_FIRST - 1 + i
        ws.set_row(row_0, 22)
        zebra = i % 2 == 0
        ws.write_string(row_0, DATA_COL_LABEL, name, formats["td_zebra"] if zebra else formats["td"])
        ws.write_number(row_0, DATA_COL_CURR, curr, formats["input_value"])
        ws.write_number(row_0, DATA_COL_PRIOR, prior, formats["input_value"])

    ws.set_row(DATA_ROW_TOTAL - 1, 24)
    ws.write_string(DATA_ROW_TOTAL - 1, DATA_COL_LABEL, "Total", formats["total_left"])
    ws.write_formula(DATA_ROW_TOTAL - 1, DATA_COL_CURR,
                     f"=SUM({col_letter(DATA_COL_CURR)}{DATA_ROW_FIRST}:{col_letter(DATA_COL_CURR)}{DATA_ROW_LAST})",
                     formats["total_right"])
    ws.write_formula(DATA_ROW_TOTAL - 1, DATA_COL_PRIOR,
                     f"=SUM({col_letter(DATA_COL_PRIOR)}{DATA_ROW_FIRST}:{col_letter(DATA_COL_PRIOR)}{DATA_ROW_LAST})",
                     formats["total_right"])
    sc.apply_page_setup(ws, sheet_title="Data")


def build_ranking(wb, formats, dat):
    ws = wb.add_worksheet("Ranking")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = 7
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 8)
    ws.set_column(2, 2, 32)
    ws.set_column(3, LAST_COL, 16)
    ws.set_column(LAST_COL + 1, LAST_COL + 1, 2)

    sc.write_hero_band(ws, formats, kicker="Sorted by current-year revenue",
                       title="Top customers by revenue",
                       last_col=LAST_COL + 2,
                       explanation=(
                           "Ranked customer list with current and prior 12-month revenue, year-on-year "
                           "change, share of total, and cumulative share (used to compute concentration "
                           "ratios). Use the cumulative share column to spot the customers driving most "
                           "of your revenue."
                       ))

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet",
                            title="Ranked top customers")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Rank", "Customer", "Current 12m", "Prior 12m", "YoY $", "YoY %", "Share %", "Cum share %"],
                        start_col=1, right_align_from=2)

    # Use LARGE() to pull the nth largest current value, and INDEX/MATCH to grab the customer name.
    r = header_row + 1
    first_data_row_1b = r + 1
    num_customers = len(dat["rows"])
    data_first = DATA_ROW_FIRST
    data_last = DATA_ROW_LAST
    curr_range = f"'Data'!${col_letter(DATA_COL_CURR)}${data_first}:${col_letter(DATA_COL_CURR)}${data_last}"
    name_range = f"'Data'!${col_letter(DATA_COL_LABEL)}${data_first}:${col_letter(DATA_COL_LABEL)}${data_last}"
    prior_range = f"'Data'!${col_letter(DATA_COL_PRIOR)}${data_first}:${col_letter(DATA_COL_PRIOR)}${data_last}"

    for i in range(num_customers):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        pct_fmt = formats["td_pct_zebra"] if zebra else formats["td_pct"]
        ws.set_row(r, 22)
        ws.write_number(r, 1, i + 1, label_fmt)
        rank = i + 1
        # Customer name = INDEX(name_range, MATCH(LARGE(curr_range, rank), curr_range, 0))
        curr_value = f"LARGE({curr_range},{rank})"
        name_f = f"=INDEX({name_range},MATCH({curr_value},{curr_range},0))"
        prior_f = f"=INDEX({prior_range},MATCH({curr_value},{curr_range},0))"
        ws.write_formula(r, 2, name_f, label_fmt)
        ws.write_formula(r, 3, f"={curr_value}", num_fmt)
        ws.write_formula(r, 4, prior_f, num_fmt)
        ws.write_formula(r, 5, f"=D{r + 1}-E{r + 1}", num_fmt)
        ws.write_formula(r, 6, f"=IFERROR((D{r + 1}-E{r + 1})/ABS(E{r + 1}),0)", pct_fmt)
        ws.write_formula(r, 7, f"=IFERROR(D{r + 1}/'Data'!${col_letter(DATA_COL_CURR)}${DATA_ROW_TOTAL},0)", pct_fmt)
        if i == 0:
            ws.write_formula(r, 8, f"=H{r + 1}", pct_fmt)
        else:
            ws.write_formula(r, 8, f"=I{r}+H{r + 1}", pct_fmt)
        r += 1
    last_data_row_1b = r

    # Total row
    st.write_total_row(ws, formats, row=r, label="Total",
                       formulas=["", f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})",
                                  f"=SUM(E{first_data_row_1b}:E{last_data_row_1b})",
                                  f"=SUM(F{first_data_row_1b}:F{last_data_row_1b})",
                                  "", "", ""],
                       cell_format="total_right")

    # CF: YoY % column — favourable high
    st.add_three_color_scale(ws, f"G{first_data_row_1b}:G{last_data_row_1b}", favourable_high=True)
    st.write_cf_legend(ws, formats, row=r + 2, col=1, favourable_high=True, metric_label="YoY % (higher is better)")

    # Tie-out
    r += 4
    st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {"name": "Sum of current 12-month revenue ties to Data total",
             "left": f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})",
             "right": f"='Data'!{col_letter(DATA_COL_CURR)}{DATA_ROW_TOTAL}"},
            {"name": "Cumulative share at bottom ranks reaches 100 per cent",
             "left": f"=I{last_data_row_1b}",
             "right": "=1",
             "is_pct": True},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Ranking")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_ranking(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
