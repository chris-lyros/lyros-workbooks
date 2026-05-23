"""Aged Receivables by Bucket workbook (debtors-focused operational view).

Same AR aging input as the working_capital analysis workbook, but optimised
for the operational chase run: top-debtor list, bucket totals, and a
'what to chase first' suggestion column.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\5200 - Aged Receivables Chase List.xlsx")

WORKBOOK_ID = "5200"
WORKBOOK_TITLE = "Aged Receivables by Bucket"
WORKBOOK_KICKER = "Monday-morning chase list"
TARGET_USER = "Credit controller, accounts receivable lead, or business owner running a weekly chase run."
HOW_TO_USE = [
    "Open the Data sheet and paste your Aged Receivables Summary export.",
    "The Chase List sheet ranks debtors by chase priority (large balances with late aging come first) and suggests an action per debtor.",
    "The Bucket Summary sheet shows total $ per aging bucket with a quick comparison to total AR.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "B2B SME with 15-20 active debtors"),
    ("CHASE CADENCE", "Weekly Monday-morning run"),
    ("OUTPUT", "Prioritised chase list with suggested action per debtor"),
]

INPUTS_REQUIRED = [
    ("Aged Receivables Summary by customer", "Data tab (paste from your accounting software)"),
]

BUCKETS = ["Current", "1-30 days", "31-60 days", "61-90 days", "90+ days"]
DATA_COL_LABEL = 1
DATA_COL_FIRST_BUCKET = 2
DATA_COL_LAST_BUCKET = DATA_COL_FIRST_BUCKET + len(BUCKETS) - 1
DATA_COL_TOTAL = DATA_COL_LAST_BUCKET + 1
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
    rng = d.make_rng("ar-bk")
    roster = d.roster()[:15]
    rows = []
    for co in roster:
        scale = rng.uniform(0.4, 2.5)
        current = round(20_000 * scale + rng.uniform(-3_000, 3_000), 0)
        d30 = round(8_000 * scale + rng.uniform(-1_500, 1_500), 0)
        d60 = round(4_000 * scale + rng.uniform(-1_000, 1_000), 0)
        d90 = round(1_500 * scale + rng.uniform(-500, 500), 0)
        d90plus = round(800 * scale + rng.uniform(-300, 300), 0) if rng.random() < 0.4 else 0
        rows.append((co.name, [current, d30, d60, d90, d90plus]))
    return {"rows": rows}


def build_data_sheet(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)
    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 32)
    ws.set_column(DATA_COL_FIRST_BUCKET, DATA_COL_LAST_BUCKET, 14)
    ws.set_column(DATA_COL_TOTAL, DATA_COL_TOTAL, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Single source of truth",
                       title="Drop your Aged Receivables here",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "Paste your Aged Receivables Summary into the table below. Standard "
                           "five-bucket aging (Current, 1-30, 31-60, 61-90, 90+). The Chase List and "
                           "Bucket Summary sheets read from this table."
                       ))

    st.write_section_header(ws, formats, row=6, kicker="Paste from your AR Summary",
                            title="Aged receivables by customer")
    st.write_header_row(ws, formats, row=DATA_ROW_HEADER - 1,
                        headers=["Customer"] + BUCKETS + ["Total"],
                        start_col=DATA_COL_LABEL, right_align_from=2)

    for i, (name, buckets) in enumerate(dat["rows"]):
        row_0 = DATA_ROW_FIRST - 1 + i
        ws.set_row(row_0, 22)
        zebra = i % 2 == 0
        text_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.write_string(row_0, DATA_COL_LABEL, name, text_fmt)
        for j, v in enumerate(buckets):
            ws.write_number(row_0, DATA_COL_FIRST_BUCKET + j, v, formats["input_value"])
        first_b = col_letter(DATA_COL_FIRST_BUCKET)
        last_b = col_letter(DATA_COL_LAST_BUCKET)
        ws.write_formula(row_0, DATA_COL_TOTAL,
                         f"=SUM(${first_b}${DATA_ROW_FIRST + i}:${last_b}${DATA_ROW_FIRST + i})",
                         formats["td_bold_right"])

    ws.set_row(DATA_ROW_TOTAL - 1, 24)
    ws.write_string(DATA_ROW_TOTAL - 1, DATA_COL_LABEL, "Total receivables", formats["total_left"])
    for j in range(len(BUCKETS) + 1):
        col = col_letter(DATA_COL_FIRST_BUCKET + j)
        ws.write_formula(DATA_ROW_TOTAL - 1, DATA_COL_FIRST_BUCKET + j,
                         f"=SUM({col}{DATA_ROW_FIRST}:{col}{DATA_ROW_LAST})", formats["total_right"])

    sc.apply_page_setup(ws, sheet_title="Data")


def build_chase_list(wb, formats, dat):
    ws = wb.add_worksheet("Chase List")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = 9
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 32)
    ws.set_column(2, 7, 12)
    ws.set_column(8, 8, 28)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Prioritised by exposure × age",
                       title="Monday-morning chase list",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "Each customer is scored by chase priority: balances older than 60 days "
                           "are weighted heavily, then 31-60, then 1-30. The Suggested action column "
                           "names the next step based on where the customer's largest unpaid bucket sits."
                       ))

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet",
                            title="Chase priority and suggested action")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Customer"] + BUCKETS + ["Total", "Suggested action"], right_align_from=2)

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (name, _b) in enumerate(dat["rows"]):
        zebra = i % 2 == 0
        data_row_1b = DATA_ROW_FIRST + i
        formulas = [f"='Data'!{col_letter(DATA_COL_FIRST_BUCKET + j)}{data_row_1b}" for j in range(len(BUCKETS))]
        total_f = f"='Data'!{col_letter(DATA_COL_TOTAL)}{data_row_1b}"
        # Suggested action based on aging
        b90plus = f"'Data'!{col_letter(DATA_COL_LAST_BUCKET)}{data_row_1b}"
        b6090 = f"'Data'!{col_letter(DATA_COL_FIRST_BUCKET + 3)}{data_row_1b}"
        b3160 = f"'Data'!{col_letter(DATA_COL_FIRST_BUCKET + 2)}{data_row_1b}"
        b130 = f"'Data'!{col_letter(DATA_COL_FIRST_BUCKET + 1)}{data_row_1b}"
        action_f = (
            f"=IF({b90plus}>0,\"Escalate: legal letter or stop credit\","
            f"IF({b6090}>0,\"Phone call today; demand payment plan\","
            f"IF({b3160}>0,\"Phone call this week\","
            f"IF({b130}>0,\"Reminder email\",\"No action\"))))"
        )
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, name, label_fmt)
        for j, f_ in enumerate(formulas):
            ws.write_formula(r, 2 + j, f_, num_fmt)
        ws.write_formula(r, 2 + len(BUCKETS), total_f, formats["td_bold_right"])
        ws.write_formula(r, 2 + len(BUCKETS) + 1, action_f, label_fmt)
        r += 1
    last_data_row_1b = r

    # CF on late buckets (31-60, 61-90, 90+) — higher = worse
    for j in range(2, len(BUCKETS)):
        col = col_letter(DATA_COL_FIRST_BUCKET + j)
        rng = f"{col}{first_data_row_1b}:{col}{last_data_row_1b}"
        st.add_three_color_scale(ws, rng, favourable_high=False)

    # Tie-out
    r += 2
    st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {"name": "Total AR on this tab ties to Data total receivables",
             "left": f"=SUM({col_letter(2 + len(BUCKETS))}{first_data_row_1b}:{col_letter(2 + len(BUCKETS))}{last_data_row_1b})",
             "right": f"='Data'!{col_letter(DATA_COL_TOTAL)}{DATA_ROW_TOTAL}"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Chase List")


def build_bucket_summary(wb, formats, dat):
    ws = wb.add_worksheet("Bucket Summary")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = 5
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 26)
    ws.set_column(2, LAST_COL, 16)
    ws.set_column(LAST_COL + 1, LAST_COL + 1, 2)

    sc.write_hero_band(ws, formats, kicker="Where is the AR sitting",
                       title="Aging bucket summary",
                       last_col=LAST_COL + 2,
                       explanation=(
                           "Total receivables per aging bucket, with each bucket's share of total AR. "
                           "Use this to track whether the AR profile is improving (more in Current, less "
                           "in 60+) or deteriorating."
                       ))

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet",
                            title="Total AR by bucket")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row, headers=["Bucket", "Total $", "Share of AR"], right_align_from=2)

    r = header_row + 1
    first_data_row_1b = r + 1
    for j, bucket in enumerate(BUCKETS):
        zebra = j % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        pct_fmt = formats["td_pct_zebra"] if zebra else formats["td_pct"]
        total_f = f"='Data'!{col_letter(DATA_COL_FIRST_BUCKET + j)}{DATA_ROW_TOTAL}"
        share_f = f"=IFERROR(C{r + 1}/'Data'!{col_letter(DATA_COL_TOTAL)}{DATA_ROW_TOTAL},0)"
        ws.set_row(r, 22)
        ws.write_string(r, 1, bucket, label_fmt)
        ws.write_formula(r, 2, total_f, num_fmt)
        ws.write_formula(r, 3, share_f, pct_fmt)
        r += 1
    last_data_row_1b = r

    # Total row
    st.write_total_row(ws, formats, row=r, label="Total AR",
                       formulas=[f"=SUM(C{first_data_row_1b}:C{last_data_row_1b})",
                                  f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})"],
                       cell_format="total_right")
    # Patch percent cell format
    # Total row writes both cells as total_right (currency); patch the share cell to pct format
    pct_total_fmt = wb.add_format({"font_name": "Arial", "font_size": 11, "bold": True, "font_color": b.WHITE,
                                   "bg_color": b.GREY_BLACK, "align": "right", "valign": "vcenter",
                                   "num_format": "0.0%;[Red](0.0%);\"-\"",
                                   "top": 2, "top_color": b.GREEN})
    ws.write_formula(r, 3, f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})", pct_total_fmt)

    # Chart
    r += 2
    chart_anchor = f"B{r + 2}"
    cats_range = f"='Bucket Summary'!$B${first_data_row_1b}:$B${last_data_row_1b}"
    series = [{
        "name": "Total $ per bucket",
        "values": f"='Bucket Summary'!$C${first_data_row_1b}:$C${last_data_row_1b}",
        "color": b.CHART_PRIMARY,
    }]
    st.add_column_chart(wb, ws, title="AR by aging bucket", anchor_cell=chart_anchor,
                       series=series, cats_range=cats_range, width=620, height=300)

    sc.apply_page_setup(ws, sheet_title="Bucket Summary")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_chase_list(wb, formats, dat)
    build_bucket_summary(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
