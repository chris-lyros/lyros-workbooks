"""Aged Receivables Analysis workbook.

Customer-level aging report from your accounting software. Adds heatmap, top-debtor exposure,
and provision suggestion logic based on aging buckets.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\5210 - Aged Receivables Analysis.xlsx")

WORKBOOK_ID = "5210"
WORKBOOK_TITLE = "Aged Receivables Analysis"
WORKBOOK_KICKER = "Aging buckets with provision suggestion"
TARGET_USER = (
    "Credit controller, Finance Controller, or external accountant running a "
    "monthly receivables review and computing a doubtful-debt provision."
)
HOW_TO_USE = [
    "Open the Data sheet and paste your Aged Receivables Summary export.",
    "Adjust the provision rate per bucket on the Settings block (default rates are conservative SME values).",
    "The Analysis sheet shows the heatmap, top-debtor exposure, and suggested provision; the Provision sheet breaks it down per customer.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "B2B SME with 15-20 active debtors"),
    ("TOTAL AR EXPOSURE", "Circa $480k at month-end"),
    ("PROVISION POLICY", "0% current, 2% 30 days, 10% 60 days, 25% 90 days, 75% 90+ days"),
]

INPUTS_REQUIRED = [
    ("Aged Receivables Summary by customer", "Data tab (paste from your accounting software AR Summary)"),
    ("Provision rate per aging bucket", "Data tab → Provision settings block"),
]

BUCKETS = ["Current", "1-30 days", "31-60 days", "61-90 days", "90+ days"]

DATA_COL_LABEL = 1
DATA_COL_FIRST_BUCKET = 2
DATA_COL_LAST_BUCKET = DATA_COL_FIRST_BUCKET + len(BUCKETS) - 1
DATA_COL_TOTAL = DATA_COL_LAST_BUCKET + 1

DATA_ROW_HEADER = 8
DATA_ROW_FIRST = 9
DATA_ROW_LAST = DATA_ROW_FIRST + 14  # 15 customer rows
DATA_ROW_TOTAL = DATA_ROW_LAST + 1
DATA_ROW_PROV_HEADER = DATA_ROW_TOTAL + 3
DATA_ROW_PROV_RATE = DATA_ROW_PROV_HEADER + 1


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
    rng = d.make_rng("ar-an")
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
    provision_rates = [0.0, 0.02, 0.10, 0.25, 0.75]
    return {"rows": rows, "provision_rates": provision_rates}


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
                       title="Drop your Aged Receivables Summary here",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "Paste your Aged Receivables Summary by customer into the table below. "
                           "The five aging buckets follow the standard Current / 1-30 / 31-60 / 61-90 / "
                           "90+ days breakdown. The Provision settings block below the table holds the "
                           "rate applied to each bucket when calculating the doubtful-debt provision."
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

    # Totals row
    ws.set_row(DATA_ROW_TOTAL - 1, 24)
    ws.write_string(DATA_ROW_TOTAL - 1, DATA_COL_LABEL, "Total receivables", formats["total_left"])
    for j in range(len(BUCKETS) + 1):
        col = col_letter(DATA_COL_FIRST_BUCKET + j)
        ws.write_formula(DATA_ROW_TOTAL - 1, DATA_COL_FIRST_BUCKET + j,
                         f"=SUM({col}{DATA_ROW_FIRST}:{col}{DATA_ROW_LAST})", formats["total_right"])

    # Provision settings
    st.write_section_header(ws, formats, row=DATA_ROW_PROV_HEADER - 2,
                            kicker="Provision settings", title="Doubtful-debt provision rate per bucket")
    ws.set_row(DATA_ROW_PROV_HEADER - 1, 24)
    ws.write_string(DATA_ROW_PROV_HEADER - 1, DATA_COL_LABEL, "Bucket", formats["th"])
    for j, bucket in enumerate(BUCKETS):
        ws.write_string(DATA_ROW_PROV_HEADER - 1, DATA_COL_FIRST_BUCKET + j, bucket, formats["th_right"])

    ws.set_row(DATA_ROW_PROV_RATE - 1, 22)
    ws.write_string(DATA_ROW_PROV_RATE - 1, DATA_COL_LABEL, "Provision rate", formats["input_label"])
    for j, rate in enumerate(dat["provision_rates"]):
        ws.write_number(DATA_ROW_PROV_RATE - 1, DATA_COL_FIRST_BUCKET + j, rate, formats["input_pct"])

    sc.apply_page_setup(ws, sheet_title="Data")


def build_analysis(wb, formats, dat):
    ws = wb.add_worksheet("Analysis")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 32)
    ws.set_column(DATA_COL_FIRST_BUCKET, DATA_COL_LAST_BUCKET, 14)
    ws.set_column(DATA_COL_TOTAL, DATA_COL_TOTAL, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Aging heatmap and exposure",
                       title="Aged receivables analysis",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "Reads the aging table from the Data sheet. Conditional formatting colours "
                           "each customer cell by exposure (higher is worse for the later buckets). "
                           "The Provision column shows the doubtful-debt provision per customer "
                           "calculated as Sum of (bucket × provision rate) using the rates on the Data tab."
                       ))

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet",
                            title="Customer exposure with suggested provision")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row, headers=["Customer"] + BUCKETS + ["Total", "Provision"],
                        right_align_from=2)

    # Need to extend column for Provision
    ws.set_column(DATA_COL_TOTAL + 1, DATA_COL_TOTAL + 1, 14)

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (name, _buckets) in enumerate(dat["rows"]):
        zebra = i % 2 == 0
        data_row_1b = DATA_ROW_FIRST + i
        formulas = [f"='Data'!{col_letter(DATA_COL_FIRST_BUCKET + j)}{data_row_1b}" for j in range(len(BUCKETS))]
        total_f = f"='Data'!{col_letter(DATA_COL_TOTAL)}{data_row_1b}"
        # Provision = SUMPRODUCT of buckets × rates
        prov_terms = "+".join(
            f"'Data'!{col_letter(DATA_COL_FIRST_BUCKET + j)}{data_row_1b}"
            f"*'Data'!{col_letter(DATA_COL_FIRST_BUCKET + j)}{DATA_ROW_PROV_RATE}"
            for j in range(len(BUCKETS))
        )
        prov_f = f"={prov_terms}"
        st.write_data_row(ws, formats, row=r, label=name, formulas=formulas + [total_f, prov_f], zebra=zebra, cell_format="td_right")
        r += 1
    last_data_row_1b = r

    # Totals row
    total_formulas = [f"=SUM({col_letter(DATA_COL_FIRST_BUCKET + j)}{first_data_row_1b}:{col_letter(DATA_COL_FIRST_BUCKET + j)}{last_data_row_1b})" for j in range(len(BUCKETS))]
    grand_total = f"=SUM({col_letter(DATA_COL_TOTAL)}{first_data_row_1b}:{col_letter(DATA_COL_TOTAL)}{last_data_row_1b})"
    grand_prov = f"=SUM({col_letter(DATA_COL_TOTAL + 1)}{first_data_row_1b}:{col_letter(DATA_COL_TOTAL + 1)}{last_data_row_1b})"
    st.write_total_row(ws, formats, row=r, label="Total", formulas=total_formulas + [grand_total, grand_prov])
    total_row_1b = r + 1
    r += 2

    # CF on the aging buckets: higher = worse (heatmap)
    for j in range(len(BUCKETS)):
        col = col_letter(DATA_COL_FIRST_BUCKET + j)
        rng = f"{col}{first_data_row_1b}:{col}{last_data_row_1b}"
        # Current bucket: lower = nothing to chase (favourable); but for heatmap we use unfavourable for late buckets
        st.add_three_color_scale(ws, rng, favourable_high=(j == 0))

    st.write_cf_legend(ws, formats, row=r, col=1, favourable_high=False, metric_label="Late buckets (lower is better)")
    r += 2

    # Tie-out
    checks_end = st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {"name": "Total AR ties to Data total receivables",
             "left": f"={col_letter(DATA_COL_TOTAL)}{total_row_1b}",
             "right": f"='Data'!{col_letter(DATA_COL_TOTAL)}{DATA_ROW_TOTAL}"},
            {"name": "Total provision equals sum of customer provisions",
             "left": f"={col_letter(DATA_COL_TOTAL + 1)}{total_row_1b}",
             "right": f"=SUM({col_letter(DATA_COL_TOTAL + 1)}{first_data_row_1b}:{col_letter(DATA_COL_TOTAL + 1)}{last_data_row_1b})"},
        ],
    )

    # Top debtor chart
    chart_anchor = f"B{checks_end + 2}"
    cats_range = f"='Analysis'!$B${first_data_row_1b}:$B${last_data_row_1b}"
    series = [{
        "name": "Total exposure",
        "values": f"='Analysis'!${col_letter(DATA_COL_TOTAL)}${first_data_row_1b}:${col_letter(DATA_COL_TOTAL)}${last_data_row_1b}",
        "color": b.CHART_PRIMARY,
    }]
    st.add_column_chart(wb, ws, title="Total AR exposure by customer", anchor_cell=chart_anchor,
                       series=series, cats_range=cats_range, width=720, height=320)

    sc.apply_page_setup(ws, sheet_title="Analysis")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_analysis(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
