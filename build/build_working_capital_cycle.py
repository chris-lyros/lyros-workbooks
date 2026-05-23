"""Working Capital Cycle workbook (DSO, DPO, DIO, CCC).

Standalone version of the working-capital section that lives inside the
Management Reporting Pack. Adds a benchmark-band sheet so users see how
their cycle compares to typical SME ranges.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\5150 - Working Capital Cycle.xlsx")

WORKBOOK_ID = "5150"
WORKBOOK_TITLE = "Working Capital Cycle"
WORKBOOK_KICKER = "Debtor days, creditor days, inventory days, and cash conversion cycle"
TARGET_USER = (
    "CFO benchmarking working-capital health quarter on quarter, or finance lead "
    "spotting deterioration in receivables or payables timing."
)
HOW_TO_USE = [
    "Open the Data sheet and enter monthly DSO, DPO, and DIO for the last 12 months.",
    "The Cycle sheet computes CCC (DSO + DIO - DPO) per month, shows a trend chart, and benchmarks the cycle against typical SME ranges.",
    "The Benchmark sheet lists the typical range per industry; adjust the band that fits your business.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "Wholesale and retail SME"),
    ("TYPICAL DSO", "30-45 days"),
    ("TYPICAL DPO", "40-55 days"),
    ("TYPICAL DIO", "25-45 days"),
    ("CYCLE TARGET", "Under 30 days CCC"),
]

INPUTS_REQUIRED = [
    ("Days Sales Outstanding (DSO) per month", "Data tab"),
    ("Days Payables Outstanding (DPO) per month", "Data tab"),
    ("Days Inventory Outstanding (DIO) per month", "Data tab"),
    ("Benchmark targets (DSO, DPO, DIO, CCC)", "Data tab"),
]

NUM_MONTHS = 12
DATA_COL_LABEL = 1
DATA_COL_M1 = 2
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1
DATA_COL_AVG = DATA_COL_M_LAST + 1

DATA_ROW_HEADER = 8
DATA_ROW_DSO = 9
DATA_ROW_DPO = 10
DATA_ROW_DIO = 11
DATA_ROW_BENCH_HEADER = 14
DATA_ROW_BENCH_FIRST = 15
DATA_ROW_BENCH_LAST = 18


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
    rng_dso = d.make_rng("wcc-dso")
    rng_dpo = d.make_rng("wcc-dpo")
    rng_dio = d.make_rng("wcc-dio")
    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    dso = [round(38 + rng_dso.uniform(-5, 6), 0) for _ in range(NUM_MONTHS)]
    dpo = [round(46 + rng_dpo.uniform(-5, 5), 0) for _ in range(NUM_MONTHS)]
    dio = [round(28 + rng_dio.uniform(-3, 6), 0) for _ in range(NUM_MONTHS)]
    return {"months": months, "dso": dso, "dpo": dpo, "dio": dio}


def build_data_sheet(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)
    LAST_COL = DATA_COL_AVG + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 32)
    ws.set_column(DATA_COL_M1, DATA_COL_M_LAST, 11)
    ws.set_column(DATA_COL_AVG, DATA_COL_AVG, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Single source of truth",
                       title="Drop your working-capital days here",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "Calculate Days Sales Outstanding (DSO), Days Payables Outstanding (DPO), and "
                           "Days Inventory Outstanding (DIO) for each month, then enter them below. The "
                           "Cycle tab computes the Cash Conversion Cycle (CCC) = DSO + DIO - DPO and "
                           "benchmarks against the targets you set in the Benchmark table."
                       ))

    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_section_header(ws, formats, row=6, kicker="Step 1   Monthly days outstanding",
                            title="DSO, DPO, DIO inputs")
    st.write_header_row(ws, formats, row=DATA_ROW_HEADER - 1,
                        headers=["Metric"] + month_headers + ["Average"],
                        start_col=DATA_COL_LABEL, right_align_from=2)

    def write_row(row_1b, label, values):
        row_0 = row_1b - 1
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_LABEL, label, formats["input_label"])
        for m, v in enumerate(values):
            ws.write_number(row_0, DATA_COL_M1 + m, v, formats["input_days"])
        first_m = col_letter(DATA_COL_M1)
        last_m = col_letter(DATA_COL_M_LAST)
        ws.write_formula(row_0, DATA_COL_AVG,
                         f"=AVERAGE(${first_m}${row_1b}:${last_m}${row_1b})",
                         formats["td_days"])

    write_row(DATA_ROW_DSO, "Days Sales Outstanding (DSO)", dat["dso"])
    write_row(DATA_ROW_DPO, "Days Payables Outstanding (DPO)", dat["dpo"])
    write_row(DATA_ROW_DIO, "Days Inventory Outstanding (DIO)", dat["dio"])

    # Benchmark table
    st.write_section_header(ws, formats, row=DATA_ROW_BENCH_HEADER - 2,
                            kicker="Step 2   Benchmark targets", title="Industry typical ranges")
    st.write_header_row(ws, formats, row=DATA_ROW_BENCH_HEADER - 1,
                        headers=["Metric", "Best in class", "Typical", "Below typical"],
                        start_col=DATA_COL_LABEL, right_align_from=2)
    bench_rows = [
        ("DSO target (lower is better)", 30, 45, 60),
        ("DPO target (higher is better)", 60, 45, 30),
        ("DIO target (lower is better)", 20, 35, 50),
        ("CCC target (lower is better)", 0, 30, 60),
    ]
    for i, (label, best, typical, below) in enumerate(bench_rows):
        row_0 = DATA_ROW_BENCH_FIRST - 1 + i
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_LABEL, label, formats["input_label"])
        ws.write_number(row_0, DATA_COL_M1, best, formats["input_days"])
        ws.write_number(row_0, DATA_COL_M1 + 1, typical, formats["input_days"])
        ws.write_number(row_0, DATA_COL_M1 + 2, below, formats["input_days"])

    sc.apply_page_setup(ws, sheet_title="Data")


def build_cycle(wb, formats, dat):
    ws = wb.add_worksheet("Cycle")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = DATA_COL_AVG + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 32)
    ws.set_column(DATA_COL_M1, DATA_COL_M_LAST, 11)
    ws.set_column(DATA_COL_AVG, DATA_COL_AVG, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="DSO + DIO - DPO",
                       title="Cash conversion cycle by month",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "DSO is the average days from invoice to collection. DPO is the average "
                           "days from supplier invoice to payment. DIO is the average days inventory "
                           "sits before sale. CCC = DSO + DIO - DPO; lower is better. Negative CCC "
                           "means suppliers fund your inventory before customers pay you."
                       ))

    months = dat["months"]
    month_headers = [m.strftime("%b %y") for m in months]

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet",
                            title="Days outstanding and cycle")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Metric"] + month_headers + ["Average"], right_align_from=2)

    rows_def = [
        ("Days Sales Outstanding (DSO)", DATA_ROW_DSO, False),
        ("Days Payables Outstanding (DPO)", DATA_ROW_DPO, True),
        ("Days Inventory Outstanding (DIO)", DATA_ROW_DIO, False),
    ]
    r = header_row + 1
    for label, drow, zebra in rows_def:
        formulas = [f"='Data'!{col_letter(DATA_COL_M1 + m)}{drow}" for m in range(NUM_MONTHS)]
        avg_f = f"=AVERAGE('Data'!{col_letter(DATA_COL_M1)}{drow}:{col_letter(DATA_COL_M_LAST)}{drow})"
        st.write_data_row(ws, formats, row=r, label=label, formulas=formulas + [avg_f], zebra=zebra, cell_format="td_days")
        r += 1

    # CCC
    ccc_formulas = [
        f"='Data'!{col_letter(DATA_COL_M1 + m)}{DATA_ROW_DSO}"
        f"+'Data'!{col_letter(DATA_COL_M1 + m)}{DATA_ROW_DIO}"
        f"-'Data'!{col_letter(DATA_COL_M1 + m)}{DATA_ROW_DPO}"
        for m in range(NUM_MONTHS)
    ]
    ccc_avg = f"=AVERAGE({col_letter(2)}{r + 1}:{col_letter(2 + NUM_MONTHS - 1)}{r + 1})"
    st.write_total_row(ws, formats, row=r, label="Cash Conversion Cycle (CCC)",
                       formulas=ccc_formulas + [ccc_avg], cell_format="total_days")
    ccc_row_1b = r + 1
    r += 2

    # CF: high DSO/DIO = bad; high DPO = good
    st.add_three_color_scale(ws, f"C{header_row + 2}:N{header_row + 2}", favourable_high=False)  # DSO
    st.add_three_color_scale(ws, f"C{header_row + 3}:N{header_row + 3}", favourable_high=True)   # DPO
    st.add_three_color_scale(ws, f"C{header_row + 4}:N{header_row + 4}", favourable_high=False)  # DIO

    st.write_cf_legend(ws, formats, row=r, col=1, favourable_high=False, metric_label="DSO and DIO (lower is better)")
    st.write_cf_legend(ws, formats, row=r + 1, col=1, favourable_high=True, metric_label="DPO (higher is better)")
    r += 3

    # Tie-out checks
    checks_end = st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {"name": "Latest month CCC equals DSO + DIO - DPO",
             "left": f"=N{ccc_row_1b}",
             "right": f"=N{header_row + 2}+N{header_row + 4}-N{header_row + 3}"},
            {"name": "Average CCC ties to mean of monthly CCC values",
             "left": f"=O{ccc_row_1b}",
             "right": f"=AVERAGE(C{ccc_row_1b}:N{ccc_row_1b})"},
        ],
    )

    # CCC trend chart
    chart_anchor = f"B{checks_end + 2}"
    cats_range = f"='Cycle'!${col_letter(2)}${header_row + 1}:${col_letter(2 + NUM_MONTHS - 1)}${header_row + 1}"
    series = [{
        "name": "Cash conversion cycle",
        "values": f"='Cycle'!${col_letter(2)}${ccc_row_1b}:${col_letter(2 + NUM_MONTHS - 1)}${ccc_row_1b}",
        "color": b.CHART_PRIMARY,
    }]
    st.add_line_chart(wb, ws, title="CCC trend", anchor_cell=chart_anchor,
                      series=series, cats_range=cats_range, width=720, height=300)

    sc.apply_page_setup(ws, sheet_title="Cycle")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_cycle(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
