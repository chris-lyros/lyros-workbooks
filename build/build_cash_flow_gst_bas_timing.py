"""GST and BAS Cash Timing workbook.

Models the BAS-quarter cash impact for the SME: GST collected on sales, GST
paid on purchases, PAYG withholding, PAYG instalments, and the net cash-out
that hits the bank on the BAS lodgement date.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\5120 - GST and BAS Cash Flow Timing.xlsx")

WORKBOOK_ID = "5120"
WORKBOOK_TITLE = "GST and BAS Cash Timing"
WORKBOOK_KICKER = "BAS quarter cash impact model"
TARGET_USER = (
    "SME Finance Controller, fractional CFO, or business owner avoiding cash "
    "surprises at quarterly BAS lodgement."
)
HOW_TO_USE = [
    "Open the Data sheet and enter monthly GST collected, GST paid, PAYG withheld, and PAYG instalments for the past 12 months.",
    "Each row shows the underlying activity; the workbook nets these by BAS quarter and reports the lodgement-day cash outflow.",
    "The Quarterly Summary tab shows the four BAS quarters with cash impact and due dates.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "GST-registered SME, quarterly BAS lodger"),
    ("PAYG OBLIGATIONS", "PAYG withholding (employees) and PAYG instalments"),
    ("OUTPUT", "Cash impact per BAS quarter with lodgement due date"),
]
INPUTS_REQUIRED = [
    ("Monthly GST collected (on sales)", "Data tab"),
    ("Monthly GST paid (on purchases)", "Data tab"),
    ("Monthly PAYG withholding", "Data tab"),
    ("Quarterly PAYG instalment amount", "Data tab"),
]

NUM_MONTHS = 12
DATA_ROW_HEADER = 8
DATA_ROW_GST_OUT = 9       # GST collected (on sales)
DATA_ROW_GST_IN = 10       # GST paid (on purchases)
DATA_ROW_PAYG_W = 11       # PAYG withheld
DATA_ROW_PAYG_I = 12       # PAYG instalment (typically quarterly; show monthly contribution)
DATA_COL_LABEL = 1
DATA_COL_M1 = 2
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1
DATA_COL_TOTAL = DATA_COL_M_LAST + 1


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
    rng_g = d.make_rng("gst")
    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    # Monthly revenue ~330k → GST collected ~30k; purchases ~70k → GST paid ~7k
    gst_collected = [round(28_000 + rng_g.uniform(-3_000, 3_000), 0) for _ in range(NUM_MONTHS)]
    gst_paid = [round(6_500 + rng_g.uniform(-1_500, 1_500), 0) for _ in range(NUM_MONTHS)]
    payg_withheld = [round(14_000 + rng_g.uniform(-1_500, 1_500), 0) for _ in range(NUM_MONTHS)]
    payg_inst = [round(8_000, 0) if m % 3 == 0 else 0 for m in range(NUM_MONTHS)]
    return {"months": months, "gst_out": gst_collected, "gst_in": gst_paid,
            "payg_w": payg_withheld, "payg_i": payg_inst}


def build_data_sheet(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)
    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 32)
    ws.set_column(DATA_COL_M1, DATA_COL_M_LAST, 10)
    ws.set_column(DATA_COL_TOTAL, DATA_COL_TOTAL, 12)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Single source of truth",
                       title="Drop your GST and PAYG monthly activity here",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "GST collected is the GST amount on sales for each month. GST paid is the "
                           "GST amount on creditable purchases. PAYG withholding is the tax withheld "
                           "from wages. PAYG instalment is the quarterly company income tax instalment "
                           "(enter the full amount in the month it falls due). All four feed the BAS "
                           "lodgement cash impact on the Quarterly Summary tab."
                       ))

    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_section_header(ws, formats, row=6, kicker="Enter monthly figures", title="GST, PAYG withholding, PAYG instalments")
    st.write_header_row(ws, formats, row=DATA_ROW_HEADER - 1,
                        headers=["Line"] + month_headers + ["12-month total"],
                        start_col=DATA_COL_LABEL, right_align_from=2)

    def write_row(row_1b, label, values):
        row_0 = row_1b - 1
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_LABEL, label, formats["input_label"])
        for m, v in enumerate(values):
            ws.write_number(row_0, DATA_COL_M1 + m, v, formats["input_value"])
        first_m = col_letter(DATA_COL_M1)
        last_m = col_letter(DATA_COL_M_LAST)
        ws.write_formula(row_0, DATA_COL_TOTAL,
                         f"=SUM(${first_m}${row_1b}:${last_m}${row_1b})",
                         formats["td_bold_right"])

    write_row(DATA_ROW_GST_OUT, "GST collected (on sales)", dat["gst_out"])
    write_row(DATA_ROW_GST_IN, "GST paid (on purchases)", dat["gst_in"])
    write_row(DATA_ROW_PAYG_W, "PAYG withholding", dat["payg_w"])
    write_row(DATA_ROW_PAYG_I, "PAYG instalment (quarterly)", dat["payg_i"])

    sc.apply_page_setup(ws, sheet_title="Data")


def build_summary(wb, formats, dat):
    ws = wb.add_worksheet("Quarterly Summary")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = 8
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 24)
    ws.set_column(2, LAST_COL, 14)
    ws.set_column(LAST_COL + 1, LAST_COL + 1, 2)

    sc.write_hero_band(ws, formats, kicker="Quarterly summary",
                       title="BAS-quarter cash impact",
                       last_col=LAST_COL + 2,
                       explanation=(
                           "Aggregates the monthly GST and PAYG figures into the four BAS quarters and "
                           "calculates the net cash outflow per lodgement. The ATO lodgement due date "
                           "for each quarter is shown in the Due column. The BAS Net Cash Out row is "
                           "what hits the bank on lodgement day."
                       ))

    months = dat["months"]
    # Group months into quarters (3-month buckets, last 12 months → 4 quarters)
    quarters = [(0, 1, 2), (3, 4, 5), (6, 7, 8), (9, 10, 11)]
    quarter_labels = []
    for start, _, end in quarters:
        quarter_labels.append(f"{months[start].strftime('%b %y')} – {months[end].strftime('%b %y')}")

    # Lodgement due dates (ATO: 28 days after quarter end, except Dec quarter = 28 Feb).
    # We just show "approx" — the user adjusts. Use end-of-quarter + 28 days.
    due_dates = []
    for _, _, end in quarters:
        end_m = months[end]
        # Approximate end of month
        next_m = date(end_m.year + (1 if end_m.month == 12 else 0), 1 if end_m.month == 12 else end_m.month + 1, 1)
        eom = next_m - timedelta(days=1)
        due_dates.append(eom + timedelta(days=28))

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet",
                            title="Net cash impact per BAS quarter")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row, headers=["Line"] + quarter_labels + ["Total"],
                        right_align_from=2)

    def quarter_formula(row_1b: int, q_idx: int) -> str:
        start_m, _, end_m = quarters[q_idx]
        cells = [f"'Data'!${col_letter(DATA_COL_M1 + start_m + k)}${row_1b}" for k in range(3)]
        return f"={'+'.join(cells)}"

    r = header_row + 1
    # GST collected
    gst_out_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "GST collected", formats["td"])
    for q in range(4):
        ws.write_formula(r, 2 + q, quarter_formula(DATA_ROW_GST_OUT, q), formats["td_right"])
    ws.write_formula(r, 6, f"=SUM(C{r + 1}:F{r + 1})", formats["td_bold_right"])
    r += 1

    # GST paid (negative impact)
    gst_in_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "GST paid", formats["td_zebra"])
    for q in range(4):
        ws.write_formula(r, 2 + q, quarter_formula(DATA_ROW_GST_IN, q), formats["td_right_zebra"])
    ws.write_formula(r, 6, f"=SUM(C{r + 1}:F{r + 1})", formats["td_bold_right"])
    r += 1

    # Net GST
    net_gst_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "Net GST payable", formats["td_bold_left"])
    for q in range(4):
        ws.write_formula(r, 2 + q, f"={col_letter(2 + q)}{gst_out_row_1b}-{col_letter(2 + q)}{gst_in_row_1b}", formats["td_bold_right"])
    ws.write_formula(r, 6, f"=SUM(C{r + 1}:F{r + 1})", formats["td_bold_right"])
    r += 1

    # PAYG withholding
    payg_w_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "PAYG withholding", formats["td_zebra"])
    for q in range(4):
        ws.write_formula(r, 2 + q, quarter_formula(DATA_ROW_PAYG_W, q), formats["td_right_zebra"])
    ws.write_formula(r, 6, f"=SUM(C{r + 1}:F{r + 1})", formats["td_bold_right"])
    r += 1

    # PAYG instalment
    payg_i_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "PAYG instalment", formats["td"])
    for q in range(4):
        ws.write_formula(r, 2 + q, quarter_formula(DATA_ROW_PAYG_I, q), formats["td_right"])
    ws.write_formula(r, 6, f"=SUM(C{r + 1}:F{r + 1})", formats["td_bold_right"])
    r += 1

    # Total BAS net cash out
    bas_total_row_1b = r + 1
    ws.set_row(r, 24)
    ws.write_string(r, 1, "BAS net cash out", formats["total_left"])
    for q in range(4):
        ws.write_formula(r, 2 + q,
                         f"={col_letter(2 + q)}{net_gst_row_1b}+{col_letter(2 + q)}{payg_w_row_1b}+{col_letter(2 + q)}{payg_i_row_1b}",
                         formats["total_right"])
    ws.write_formula(r, 6, f"=SUM(C{r + 1}:F{r + 1})", formats["total_right"])
    r += 2

    # Due dates row
    ws.set_row(r, 22)
    ws.write_string(r, 1, "Approximate lodgement due", formats["td_bold_left"])
    date_fmt = wb.add_format({"font_name": "Arial", "font_size": 10, "align": "right", "valign": "vcenter",
                              "border": 1, "border_color": b.BORDER_GREY, "num_format": "yyyy-mm-dd"})
    for q in range(4):
        ws.write_datetime(r, 2 + q, due_dates[q], date_fmt)
    r += 2

    # Tie-out checks
    st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {"name": "GST collected total ties to Data row total",
             "left": f"={col_letter(2 + 4)}{gst_out_row_1b}",
             "right": f"='Data'!{col_letter(DATA_COL_TOTAL)}{DATA_ROW_GST_OUT}"},
            {"name": "PAYG instalment total ties to Data row total",
             "left": f"={col_letter(2 + 4)}{payg_i_row_1b}",
             "right": f"='Data'!{col_letter(DATA_COL_TOTAL)}{DATA_ROW_PAYG_I}"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Quarterly Summary")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_summary(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
