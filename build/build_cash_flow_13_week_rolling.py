"""Build the 13-Week Rolling Cash Flow workbook.

Data shape: weekly receipts and payments by category for 13 weeks ahead,
plus an opening cash balance. Closing balances roll forward weekly.

Output:
  C:\\dev\\lyros-workbooks\\library\\cash_flow
        \\lyros_lib_cash_flow_13_week_rolling.xlsx
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\5100 - 13 Week Rolling Cash Flow.xlsx")

WORKBOOK_ID = "5100"
WORKBOOK_TITLE = "13-Week Rolling Cash Flow"
WORKBOOK_KICKER = "Weekly cash forecast across 13 weeks"
TARGET_USER = (
    "Founder-CEO, Finance Controller, or fractional CFO needing weekly cash "
    "visibility for the next quarter."
)
HOW_TO_USE = [
    "Open the Data sheet and enter the opening cash balance, expected weekly receipts, and expected weekly payments.",
    "The Forecast sheet shows the closing balance per week with a minimum-balance flag and a trend chart.",
    "Each week's closing balance carries forward as the next week's opening balance automatically.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "SME with regular weekly receipts and payments"),
    ("OPENING CASH", "Circa $250k"),
    ("FORECAST HORIZON", "13 weeks forward from the build date"),
    ("MIN BALANCE TRIGGER", "$100k (used by the alert column)"),
]

INPUTS_REQUIRED = [
    ("Opening cash balance", "Data tab"),
    ("Weekly expected receipts by category", "Data tab"),
    ("Weekly expected payments by category", "Data tab"),
    ("Minimum balance trigger", "Data tab"),
]

NUM_WEEKS = 13
RECEIPT_CATEGORIES = [
    ("Customer receipts - operating", "operating"),
    ("Customer receipts - large invoices", "lumpy"),
    ("Other receipts", "other"),
]
PAYMENT_CATEGORIES = [
    ("Payroll", "payroll"),
    ("Suppliers - operating", "suppliers"),
    ("Rent and utilities", "fixed"),
    ("Marketing and other", "discretionary"),
    ("BAS and PAYG", "tax"),
    ("Loan repayments", "debt"),
]

DATA_COL_LABEL = 1
DATA_COL_W1 = 2
DATA_COL_W_LAST = DATA_COL_W1 + NUM_WEEKS - 1
DATA_COL_TOTAL = DATA_COL_W_LAST + 1

DATA_ROW_OPENING = 7
DATA_ROW_RECEIPTS_HEADER = 10
DATA_ROW_RECEIPTS_FIRST = 11
DATA_ROW_RECEIPTS_LAST = DATA_ROW_RECEIPTS_FIRST + len(RECEIPT_CATEGORIES) - 1
DATA_ROW_PAYMENTS_HEADER = DATA_ROW_RECEIPTS_LAST + 3
DATA_ROW_PAYMENTS_FIRST = DATA_ROW_PAYMENTS_HEADER + 1
DATA_ROW_PAYMENTS_LAST = DATA_ROW_PAYMENTS_FIRST + len(PAYMENT_CATEGORIES) - 1
DATA_ROW_MIN_BALANCE = DATA_ROW_PAYMENTS_LAST + 3


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
    rng = d.make_rng("13wk")
    start = date.today() - timedelta(days=date.today().weekday())
    weeks = [start + timedelta(weeks=i) for i in range(NUM_WEEKS)]

    # Receipts by week and category
    receipts = []
    for cat_name, kind in RECEIPT_CATEGORIES:
        if kind == "operating":
            row = [round(72_000 + rng.uniform(-6_000, 6_000), 0) for _ in range(NUM_WEEKS)]
        elif kind == "lumpy":
            row = [0] * NUM_WEEKS
            for w in [2, 6, 10]:
                row[w] = round(45_000 + rng.uniform(-5_000, 5_000), 0)
        else:
            row = [round(4_000 + rng.uniform(-1_500, 1_500), 0) for _ in range(NUM_WEEKS)]
        receipts.append(row)

    # Payments by week and category
    payments = []
    for cat_name, kind in PAYMENT_CATEGORIES:
        if kind == "payroll":
            row = [round(26_000 + rng.uniform(-1_500, 1_500), 0) if w % 2 == 0 else 0 for w in range(NUM_WEEKS)]
        elif kind == "suppliers":
            row = [round(36_000 + rng.uniform(-5_000, 5_000), 0) for _ in range(NUM_WEEKS)]
        elif kind == "fixed":
            row = [round(7_500, 0) for _ in range(NUM_WEEKS)]
        elif kind == "discretionary":
            row = [round(5_500 + rng.uniform(-1_500, 1_500), 0) for _ in range(NUM_WEEKS)]
        elif kind == "tax":
            row = [0] * NUM_WEEKS
            row[3] = 28_000  # quarterly BAS in week 4
            row[10] = 18_000  # PAYG instalment
        elif kind == "debt":
            row = [round(3_200, 0) if w % 4 == 0 else 0 for w in range(NUM_WEEKS)]
        payments.append(row)

    return {"weeks": weeks, "receipts": receipts, "payments": payments,
            "opening_cash": 250_000, "min_balance": 100_000}


def build_data_sheet(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 32)
    ws.set_column(DATA_COL_W1, DATA_COL_W_LAST, 10)
    ws.set_column(DATA_COL_TOTAL, DATA_COL_TOTAL, 12)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Single source of truth",
        title="Drop your 13-week cash inputs here",
        last_col=LAST_COL + 1,
        explanation=(
            "Enter the opening cash balance, weekly expected receipts by category, and "
            "weekly expected payments by category. The Forecast tab rolls the closing "
            "balance forward week by week and flags any week where the projected closing "
            "balance dips below the minimum balance trigger."
        ),
    )

    # Opening cash input
    ws.set_row(DATA_ROW_OPENING - 1, 24)
    ws.write_string(DATA_ROW_OPENING - 1, DATA_COL_LABEL, "Opening cash balance (week 1)", formats["input_label"])
    ws.write_number(DATA_ROW_OPENING - 1, DATA_COL_W1, dat["opening_cash"], formats["input_value"])

    # Week headers
    week_headers = [f"W{i + 1}  {dat['weeks'][i].strftime('%d %b')}" for i in range(NUM_WEEKS)]

    # Receipts table
    st.write_section_header(ws, formats, row=DATA_ROW_RECEIPTS_HEADER - 2,
                            kicker="Step 1   Weekly expected receipts", title="Receipts")
    st.write_header_row(ws, formats, row=DATA_ROW_RECEIPTS_HEADER - 1,
                        headers=["Category"] + week_headers + ["13-wk total"],
                        start_col=DATA_COL_LABEL, right_align_from=2)

    for i, (cat_name, _kind) in enumerate(RECEIPT_CATEGORIES):
        row_0 = DATA_ROW_RECEIPTS_FIRST - 1 + i
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_LABEL, cat_name, formats["input_label"])
        for w, v in enumerate(dat["receipts"][i]):
            ws.write_number(row_0, DATA_COL_W1 + w, v, formats["input_value"])
        first_w = col_letter(DATA_COL_W1)
        last_w = col_letter(DATA_COL_W_LAST)
        ws.write_formula(row_0, DATA_COL_TOTAL,
                         f"=SUM(${first_w}${DATA_ROW_RECEIPTS_FIRST + i}:${last_w}${DATA_ROW_RECEIPTS_FIRST + i})",
                         formats["td_bold_right"])

    # Payments table
    st.write_section_header(ws, formats, row=DATA_ROW_PAYMENTS_HEADER - 2,
                            kicker="Step 2   Weekly expected payments", title="Payments")
    st.write_header_row(ws, formats, row=DATA_ROW_PAYMENTS_HEADER - 1,
                        headers=["Category"] + week_headers + ["13-wk total"],
                        start_col=DATA_COL_LABEL, right_align_from=2)

    for i, (cat_name, _kind) in enumerate(PAYMENT_CATEGORIES):
        row_0 = DATA_ROW_PAYMENTS_FIRST - 1 + i
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_LABEL, cat_name, formats["input_label"])
        for w, v in enumerate(dat["payments"][i]):
            ws.write_number(row_0, DATA_COL_W1 + w, v, formats["input_value"])
        first_w = col_letter(DATA_COL_W1)
        last_w = col_letter(DATA_COL_W_LAST)
        ws.write_formula(row_0, DATA_COL_TOTAL,
                         f"=SUM(${first_w}${DATA_ROW_PAYMENTS_FIRST + i}:${last_w}${DATA_ROW_PAYMENTS_FIRST + i})",
                         formats["td_bold_right"])

    # Minimum balance trigger
    ws.set_row(DATA_ROW_MIN_BALANCE - 1, 24)
    ws.write_string(DATA_ROW_MIN_BALANCE - 1, DATA_COL_LABEL, "Minimum balance trigger", formats["input_label"])
    ws.write_number(DATA_ROW_MIN_BALANCE - 1, DATA_COL_W1, dat["min_balance"], formats["input_value"])

    sc.apply_page_setup(ws, sheet_title="Data")
    return ws


def build_forecast(wb, formats, dat):
    ws = wb.add_worksheet("Forecast")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = DATA_COL_TOTAL + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 28)
    ws.set_column(DATA_COL_W1, DATA_COL_W_LAST, 10)
    ws.set_column(DATA_COL_TOTAL, DATA_COL_TOTAL, 12)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="13 weeks forward",
        title="Rolling cash forecast",
        last_col=LAST_COL + 1,
        explanation=(
            "Opening balance plus receipts minus payments equals the closing balance for "
            "each week. Each week's closing balance becomes the next week's opening balance. "
            "The Below trigger row flags any week where the projected closing balance falls "
            "under the minimum balance set on the Data tab. BAS is Business Activity Statement; "
            "PAYG is Pay As You Go withholding."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet",
                            title="Weekly cash position")

    header_row = section_row + 2
    week_headers = [f"W{i + 1}  {dat['weeks'][i].strftime('%d %b')}" for i in range(NUM_WEEKS)]
    st.write_header_row(ws, formats, row=header_row, headers=["Line"] + week_headers + ["13-wk total"],
                        right_align_from=2)

    r = header_row + 1
    # Opening balance row
    opening_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "Opening balance", formats["td_bold_left"])
    # Week 1 opening = Data opening
    ws.write_formula(r, 2, f"='Data'!${col_letter(DATA_COL_W1)}${DATA_ROW_OPENING}", formats["td_bold_right"])
    # Subsequent weeks: opening = prior week closing (cell two rows below, same column - 1)
    for w in range(1, NUM_WEEKS):
        ws.write_formula(r, 2 + w, f"={col_letter(2 + w - 1)}{r + 5}", formats["td_bold_right"])
    ws.write_blank(r, DATA_COL_TOTAL, None, formats["td_bold_right"])
    r += 1

    # Receipts (sum of receipts from Data)
    receipts_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "Receipts", formats["td"])
    for w in range(NUM_WEEKS):
        col = col_letter(DATA_COL_W1 + w)
        f_ = f"=SUM('Data'!${col}${DATA_ROW_RECEIPTS_FIRST}:${col}${DATA_ROW_RECEIPTS_LAST})"
        ws.write_formula(r, 2 + w, f_, formats["td_right"])
    ws.write_formula(r, DATA_COL_TOTAL, f"=SUM({col_letter(2)}{r + 1}:{col_letter(2 + NUM_WEEKS - 1)}{r + 1})",
                     formats["td_bold_right"])
    r += 1

    # Payments
    payments_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "Payments", formats["td_zebra"])
    for w in range(NUM_WEEKS):
        col = col_letter(DATA_COL_W1 + w)
        f_ = f"=-SUM('Data'!${col}${DATA_ROW_PAYMENTS_FIRST}:${col}${DATA_ROW_PAYMENTS_LAST})"
        ws.write_formula(r, 2 + w, f_, formats["td_right_zebra"])
    ws.write_formula(r, DATA_COL_TOTAL, f"=SUM({col_letter(2)}{r + 1}:{col_letter(2 + NUM_WEEKS - 1)}{r + 1})",
                     formats["td_bold_right"])
    r += 1

    # Net movement
    net_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "Net movement", formats["td_bold_left"])
    for w in range(NUM_WEEKS):
        ws.write_formula(r, 2 + w,
                         f"={col_letter(2 + w)}{receipts_row_1b}+{col_letter(2 + w)}{payments_row_1b}",
                         formats["td_bold_right"])
    ws.write_formula(r, DATA_COL_TOTAL,
                     f"={col_letter(2 + NUM_WEEKS)}{receipts_row_1b}+{col_letter(2 + NUM_WEEKS)}{payments_row_1b}",
                     formats["td_bold_right"])
    r += 1

    # Closing balance
    closing_row_1b = r + 1
    ws.set_row(r, 24)
    ws.write_string(r, 1, "Closing balance", formats["total_left"])
    for w in range(NUM_WEEKS):
        ws.write_formula(r, 2 + w,
                         f"={col_letter(2 + w)}{opening_row_1b}+{col_letter(2 + w)}{net_row_1b}",
                         formats["total_right"])
    ws.write_formula(r, DATA_COL_TOTAL, f"={col_letter(2 + NUM_WEEKS - 1)}{r + 1}", formats["total_right"])
    r += 1

    # Below trigger flag
    trigger_row_1b = r + 1
    ws.set_row(r, 22)
    ws.write_string(r, 1, "Below trigger", formats["td_bold_left"])
    for w in range(NUM_WEEKS):
        ws.write_formula(r, 2 + w,
                         f"=IF({col_letter(2 + w)}{closing_row_1b}<'Data'!${col_letter(DATA_COL_W1)}${DATA_ROW_MIN_BALANCE},\"FLAG\",\"OK\")",
                         formats["check_status_neutral"])
    ws.write_blank(r, DATA_COL_TOTAL, None, formats["td"])
    # CF for OK/FLAG
    status_range = f"{col_letter(2)}{r + 1}:{col_letter(2 + NUM_WEEKS - 1)}{r + 1}"
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["check_status_ok"]})
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "FLAG", "format": formats["check_status_flag"]})
    r += 2

    # Tie-out checks
    checks_end_one = st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {
                "name": "Closing balance W13 equals opening + total net movement",
                "left": f"={col_letter(2 + NUM_WEEKS - 1)}{closing_row_1b}",
                "right": f"={col_letter(2)}{opening_row_1b}+{col_letter(2 + NUM_WEEKS)}{net_row_1b}",
            },
            {
                "name": "Total receipts equals sum of receipt categories on Data",
                "left": f"={col_letter(2 + NUM_WEEKS)}{receipts_row_1b}",
                "right": f"=SUM('Data'!{col_letter(DATA_COL_TOTAL)}{DATA_ROW_RECEIPTS_FIRST}:{col_letter(DATA_COL_TOTAL)}{DATA_ROW_RECEIPTS_LAST})",
            },
        ],
    )

    # Closing balance trend chart
    chart_anchor = f"B{checks_end_one + 2}"
    cats_range = f"='Forecast'!${col_letter(2)}${header_row + 1}:${col_letter(2 + NUM_WEEKS - 1)}${header_row + 1}"
    series = [{
        "name": "Closing balance",
        "values": f"='Forecast'!${col_letter(2)}${closing_row_1b}:${col_letter(2 + NUM_WEEKS - 1)}${closing_row_1b}",
        "color": b.CHART_PRIMARY,
    }]
    st.add_line_chart(wb, ws, title="Closing balance by week", anchor_cell=chart_anchor,
                      series=series, cats_range=cats_range, width=720, height=300)

    sc.apply_page_setup(ws, sheet_title="Forecast")
    return ws


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_forecast(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
