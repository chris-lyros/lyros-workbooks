"""Cash Position by Bank Account workbook.

Daily-ish (monthly closing) cash by bank account with trend chart and
minimum-balance flags per account.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\5110 - Cash Position by Bank.xlsx")

WORKBOOK_ID = "5110"
WORKBOOK_TITLE = "Cash Position by Bank Account"
WORKBOOK_KICKER = "Monthly closing cash by bank account"
TARGET_USER = "SME with multiple bank accounts (operating, savings, trust) needing visibility on each."
HOW_TO_USE = [
    "Open the Data sheet and enter the monthly closing balance per bank account.",
    "Set the minimum-balance trigger for each account.",
    "The Position sheet shows the trend and flags any month a balance dipped below trigger.",
]
EXAMPLE_PROFILE = [
    ("ACCOUNTS", "Operating, Savings, Trust"),
    ("REPORTING CADENCE", "Month-end snapshots"),
    ("WHO USES", "Founder, FC, or external accountant"),
]
INPUTS_REQUIRED = [
    ("Monthly closing balance per bank account", "Data tab"),
    ("Minimum balance trigger per account", "Data tab"),
]

NUM_MONTHS = 12
BANK_ACCOUNTS = ["Operating", "Savings", "Trust"]

DATA_COL_LABEL = 1
DATA_COL_M1 = 2
DATA_COL_M_LAST = DATA_COL_M1 + NUM_MONTHS - 1
DATA_COL_TRIGGER = DATA_COL_M_LAST + 1

DATA_ROW_HEADER = 8
DATA_ROW_FIRST = 9
DATA_ROW_LAST = DATA_ROW_FIRST + len(BANK_ACCOUNTS) - 1
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
    rng_op = d.make_rng("cpb-op")
    rng_sv = d.make_rng("cpb-sv")
    rng_tr = d.make_rng("cpb-tr")
    months = d.months_back(d.LAST_REPORTED_MONTH, NUM_MONTHS)
    op = [round(220_000 + 6_000 * i + rng_op.uniform(-15_000, 15_000), 0) for i in range(NUM_MONTHS)]
    sv = [round(380_000 + 3_000 * i + rng_sv.uniform(-8_000, 8_000), 0) for i in range(NUM_MONTHS)]
    tr = [round(60_000 + 1_500 * i + rng_tr.uniform(-4_000, 4_000), 0) for i in range(NUM_MONTHS)]
    return {"months": months, "balances": [op, sv, tr], "triggers": [100_000, 200_000, 30_000]}


def build_data_sheet(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)
    LAST_COL = DATA_COL_TRIGGER + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 24)
    ws.set_column(DATA_COL_M1, DATA_COL_M_LAST, 11)
    ws.set_column(DATA_COL_TRIGGER, DATA_COL_TRIGGER, 16)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Single source of truth",
                       title="Drop your bank balances here", last_col=LAST_COL + 1,
                       explanation=(
                           "Enter the closing balance for each bank account at the end of each month. "
                           "Set a minimum-balance trigger per account; the Position sheet flags any "
                           "month the closing balance dropped below the trigger."
                       ))

    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_section_header(ws, formats, row=6, kicker="Paste from your bank report",
                            title="Monthly closing balances")
    st.write_header_row(ws, formats, row=DATA_ROW_HEADER - 1,
                        headers=["Bank account"] + month_headers + ["Min trigger"],
                        start_col=DATA_COL_LABEL, right_align_from=2)
    for i, acct in enumerate(BANK_ACCOUNTS):
        row_0 = DATA_ROW_FIRST - 1 + i
        ws.set_row(row_0, 22)
        ws.write_string(row_0, DATA_COL_LABEL, acct, formats["input_label"])
        for m, v in enumerate(dat["balances"][i]):
            ws.write_number(row_0, DATA_COL_M1 + m, v, formats["input_value"])
        ws.write_number(row_0, DATA_COL_TRIGGER, dat["triggers"][i], formats["input_value"])

    # Total row
    ws.set_row(DATA_ROW_TOTAL - 1, 24)
    ws.write_string(DATA_ROW_TOTAL - 1, DATA_COL_LABEL, "Total cash", formats["total_left"])
    for m in range(NUM_MONTHS):
        col = col_letter(DATA_COL_M1 + m)
        ws.write_formula(DATA_ROW_TOTAL - 1, DATA_COL_M1 + m,
                         f"=SUM({col}{DATA_ROW_FIRST}:{col}{DATA_ROW_LAST})", formats["total_right"])

    sc.apply_page_setup(ws, sheet_title="Data")


def build_position(wb, formats, dat):
    ws = wb.add_worksheet("Position")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = DATA_COL_TRIGGER + 1
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, NUM_MONTHS + 1, 11)
    ws.set_column(NUM_MONTHS + 2, NUM_MONTHS + 2, 12)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Closing balance and minimum-balance flag",
                       title="Cash position by bank", last_col=LAST_COL + 1,
                       explanation=(
                           "Reads the closing balances from the Data sheet. Each account row carries "
                           "an OK / FLAG status for each month based on the minimum-balance trigger "
                           "set on Data. The Total row sums the closing balances across all accounts."
                       ))

    month_headers = [m.strftime("%b %y") for m in dat["months"]]
    st.write_section_header(ws, formats, row=6, kicker="Drawn from the Data sheet",
                            title="Monthly closing balance with flags")
    header_row = 8
    st.write_header_row(ws, formats, row=header_row, headers=["Bank account"] + month_headers + ["Latest"],
                        right_align_from=2)

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, acct in enumerate(BANK_ACCOUNTS):
        zebra = i % 2 == 0
        idm_row_1b = DATA_ROW_FIRST + i
        formulas = [f"='Data'!{col_letter(DATA_COL_M1 + m)}{idm_row_1b}" for m in range(NUM_MONTHS)]
        latest = f"='Data'!{col_letter(DATA_COL_M_LAST)}{idm_row_1b}"
        st.write_data_row(ws, formats, row=r, label=acct, formulas=formulas + [latest], zebra=zebra, cell_format="td_right")
        r += 1

    # Total row
    total_formulas = [f"='Data'!{col_letter(DATA_COL_M1 + m)}{DATA_ROW_TOTAL}" for m in range(NUM_MONTHS)]
    total_latest = f"='Data'!{col_letter(DATA_COL_M_LAST)}{DATA_ROW_TOTAL}"
    st.write_total_row(ws, formats, row=r, label="Total cash", formulas=total_formulas + [total_latest])
    total_row_1b = r + 1
    r += 2

    # Trigger flag block
    st.write_section_header(ws, formats, row=r, kicker="Minimum balance check",
                            title="Months below trigger by account")
    flag_header_row = r + 2
    st.write_header_row(ws, formats, row=flag_header_row, headers=["Bank account"] + month_headers + ["Trigger"],
                        right_align_from=2)
    r = flag_header_row + 1
    for i, acct in enumerate(BANK_ACCOUNTS):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        idm_row_1b = DATA_ROW_FIRST + i
        ws.set_row(r, 22)
        ws.write_string(r, 1, acct, label_fmt)
        for m in range(NUM_MONTHS):
            f_ = (
                f"=IF('Data'!{col_letter(DATA_COL_M1 + m)}{idm_row_1b}<'Data'!${col_letter(DATA_COL_TRIGGER)}${idm_row_1b},\"FLAG\",\"OK\")"
            )
            ws.write_formula(r, 2 + m, f_, formats["check_status_neutral"])
        trigger_f = f"='Data'!${col_letter(DATA_COL_TRIGGER)}${idm_row_1b}"
        ws.write_formula(r, 2 + NUM_MONTHS, trigger_f, formats["td_right"])
        status_range = f"{col_letter(2)}{r + 1}:{col_letter(2 + NUM_MONTHS - 1)}{r + 1}"
        ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["check_status_ok"]})
        ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "FLAG", "format": formats["check_status_flag"]})
        r += 1

    # Tie-out checks
    checks_end = st.write_checks_block(
        ws, formats,
        row=r + 2,
        checks=[
            {"name": "Sum of bank account totals equals Total cash row latest column",
             "left": f"=SUM({col_letter(2 + NUM_MONTHS)}{first_data_row_1b}:{col_letter(2 + NUM_MONTHS)}{first_data_row_1b + len(BANK_ACCOUNTS) - 1})",
             "right": f"={col_letter(2 + NUM_MONTHS)}{total_row_1b}"},
        ],
    )

    # Chart
    chart_anchor = f"B{checks_end + 2}"
    cats_range = f"='Position'!${col_letter(2)}${header_row + 1}:${col_letter(2 + NUM_MONTHS - 1)}${header_row + 1}"
    series = []
    for i, acct in enumerate(BANK_ACCOUNTS):
        ROW = first_data_row_1b + i
        series.append({
            "name": acct,
            "values": f"='Position'!${col_letter(2)}${ROW}:${col_letter(2 + NUM_MONTHS - 1)}${ROW}",
            "color": b.GREEN_SPECTRUM[1 + i * 2],
        })
    series.append({
        "name": "Total cash",
        "values": f"='Position'!${col_letter(2)}${total_row_1b}:${col_letter(2 + NUM_MONTHS - 1)}${total_row_1b}",
        "color": b.CHART_PRIMARY,
    })
    st.add_line_chart(wb, ws, title="Cash position trend", anchor_cell=chart_anchor,
                      series=series, cats_range=cats_range, width=720, height=300)

    sc.apply_page_setup(ws, sheet_title="Position")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_position(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
