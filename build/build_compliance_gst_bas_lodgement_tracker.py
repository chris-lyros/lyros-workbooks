"""GST and BAS Lodgement Tracker (6100).

Quarterly lodgement and payment register for one or more GST-registered
entities. Tracks due date, lodgement date, BAS amount, payment date, and
status across the next eight quarters with rollover.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\6100 - GST and BAS Lodgement Tracker.xlsx")

WORKBOOK_ID = "6100"
WORKBOOK_TITLE = "GST and BAS Lodgement Tracker"
WORKBOOK_KICKER = "Quarterly BAS register across entities"
TARGET_USER = (
    "Bookkeeper or operator tracking BAS lodgement and payment dates "
    "across one or more GST-registered entities."
)
HOW_TO_USE = [
    "List each GST-registered entity on the Tracker sheet, one row per quarter per entity.",
    "Update the Lodged on, Amount, Paid on, and Status columns as you progress each quarter.",
    "The Summary sheet rolls the tracker up by status (Due, Overdue, Lodged, Paid).",
]

EXAMPLE_PROFILE = [
    ("ENTITIES", "Three Australian Pty Ltd companies, all quarterly cycle"),
    ("LODGEMENT CHANNEL", "Tax agent (28-day extension applies to quarterly BAS due dates)"),
    ("PAYMENT CADENCE", "Pay-on-lodgement; ATO direct debit for two of three entities"),
]

INPUTS_REQUIRED = [
    ("Entity name and ABN", "Tracker tab"),
    ("Quarter period, due date, lodged date, amount, paid date, status", "Tracker tab"),
]

PERIOD_LABELS = [
    "Q1 FY25 (Jul-Sep 24)", "Q2 FY25 (Oct-Dec 24)", "Q3 FY25 (Jan-Mar 25)", "Q4 FY25 (Apr-Jun 25)",
    "Q1 FY26 (Jul-Sep 25)", "Q2 FY26 (Oct-Dec 25)", "Q3 FY26 (Jan-Mar 26)", "Q4 FY26 (Apr-Jun 26)",
]
# Standard agent-lodged due dates (28-day extension on the 28th of the month
# following quarter-end where applicable).
DUE_DATES = [
    date(2024, 11, 25), date(2025, 2, 28), date(2025, 5, 26), date(2025, 8, 25),
    date(2025, 11, 25), date(2026, 2, 28), date(2026, 5, 26), date(2026, 8, 25),
]
STATUS_OPTIONS = ["Due", "Overdue", "Lodged", "Paid"]


def _synthetic_data() -> dict:
    rng = d.make_rng("bas")
    entities = [
        ("Yarra Supplies Pty Ltd", "99 102 478 311"),
        ("Southbank Trading Co", "99 211 884 027"),
        ("Latrobe Manufacturing Pty Ltd", "99 326 119 540"),
    ]
    today = d.ANCHOR_BUILD_DATE
    rows = []
    for ent_name, ent_abn in entities:
        for label, due in zip(PERIOD_LABELS, DUE_DATES):
            base_bas = rng.uniform(4_500, 38_000)
            if due < today - timedelta(days=14):  # past quarter
                lodged = due - timedelta(days=int(rng.uniform(1, 12)))
                paid = lodged + timedelta(days=int(rng.uniform(0, 14)))
                amount = round(base_bas, 0)
                status = "Paid"
            elif due < today:  # within grace window
                lodged = due - timedelta(days=int(rng.uniform(0, 4)))
                paid = None
                amount = round(base_bas, 0)
                status = "Lodged"
            else:  # future
                lodged = None
                paid = None
                amount = 0
                status = "Due"
            rows.append((ent_name, ent_abn, label, due, lodged, amount, paid, status))
    return {"rows": rows}


def build_tracker(wb, formats, dat):
    ws = wb.add_worksheet("Tracker")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 9
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)   # Entity
    ws.set_column(2, 2, 16)   # ABN
    ws.set_column(3, 3, 22)   # Period
    ws.set_column(4, 4, 14)   # Due
    ws.set_column(5, 5, 14)   # Lodged
    ws.set_column(6, 6, 14)   # Amount
    ws.set_column(7, 7, 14)   # Paid
    ws.set_column(8, 8, 12)   # Status
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Lodgement register",
        title="Quarterly BAS by entity",
        last_col=LAST_COL + 1,
        explanation=(
            "One row per entity per quarter. Update the Lodged on, Amount, Paid on, and "
            "Status columns as each quarter progresses. The Status column drives the "
            "Summary sheet's roll-up and the colour highlighting on this tab."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Step 1   Maintain the quarterly register",
                            title="BAS quarters")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Entity", "ABN", "Period", "Due", "Lodged on", "Amount", "Paid on", "Status"],
        right_align_from=6,
    )

    r = header_row + 1
    first_data_row_1b = r + 1
    date_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "font_color": b.ACCENT_SOFT,
        "align": "right", "valign": "vcenter", "bg_color": "#FFFEF7",
        "border": 1, "border_color": b.AMBER, "num_format": "yyyy-mm-dd",
    })
    for i, (ent, abn, period, due, lodged, amount, paid, status) in enumerate(dat["rows"]):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, ent, label_fmt)
        ws.write_string(r, 2, abn, label_fmt)
        ws.write_string(r, 3, period, label_fmt)
        ws.write_datetime(r, 4, due, date_fmt)
        if lodged:
            ws.write_datetime(r, 5, lodged, date_fmt)
        else:
            ws.write_blank(r, 5, None, formats["input_text"])
        if amount:
            ws.write_number(r, 6, amount, formats["input_value"])
        else:
            ws.write_blank(r, 6, None, formats["input_value"])
        if paid:
            ws.write_datetime(r, 7, paid, date_fmt)
        else:
            ws.write_blank(r, 7, None, formats["input_text"])
        ws.write_string(r, 8, status, formats["input_text"])
        r += 1
    last_data_row_1b = r

    ws.data_validation(first_data_row_1b - 1, 8, last_data_row_1b - 1, 8,
                       {"validate": "list", "source": STATUS_OPTIONS})

    status_range = f"I{first_data_row_1b}:I{last_data_row_1b}"
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "Paid", "format": formats["check_status_ok"]})
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "Lodged", "format": formats["check_status_ok"]})
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "Overdue", "format": formats["check_status_flag"]})
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "Due", "format": formats["check_status_neutral"]})

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Every paid row has a paid date",
             "left": f"=COUNTIFS(I{first_data_row_1b}:I{last_data_row_1b},\"Paid\",H{first_data_row_1b}:H{last_data_row_1b},\">0\")",
             "right": f"=COUNTIF(I{first_data_row_1b}:I{last_data_row_1b},\"Paid\")"},
            {"name": "Every lodged or paid row has an amount",
             "left": f"=COUNTIFS(I{first_data_row_1b}:I{last_data_row_1b},\"Paid\",G{first_data_row_1b}:G{last_data_row_1b},\">0\")+COUNTIFS(I{first_data_row_1b}:I{last_data_row_1b},\"Lodged\",G{first_data_row_1b}:G{last_data_row_1b},\">0\")",
             "right": f"=COUNTIF(I{first_data_row_1b}:I{last_data_row_1b},\"Paid\")+COUNTIF(I{first_data_row_1b}:I{last_data_row_1b},\"Lodged\")"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Tracker")


def build_summary(wb, formats, dat):
    ws = wb.add_worksheet("Summary")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 7
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 18)
    ws.set_column(4, 4, 18)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Roll-up",
        title="Status snapshot across entities",
        last_col=LAST_COL + 1,
        explanation=(
            "Rolls the Tracker sheet up by status. Use this view to spot quarters that "
            "have lodged but not been paid, or that remain Due with the deadline near."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="By status", title="Quarters by lodgement state")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Status", "Count", "Amount lodged", "Amount paid"],
                        right_align_from=2)

    r = header_row + 1
    statuses = ["Due", "Overdue", "Lodged", "Paid"]
    for i, s in enumerate(statuses):
        zebra = i % 2 == 0
        ws.set_row(r, 22)
        ws.write_string(r, 1, s, formats["td_zebra"] if zebra else formats["td"])
        ws.write_formula(r, 2, f"=COUNTIF(Tracker!I:I,\"{s}\")",
                         formats["td_right_zebra"] if zebra else formats["td_right"])
        ws.write_formula(r, 3, f"=SUMIF(Tracker!I:I,\"{s}\",Tracker!G:G)",
                         formats["td_right_zebra"] if zebra else formats["td_right"])
        # Amount paid only meaningful for Paid status
        if s == "Paid":
            ws.write_formula(r, 4, f"=SUMIF(Tracker!I:I,\"{s}\",Tracker!G:G)",
                             formats["td_right_zebra"] if zebra else formats["td_right"])
        else:
            ws.write_blank(r, 4, None, formats["td_zebra"] if zebra else formats["td"])
        r += 1

    st.write_total_row(
        ws, formats, row=r, label="All quarters",
        formulas=[
            f"=SUM(C{header_row + 2}:C{r})",
            f"=SUM(D{header_row + 2}:D{r})",
            f"=SUM(E{header_row + 2}:E{r})",
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Summary")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE,
                       workbook_kicker=WORKBOOK_KICKER,
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_tracker(wb, formats, dat)
    build_summary(wb, formats, dat)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
