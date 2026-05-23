"""Month-End Close Checklist (2010).

Canonical close checklist with sign-off gates and ageing thresholds.
Used monthly; flags any open item past the sign-off threshold.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\2010 - Month-End Close Checklist.xlsx")

WORKBOOK_ID = "2010"
WORKBOOK_TITLE = "Month-End Close Checklist"
WORKBOOK_KICKER = "Repeatable close process for SMEs"
TARGET_USER = "Bookkeeper or finance lead running a monthly close in under five business days."
HOW_TO_USE = [
    "Update the Period header on the Checklist sheet to the month being closed.",
    "Tick each step as Done, Skipped, or N/A, recording the owner and date.",
    "The Summary sheet shows completion rate and flags any step still Open past its day target.",
]

EXAMPLE_PROFILE = [
    ("BUSINESS SIZE", "Owner-operated SME, one bookkeeper plus part-time finance lead"),
    ("CLOSE CADENCE", "Five business days from month-end to manager sign-off"),
    ("CRITICAL GATES", "Bank rec by day 2, payroll posted by day 3, sign-off by day 5"),
]

INPUTS_REQUIRED = [
    ("Period being closed (month-end date)", "Checklist tab top of sheet"),
    ("Status, owner, and completion date per step", "Checklist tab rows"),
]

STATUS_OPTIONS = ["Open", "In progress", "Done", "Skipped", "N/A"]

# (category, step, target_day, default_owner)
STEPS = [
    ("Banking and cash", "Reconcile every bank account to statement", 2, "Bookkeeper"),
    ("Banking and cash", "Clear unpresented cheques older than 90 days", 2, "Bookkeeper"),
    ("Banking and cash", "Reconcile credit card statements to GL", 2, "Bookkeeper"),
    ("AR", "Post all sales invoices for the period", 1, "Bookkeeper"),
    ("AR", "Review ageing and chase 60+ day balances", 3, "Bookkeeper"),
    ("AR", "Post any bad debt write-offs and reverse if recovered", 3, "Finance lead"),
    ("AP", "Post all supplier invoices for the period", 1, "Bookkeeper"),
    ("AP", "Review unprocessed invoices and accrue if material", 2, "Bookkeeper"),
    ("AP", "Match purchase orders to invoices and resolve variances", 2, "Bookkeeper"),
    ("Inventory", "Roll forward inventory movement and post adjustments", 3, "Finance lead"),
    ("Inventory", "Reconcile inventory subsidiary to GL control account", 3, "Finance lead"),
    ("Fixed assets", "Post depreciation journal for the month", 3, "Bookkeeper"),
    ("Fixed assets", "Add new asset acquisitions to the register", 3, "Bookkeeper"),
    ("Payroll", "Post final pay run for the period", 3, "Bookkeeper"),
    ("Payroll", "Accrue wages from last pay date to month-end", 3, "Bookkeeper"),
    ("Payroll", "Reconcile superannuation payable to subsidiary", 4, "Finance lead"),
    ("Tax", "Reconcile GST receivable and payable to BAS-shaped report", 4, "Finance lead"),
    ("Tax", "Reconcile PAYG withholding to STP filings", 4, "Finance lead"),
    ("Tax", "Reconcile income tax payable movement", 4, "Finance lead"),
    ("Reporting", "Run trial balance and review for unusual movements", 4, "Finance lead"),
    ("Reporting", "Run P&L and balance sheet and review variances", 4, "Finance lead"),
    ("Reporting", "Prepare management pack and commentary", 5, "Finance lead"),
    ("Sign-off", "Bookkeeper signs off transactional close", 4, "Bookkeeper"),
    ("Sign-off", "Finance lead signs off month-end reporting", 5, "Finance lead"),
    ("Sign-off", "Management pack issued to leadership team", 5, "Finance lead"),
]


def _synthetic_data() -> dict:
    rng = d.make_rng("close")
    period = d.LAST_REPORTED_MONTH
    rows = []
    today_day = 6  # imagine we are on day 6 of close
    for cat, step, day, owner in STEPS:
        # Most steps before day 4 are done in this example; later ones progressing
        if day <= 3:
            status = "Done"
            done_on = period + timedelta(days=int(rng.uniform(1, day)))
            done_by = owner
        elif day == 4:
            status = "Done" if rng.random() > 0.3 else "In progress"
            done_on = period + timedelta(days=day) if status == "Done" else None
            done_by = owner if status == "Done" else ""
        else:
            status = "In progress"
            done_on = None
            done_by = ""
        rows.append((cat, step, day, owner, status, done_by, done_on))
    return {"period": period, "rows": rows, "today_day": today_day}


def build_checklist(wb, formats, dat):
    ws = wb.add_worksheet("Checklist")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 9
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 16)   # Category
    ws.set_column(2, 2, 50)   # Step
    ws.set_column(3, 3, 8)    # Target day
    ws.set_column(4, 4, 16)   # Owner
    ws.set_column(5, 5, 14)   # Status
    ws.set_column(6, 6, 18)   # Done by
    ws.set_column(7, 7, 14)   # Done on
    ws.set_column(8, 8, 12)   # Flag
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Step-by-step close",
        title=f"Close for {dat['period'].strftime('%B %Y')}",
        last_col=LAST_COL + 1,
        explanation=(
            "Each row is one step in the monthly close. Target day is the business-day "
            "number from month-end (day 1 is the first business day of the new month). "
            "Mark Status, Done by, and Done on as you complete each step. The Flag "
            "column highlights any step still Open or In progress past its target day."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Step 1   Work through every line",
                            title="Close steps")
    header_row = section_row + 2

    # Day counter input
    ws.set_row(section_row + 1, 22)
    ws.write_string(section_row + 1, 7, "Day of close:", formats["cover_kicker"])
    ws.write_number(section_row + 1, 8, dat["today_day"], formats["input_value"])

    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Category", "Step", "Day", "Owner", "Status", "Done by", "Done on", "Flag"],
        right_align_from=3,
    )

    r = header_row + 1
    first_data_row_1b = r + 1
    date_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "font_color": b.ACCENT_SOFT,
        "align": "right", "valign": "vcenter", "bg_color": "#FFFEF7",
        "border": 1, "border_color": b.AMBER, "num_format": "yyyy-mm-dd",
    })
    for i, (cat, step, day, owner, status, done_by, done_on) in enumerate(dat["rows"]):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, cat, label_fmt)
        ws.write_string(r, 2, step, label_fmt)
        ws.write_number(r, 3, day, formats["td_right_zebra"] if zebra else formats["td_right"])
        ws.write_string(r, 4, owner, formats["input_text"])
        ws.write_string(r, 5, status, formats["input_text"])
        ws.write_string(r, 6, done_by, formats["input_text"])
        if done_on:
            ws.write_datetime(r, 7, done_on, date_fmt)
        else:
            ws.write_blank(r, 7, None, formats["input_text"])
        # Flag: OVERDUE if not Done/Skipped/N/A and day < day_of_close
        flag_f = (
            f'=IF(OR(F{r + 1}="Done",F{r + 1}="Skipped",F{r + 1}="N/A"),"OK",'
            f'IF(D{r + 1}<$I${section_row + 2},"OVERDUE","ON TRACK"))'
        )
        ws.write_formula(r, 8, flag_f, formats["check_status_neutral"])
        r += 1
    last_data_row_1b = r

    ws.data_validation(first_data_row_1b - 1, 5, last_data_row_1b - 1, 5,
                       {"validate": "list", "source": STATUS_OPTIONS})

    flag_range = f"I{first_data_row_1b}:I{last_data_row_1b}"
    ws.conditional_format(flag_range, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["check_status_ok"]})
    ws.conditional_format(flag_range, {"type": "text", "criteria": "containing", "value": "OVERDUE", "format": formats["check_status_flag"]})
    ws.conditional_format(flag_range, {"type": "text", "criteria": "containing", "value": "ON TRACK", "format": formats["check_status_neutral"]})

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "All steps have a status set",
             "left": f"=COUNTA(F{first_data_row_1b}:F{last_data_row_1b})",
             "right": f"=COUNTA(C{first_data_row_1b}:C{last_data_row_1b})"},
            {"name": "No overdue items still open",
             "left": f"=COUNTIF(I{first_data_row_1b}:I{last_data_row_1b},\"OVERDUE\")",
             "right": "=0"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Checklist")


def build_summary(wb, formats, dat):
    ws = wb.add_worksheet("Summary")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 6
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 3, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Roll-up",
        title="Close progress by category",
        last_col=LAST_COL + 1,
        explanation=(
            "Completion rate and overdue counts per category. Use this view to identify "
            "where the close is stuck before reading through every line of the checklist."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="By category", title="Status by category")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Category", "Steps", "Done", "Overdue"],
                        right_align_from=2)

    categories = []
    for cat, *_ in STEPS:
        if cat not in categories:
            categories.append(cat)

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, cat in enumerate(categories):
        zebra = i % 2 == 0
        ws.set_row(r, 22)
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.write_string(r, 1, cat, label_fmt)
        ws.write_formula(r, 2, f"=COUNTIF(Checklist!B:B,\"{cat}\")", num_fmt)
        ws.write_formula(r, 3, f"=COUNTIFS(Checklist!B:B,\"{cat}\",Checklist!F:F,\"Done\")", num_fmt)
        ws.write_formula(r, 4, f"=COUNTIFS(Checklist!B:B,\"{cat}\",Checklist!I:I,\"OVERDUE\")", num_fmt)
        r += 1
    last_data_row_1b = r

    st.write_total_row(
        ws, formats, row=r, label="All categories",
        formulas=[
            f"=SUM(C{first_data_row_1b}:C{last_data_row_1b})",
            f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})",
            f"=SUM(E{first_data_row_1b}:E{last_data_row_1b})",
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
    build_checklist(wb, formats, dat)
    build_summary(wb, formats, dat)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
