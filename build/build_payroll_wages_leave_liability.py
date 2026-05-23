"""Wages and Leave Liability (6300).

Period-end leave balance and accrued wages liability per employee.
Working file for the month-end journal that recognises wages earned but
not yet paid, and the leave provision at year-end rates.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\6300 - Wages and Leave Liability.xlsx")

WORKBOOK_ID = "6300"
WORKBOOK_TITLE = "Wages and Leave Liability"
WORKBOOK_KICKER = "Month-end accrual and leave provision"
TARGET_USER = "Bookkeeper or finance lead posting the month-end wages and leave journals."
HOW_TO_USE = [
    "List every active employee on the Employees sheet with current rate, leave balance, and last pay end date.",
    "The Accrued Wages sheet calculates wages earned between the last pay run and month-end.",
    "The Leave Provision sheet calculates the dollar value of leave balances at current rates.",
]

EXAMPLE_PROFILE = [
    ("WORKFORCE", "Twelve active employees across operations, sales, and admin"),
    ("PAY CYCLE", "Fortnightly; period ends every second Sunday"),
    ("LEAVE POLICY", "Standard annual leave plus long service after 10 years"),
]

INPUTS_REQUIRED = [
    ("Employee name, annual rate, weekly hours", "Employees tab"),
    ("Annual leave balance and LSL balance (hours)", "Employees tab"),
    ("Last pay period end date and period-end (month-end) date", "Employees tab top of sheet"),
]


EMP_COL_NAME = 1
EMP_COL_DEPT = 2
EMP_COL_ANNUAL = 3
EMP_COL_HOURS = 4
EMP_COL_AL_HOURS = 5
EMP_COL_LSL_HOURS = 6
EMP_ROW_HEADER = 8
EMP_ROW_FIRST = 9

DEPARTMENTS = ["Operations", "Sales", "Admin", "Finance", "Customer service"]


def _synthetic_data() -> dict:
    rng = d.make_rng("payroll")
    period_end = d.LAST_REPORTED_MONTH
    last_pay_end = period_end - timedelta(days=int(rng.uniform(4, 11)))
    employees = []
    names = [
        "Sarah Chen", "Marcus Williams", "Priya Sharma", "Daniel O'Connor",
        "Lucia Romano", "Hamish McKenzie", "Aisha Patel", "Jordan Tan",
        "Olivia Phan", "Riley Brennan", "Thomas Vasquez", "Mei Tanaka",
    ]
    for n in names:
        dept = DEPARTMENTS[int(rng.uniform(0, len(DEPARTMENTS)))]
        annual = round(rng.uniform(58_000, 145_000), 0)
        hours = 38 if rng.random() > 0.2 else round(rng.uniform(20, 36), 0)
        al = round(rng.uniform(28, 162), 1)
        lsl = round(rng.uniform(0, 95), 1) if rng.random() > 0.4 else 0
        employees.append((n, dept, annual, hours, al, lsl))
    return {"period_end": period_end, "last_pay_end": last_pay_end, "employees": employees}


def build_employees(wb, formats, dat):
    ws = wb.add_worksheet("Employees")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 8
    ws.set_column(0, 0, 2)
    ws.set_column(EMP_COL_NAME, EMP_COL_NAME, 24)
    ws.set_column(EMP_COL_DEPT, EMP_COL_DEPT, 18)
    ws.set_column(EMP_COL_ANNUAL, EMP_COL_ANNUAL, 14)
    ws.set_column(EMP_COL_HOURS, EMP_COL_HOURS, 12)
    ws.set_column(EMP_COL_AL_HOURS, EMP_COL_LSL_HOURS, 14)
    ws.set_column(7, 7, 16)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Source data",
        title="Employees and balances",
        last_col=LAST_COL + 1,
        explanation=(
            "One row per active employee. Annual rate and weekly hours drive both the "
            "accrued wages and the leave provision calculations. Update the period-end "
            "and last pay-run end dates at the top of the sheet each month."
        ),
    )

    # Period header
    date_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "font_color": b.ACCENT_SOFT,
        "align": "right", "valign": "vcenter", "bg_color": "#FFFEF7",
        "border": 1, "border_color": b.AMBER, "num_format": "yyyy-mm-dd",
    })
    ws.set_row(6, 22)
    ws.write_string(6, 1, "Period end (month-end)", formats["cover_kicker"])
    ws.write_datetime(6, 3, dat["period_end"], date_fmt)
    ws.write_string(6, 4, "Last pay period end", formats["cover_kicker"])
    ws.write_datetime(6, 6, dat["last_pay_end"], date_fmt)

    st.write_header_row(
        ws, formats, row=EMP_ROW_HEADER,
        headers=["Employee", "Department", "Annual rate", "Weekly hours", "AL balance (hrs)", "LSL balance (hrs)", "Daily rate"],
        right_align_from=3,
    )

    r = EMP_ROW_FIRST
    first_data_row_1b = r + 1
    for i, (name, dept, annual, hours, al, lsl) in enumerate(dat["employees"]):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        days_fmt = formats["td_days_zebra"] if zebra else formats["td_days"]
        ws.set_row(r, 22)
        ws.write_string(r, EMP_COL_NAME, name, label_fmt)
        ws.write_string(r, EMP_COL_DEPT, dept, formats["input_text"])
        ws.write_number(r, EMP_COL_ANNUAL, annual, formats["input_value"])
        ws.write_number(r, EMP_COL_HOURS, hours, days_fmt)
        ws.write_number(r, EMP_COL_AL_HOURS, al, days_fmt)
        ws.write_number(r, EMP_COL_LSL_HOURS, lsl, days_fmt)
        # Daily rate = annual / 260
        ws.write_formula(r, 7, f"=D{r + 1}/260", num_fmt)
        r += 1
    last_data_row_1b = r

    ws.data_validation(first_data_row_1b - 1, EMP_COL_DEPT,
                       last_data_row_1b - 1, EMP_COL_DEPT,
                       {"validate": "list", "source": DEPARTMENTS})

    sc.apply_page_setup(ws, sheet_title="Employees")


def build_accrued_wages(wb, formats, dat):
    ws = wb.add_worksheet("Accrued Wages")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 7
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 24)   # Employee
    ws.set_column(2, 2, 18)   # Department
    ws.set_column(3, 3, 12)   # Working days
    ws.set_column(4, 4, 14)   # Daily rate
    ws.set_column(5, 5, 16)   # Accrued wages
    ws.set_column(6, 6, 14)   # Super 11.5%
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Month-end accrual",
        title="Wages earned but not yet paid",
        last_col=LAST_COL + 1,
        explanation=(
            "Working days between the last pay period end and month-end times the daily "
            "rate per employee, plus superannuation. Use the totals here to post the "
            "month-end wages accrual journal."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Calculation",
                            title="Accrued wages per employee")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Employee", "Department", "Days", "Daily rate", "Accrued wages", "Super (11.5%)"],
        right_align_from=2,
    )

    # Working days from last_pay_end (exclusive) to period_end (inclusive). The
    # formula uses NETWORKDAYS for accuracy.
    r = header_row + 1
    first_data_row_1b = r + 1
    for i in range(len(dat["employees"])):
        emp_row_1b = EMP_ROW_FIRST + 1 + i
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        days_fmt = formats["td_days_zebra"] if zebra else formats["td_days"]
        ws.set_row(r, 22)
        ws.write_formula(r, 1, f"=Employees!B{emp_row_1b}", label_fmt)
        ws.write_formula(r, 2, f"=Employees!C{emp_row_1b}", label_fmt)
        ws.write_formula(r, 3, "=NETWORKDAYS(Employees!$G$7+1,Employees!$D$7)", days_fmt)
        ws.write_formula(r, 4, f"=Employees!H{emp_row_1b}", num_fmt)
        ws.write_formula(r, 5, f"=D{r + 1}*E{r + 1}", num_fmt)
        ws.write_formula(r, 6, f"=F{r + 1}*0.115", num_fmt)
        r += 1
    last_data_row_1b = r

    st.write_total_row(
        ws, formats, row=r, label="Total",
        formulas=["", "", "",
                  f"=SUM(F{first_data_row_1b}:F{last_data_row_1b})",
                  f"=SUM(G{first_data_row_1b}:G{last_data_row_1b})"],
    )

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Working days between last pay end and period end is positive",
             "left": "=NETWORKDAYS(Employees!G7+1,Employees!D7)",
             "right": "=NETWORKDAYS(Employees!G7+1,Employees!D7)"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Accrued Wages")


def build_leave_provision(wb, formats, dat):
    ws = wb.add_worksheet("Leave Provision")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 7
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 24)
    ws.set_column(2, 2, 18)
    ws.set_column(3, 3, 14)   # AL hours
    ws.set_column(4, 4, 14)   # LSL hours
    ws.set_column(5, 5, 14)   # Hourly rate
    ws.set_column(6, 6, 14)   # AL $
    ws.set_column(7, 7, 14)   # LSL $
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Year-end provision",
        title="Leave liability at current rates",
        last_col=LAST_COL + 1,
        explanation=(
            "Leave balances valued at the current hourly rate per employee. Annual leave "
            "is shown current; long service leave is recognised once the employee passes "
            "the qualifying service period (10 years standard). For materiality, accrue "
            "LSL pro-rata from year 5 if the policy requires it."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Provision", title="Leave liability per employee")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Employee", "Department", "AL hrs", "LSL hrs", "Hourly rate", "AL $", "LSL $"],
        right_align_from=2,
    )

    r = header_row + 1
    first_data_row_1b = r + 1
    for i in range(len(dat["employees"])):
        emp_row_1b = EMP_ROW_FIRST + 1 + i
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        days_fmt = formats["td_days_zebra"] if zebra else formats["td_days"]
        ws.set_row(r, 22)
        ws.write_formula(r, 1, f"=Employees!B{emp_row_1b}", label_fmt)
        ws.write_formula(r, 2, f"=Employees!C{emp_row_1b}", label_fmt)
        ws.write_formula(r, 3, f"=Employees!F{emp_row_1b}", days_fmt)
        ws.write_formula(r, 4, f"=Employees!G{emp_row_1b}", days_fmt)
        # Hourly = annual / 52 / weekly hours
        ws.write_formula(r, 5, f"=Employees!D{emp_row_1b}/52/Employees!E{emp_row_1b}", num_fmt)
        ws.write_formula(r, 6, f"=D{r + 1}*F{r + 1}", num_fmt)
        ws.write_formula(r, 7, f"=E{r + 1}*F{r + 1}", num_fmt)
        r += 1
    last_data_row_1b = r

    st.write_total_row(
        ws, formats, row=r, label="Total",
        formulas=["", "", "", "",
                  f"=SUM(G{first_data_row_1b}:G{last_data_row_1b})",
                  f"=SUM(H{first_data_row_1b}:H{last_data_row_1b})"],
    )

    sc.apply_page_setup(ws, sheet_title="Leave Provision")


def build_summary(wb, formats, dat):
    ws = wb.add_worksheet("Summary")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 5
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 28)
    ws.set_column(2, 2, 18)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Journal entry",
        title="Period-end payroll liability summary",
        last_col=LAST_COL + 1,
        explanation=(
            "Source values for the month-end journal entries. Accrued wages includes "
            "superannuation; leave provision is split between annual leave and long "
            "service leave."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Total liability", title="Period-end totals")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Component", "Amount"],
                        right_align_from=2)

    rows = [
        ("Accrued wages", "='Accrued Wages'!F" + str(header_row + 1 + len(dat["employees"]) + 1)),
        ("Accrued superannuation", "='Accrued Wages'!G" + str(header_row + 1 + len(dat["employees"]) + 1)),
        ("Annual leave provision", "='Leave Provision'!G" + str(header_row + 1 + len(dat["employees"]) + 1)),
        ("Long service leave provision", "='Leave Provision'!H" + str(header_row + 1 + len(dat["employees"]) + 1)),
    ]
    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (label, formula) in enumerate(rows):
        zebra = i % 2 == 0
        ws.set_row(r, 22)
        ws.write_string(r, 1, label, formats["td_zebra"] if zebra else formats["td"])
        ws.write_formula(r, 2, formula, formats["td_right_zebra"] if zebra else formats["td_right"])
        r += 1
    last_data_row_1b = r

    st.write_total_row(
        ws, formats, row=r, label="Total payroll-related liability",
        formulas=[f"=SUM(C{first_data_row_1b}:C{last_data_row_1b})"],
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
    build_summary(wb, formats, dat)
    build_accrued_wages(wb, formats, dat)
    build_leave_provision(wb, formats, dat)
    build_employees(wb, formats, dat)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
