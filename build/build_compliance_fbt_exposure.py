"""FBT Exposure Quick Check (6200).

Screening tool for common fringe benefits tax triggers. Not an FBT return.
Where this workbook flags exposure the next step is a conversation with a
tax accountant.
"""

from __future__ import annotations
import sys
from pathlib import Path
import xlsxwriter

sys.path.insert(0, str(Path(__file__).parent))
from _shared import branding as b
from _shared import data as d
from _shared import scaffold as sc
from _shared import styles as st


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\6200 - FBT Exposure Quick Check.xlsx")

WORKBOOK_ID = "6200"
WORKBOOK_TITLE = "FBT Exposure Quick Check"
WORKBOOK_KICKER = "Screening tool for common FBT triggers"
TARGET_USER = "Finance lead or bookkeeper screening for FBT exposure before year-end."
HOW_TO_USE = [
    "Work through each category and tick whether the trigger applies to your business.",
    "Enter the indicative annual value of each triggered benefit; the workbook estimates gross-up and FBT.",
    "Review the Exposure Estimate sheet and take any FLAG items to a tax accountant.",
]

EXAMPLE_PROFILE = [
    ("BUSINESS TYPE", "Professional services Pty Ltd with five vehicles and modest entertainment spend"),
    ("FBT REGISTRATION", "Registered with the ATO; lodges 1 April to 31 March year"),
    ("TYPICAL EXPOSURE", "Two cars under statutory method, one client lunch programme, no employee loans"),
]

INPUTS_REQUIRED = [
    ("Trigger applicability (Yes/No)", "Screening tab"),
    ("Indicative annual value per trigger", "Screening tab"),
]

# FBT rate FY26 is 47%. Type 1 gross-up 2.0802. Type 2 gross-up 1.8868.
FBT_RATE = 0.47
GROSS_UP_T1 = 2.0802
GROSS_UP_T2 = 1.8868

# (category, trigger, default_gross_up, hint)
TRIGGERS = [
    ("Cars", "Company car available for private use", "T1", "Statutory formula or operating cost method"),
    ("Cars", "Employee novated lease where employer pays running costs", "T1", "Treat carefully if employee contributions are nil"),
    ("Entertainment", "Client lunches and after-work drinks (employees attend)", "T1", "Employee portion is FBT; client portion is not"),
    ("Entertainment", "Year-end party or off-site events", "T1", "Minor benefits exemption may apply under $300 per employee"),
    ("Expense payments", "Reimbursing employees for personal expenses", "T1", "Travel, phone, internet if not otherwise deductible"),
    ("Property", "Discounted or free goods provided to employees", "T1", "Common in retail and hospitality"),
    ("Residual", "Use of business premises out of hours", "T2", "Function rooms, gyms, holiday houses"),
    ("Loans", "Employee loans below ATO benchmark interest", "T2", "Benchmark rate published annually"),
    ("Living away", "Living-away-from-home allowance paid", "T2", "Strict declaration requirements"),
    ("Car parking", "Employer-provided car parking near commercial parking", "T1", "Commercial station 1km test"),
]


def build_screening(wb, formats):
    ws = wb.add_worksheet("Screening")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 8
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 16)   # Category
    ws.set_column(2, 2, 44)   # Trigger
    ws.set_column(3, 3, 10)   # Applies
    ws.set_column(4, 4, 16)   # Indicative annual value
    ws.set_column(5, 5, 8)    # Gross-up
    ws.set_column(6, 6, 14)   # Grossed up
    ws.set_column(7, 7, 14)   # FBT exposure
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Trigger screen",
        title="Common FBT triggers screened against your spend",
        last_col=LAST_COL + 1,
        explanation=(
            "For each trigger, mark Yes or No and enter an indicative annual value if the "
            "trigger applies. The workbook applies the standard gross-up (Type 1 or Type 2) "
            "and the current FBT rate. This is a screening tool only; actual FBT requires "
            "an FBT return and may need specialist advice."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Step 1   Work through each category",
                            title="FBT trigger screen")

    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Category", "Trigger", "Applies", "Annual value", "Type", "Grossed up", "FBT exposure"],
        right_align_from=4,
    )

    r = header_row + 1
    first_data_row_1b = r + 1
    rng = d.make_rng("fbt")
    for i, (cat, trig, t_type, hint) in enumerate(TRIGGERS):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        # Synthetic example data
        applies = "Yes" if rng.random() > 0.4 else "No"
        sample_val = round(rng.uniform(3_000, 32_000), 0) if applies == "Yes" else 0
        ws.set_row(r, 22)
        ws.write_string(r, 1, cat, label_fmt)
        ws.write_string(r, 2, trig, label_fmt)
        ws.write_string(r, 3, applies, formats["input_text"])
        ws.write_number(r, 4, sample_val, formats["input_value"])
        ws.write_string(r, 5, t_type, label_fmt)
        gross_up_factor = GROSS_UP_T1 if t_type == "T1" else GROSS_UP_T2
        ws.write_formula(r, 6,
                         f'=IF(D{r + 1}="Yes",E{r + 1}*{gross_up_factor},0)',
                         num_fmt)
        ws.write_formula(r, 7,
                         f'=G{r + 1}*{FBT_RATE}',
                         num_fmt)
        r += 1
    last_data_row_1b = r

    ws.data_validation(first_data_row_1b - 1, 3, last_data_row_1b - 1, 3,
                       {"validate": "list", "source": ["Yes", "No"]})

    applies_range = f"D{first_data_row_1b}:D{last_data_row_1b}"
    ws.conditional_format(applies_range, {"type": "text", "criteria": "containing", "value": "Yes", "format": formats["check_status_flag"]})
    ws.conditional_format(applies_range, {"type": "text", "criteria": "containing", "value": "No", "format": formats["check_status_ok"]})

    st.write_total_row(
        ws, formats, row=r, label="Total exposure (indicative)",
        formulas=[
            "", "", "",
            f"=SUM(E{first_data_row_1b}:E{last_data_row_1b})",
            "",
            f"=SUM(G{first_data_row_1b}:G{last_data_row_1b})",
            f"=SUM(H{first_data_row_1b}:H{last_data_row_1b})",
        ],
    )

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "FBT exposure equals grossed-up value times FBT rate",
             "left": f"=SUM(H{first_data_row_1b}:H{last_data_row_1b})",
             "right": f"=SUM(G{first_data_row_1b}:G{last_data_row_1b})*{FBT_RATE}"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Screening")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE,
                       workbook_kicker=WORKBOOK_KICKER,
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_screening(wb, formats)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
