"""Chart of Accounts Setup (1010).

Setup register for the chart of accounts your bookkeeper will load into
your accounting software. Pre-populated with a Lyros-recommended SME COA
that mirrors the header structure of the management reporting pack.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\1010 - Chart of Accounts Setup.xlsx")

WORKBOOK_ID = "1010"
WORKBOOK_TITLE = "Chart of Accounts Setup"
WORKBOOK_KICKER = "Setup register for your accounting software"
TARGET_USER = (
    "Bookkeeper setting up a new accounting-software file, or finance "
    "lead cleaning up an existing chart that has grown organically."
)
HOW_TO_USE = [
    "Open the Register sheet and review the recommended chart. Add or remove rows to suit your business.",
    "Pick the Account type and Tax default for each account from the dropdowns.",
    "The Validation block flags any account that is missing a code, name, type, or tax default.",
]

EXAMPLE_PROFILE = [
    ("BUSINESS TYPE", "Trading SME under $20M turnover, GST registered"),
    ("STRUCTURE", "Single entity with three departments (Sales, Operations, Admin)"),
    ("REPORTING SHAPE", "Header accounts mirror the Management Reporting Pack so monthly reads come out tidy"),
]

INPUTS_REQUIRED = [
    ("Account code, name, type, tax default", "Register tab"),
    ("Header (parent) mapping per account", "Register tab"),
]

ACCOUNT_TYPES = [
    "Current Asset", "Non-current Asset", "Current Liability",
    "Non-current Liability", "Equity", "Revenue", "Direct Cost",
    "Operating Expense", "Other Income", "Other Expense",
]
TAX_DEFAULTS = [
    "GST on Income", "GST on Expenses", "GST Free Income", "GST Free Expenses",
    "BAS Excluded", "Input Taxed Sales", "Capital Acquisition",
]

# Recommended SME chart. Header column groups for reporting roll-up.
COA = [
    # code, name, type, tax_default, header
    ("100", "Cash at bank - operating", "Current Asset", "BAS Excluded", "Bank"),
    ("101", "Cash at bank - tax set aside", "Current Asset", "BAS Excluded", "Bank"),
    ("110", "Trade debtors", "Current Asset", "BAS Excluded", "Receivables"),
    ("120", "Inventory on hand", "Current Asset", "BAS Excluded", "Inventory"),
    ("130", "Prepayments", "Current Asset", "BAS Excluded", "Other current assets"),
    ("140", "GST receivable", "Current Asset", "BAS Excluded", "GST and tax"),
    ("200", "Plant and equipment - cost", "Non-current Asset", "BAS Excluded", "Fixed assets"),
    ("201", "Plant and equipment - accumulated depreciation", "Non-current Asset", "BAS Excluded", "Fixed assets"),
    ("210", "Motor vehicles - cost", "Non-current Asset", "BAS Excluded", "Fixed assets"),
    ("211", "Motor vehicles - accumulated depreciation", "Non-current Asset", "BAS Excluded", "Fixed assets"),
    ("300", "Trade creditors", "Current Liability", "BAS Excluded", "Payables"),
    ("310", "GST payable", "Current Liability", "BAS Excluded", "GST and tax"),
    ("320", "PAYG withholding payable", "Current Liability", "BAS Excluded", "GST and tax"),
    ("330", "Superannuation payable", "Current Liability", "BAS Excluded", "Payroll liabilities"),
    ("340", "Wages payable", "Current Liability", "BAS Excluded", "Payroll liabilities"),
    ("350", "Provision for annual leave", "Current Liability", "BAS Excluded", "Payroll liabilities"),
    ("360", "Income tax payable", "Current Liability", "BAS Excluded", "GST and tax"),
    ("400", "Bank loan - long term", "Non-current Liability", "BAS Excluded", "Borrowings"),
    ("500", "Issued capital", "Equity", "BAS Excluded", "Equity"),
    ("510", "Retained earnings", "Equity", "BAS Excluded", "Equity"),
    ("600", "Sales - core", "Revenue", "GST on Income", "Revenue"),
    ("601", "Sales - secondary", "Revenue", "GST on Income", "Revenue"),
    ("610", "Sales - export", "Revenue", "GST Free Income", "Revenue"),
    ("700", "Cost of goods sold - materials", "Direct Cost", "GST on Expenses", "Cost of sales"),
    ("710", "Cost of goods sold - freight in", "Direct Cost", "GST on Expenses", "Cost of sales"),
    ("720", "Direct labour", "Direct Cost", "BAS Excluded", "Cost of sales"),
    ("800", "Wages and salaries", "Operating Expense", "BAS Excluded", "Wages and on-costs"),
    ("801", "Superannuation expense", "Operating Expense", "BAS Excluded", "Wages and on-costs"),
    ("802", "Workers compensation", "Operating Expense", "GST on Expenses", "Wages and on-costs"),
    ("810", "Rent", "Operating Expense", "GST on Expenses", "Occupancy"),
    ("811", "Light, power, heating", "Operating Expense", "GST on Expenses", "Occupancy"),
    ("820", "Marketing and advertising", "Operating Expense", "GST on Expenses", "Marketing"),
    ("830", "Office expenses", "Operating Expense", "GST on Expenses", "Administration"),
    ("840", "Accounting and legal fees", "Operating Expense", "GST on Expenses", "Administration"),
    ("850", "Insurance", "Operating Expense", "GST on Expenses", "Administration"),
    ("860", "Bank fees", "Operating Expense", "GST Free Expenses", "Administration"),
    ("870", "Depreciation - plant and equipment", "Operating Expense", "BAS Excluded", "Depreciation"),
    ("871", "Depreciation - motor vehicles", "Operating Expense", "BAS Excluded", "Depreciation"),
    ("880", "Interest expense", "Operating Expense", "BAS Excluded", "Finance costs"),
    ("900", "Interest income", "Other Income", "BAS Excluded", "Other income"),
    ("910", "Gain on disposal of assets", "Other Income", "BAS Excluded", "Other income"),
    ("990", "Income tax expense", "Other Expense", "BAS Excluded", "Income tax"),
]


def build_register(wb, formats):
    ws = wb.add_worksheet("Register")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 8
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 7)    # Code
    ws.set_column(2, 2, 36)   # Name
    ws.set_column(3, 3, 22)   # Type
    ws.set_column(4, 4, 22)   # Tax default
    ws.set_column(5, 5, 22)   # Header
    ws.set_column(6, 6, 12)   # Status
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Setup register",
        title="Chart of accounts",
        last_col=LAST_COL + 1,
        explanation=(
            "Each row is one ledger account ready to be loaded into your accounting "
            "software. The Type and Tax default dropdowns pull from a standard list. "
            "The Header column drives the reporting roll-up in the Management Reporting "
            "Pack and other workbooks. The Status column flags any incomplete row."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Step 1   Review and amend the chart",
                            title="Accounts to load")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Code", "Account name", "Type", "Tax default", "Header", "Status"],
        right_align_from=99,
    )

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (code, name, typ, tax, header) in enumerate(COA):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, code, label_fmt)
        ws.write_string(r, 2, name, label_fmt)
        ws.write_string(r, 3, typ, formats["input_text"])
        ws.write_string(r, 4, tax, formats["input_text"])
        ws.write_string(r, 5, header, formats["input_text"])
        status_f = (
            f'=IF(AND(B{r + 1}<>"",C{r + 1}<>"",D{r + 1}<>"",E{r + 1}<>"",F{r + 1}<>""),'
            f'"OK","INCOMPLETE")'
        )
        ws.write_formula(r, 6, status_f, formats["check_status_neutral"])
        r += 1
    last_data_row_1b = r

    ws.data_validation(first_data_row_1b - 1, 3, last_data_row_1b - 1, 3,
                       {"validate": "list", "source": ACCOUNT_TYPES})
    ws.data_validation(first_data_row_1b - 1, 4, last_data_row_1b - 1, 4,
                       {"validate": "list", "source": TAX_DEFAULTS})

    status_range = f"G{first_data_row_1b}:G{last_data_row_1b}"
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["check_status_ok"]})
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "INCOMPLETE", "format": formats["check_status_flag"]})

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Every account has a complete setup (no blanks)",
             "left": f"=COUNTIF(G{first_data_row_1b}:G{last_data_row_1b},\"OK\")",
             "right": f"=COUNTA(B{first_data_row_1b}:B{last_data_row_1b})"},
            {"name": "Account codes are unique",
             "left": f"=SUMPRODUCT(1/COUNTIF(B{first_data_row_1b}:B{last_data_row_1b},B{first_data_row_1b}:B{last_data_row_1b}))",
             "right": f"=COUNTA(B{first_data_row_1b}:B{last_data_row_1b})"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Register")


def build_reference(wb, formats):
    ws = wb.add_worksheet("Reference")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 4
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 26)
    ws.set_column(2, 2, 60)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Reference",
        title="Account types and tax defaults",
        last_col=LAST_COL + 2,
        explanation=(
            "Reference list for the dropdowns on the Register sheet. Type drives where "
            "the account sits on the balance sheet or P&L. Tax default drives how new "
            "transactions are coded when no specific rule is set."
        ),
    )

    r = 6
    st.write_section_header(ws, formats, row=r, kicker="Reference", title="Account types")
    st.write_header_row(ws, formats, row=r + 2, headers=["Type", "Sits under"], right_align_from=99)
    descriptions = {
        "Current Asset": "Balance sheet, expected to convert to cash within 12 months",
        "Non-current Asset": "Balance sheet, longer than 12 months",
        "Current Liability": "Balance sheet, due within 12 months",
        "Non-current Liability": "Balance sheet, longer than 12 months",
        "Equity": "Balance sheet, owners' interest",
        "Revenue": "Profit and loss, top line",
        "Direct Cost": "Profit and loss, cost of sales",
        "Operating Expense": "Profit and loss, overheads",
        "Other Income": "Profit and loss, below operating",
        "Other Expense": "Profit and loss, below operating",
    }
    for i, t in enumerate(ACCOUNT_TYPES):
        rr = r + 3 + i
        zebra = i % 2 == 0
        fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(rr, 22)
        ws.write_string(rr, 1, t, fmt)
        ws.write_string(rr, 2, descriptions[t], fmt)

    r = r + 3 + len(ACCOUNT_TYPES) + 2
    st.write_section_header(ws, formats, row=r, kicker="Reference", title="Tax defaults")
    st.write_header_row(ws, formats, row=r + 2, headers=["Default", "When to use"], right_align_from=99)
    tax_hints = {
        "GST on Income": "Sales subject to 10 per cent GST",
        "GST on Expenses": "Purchases subject to 10 per cent GST input credit",
        "GST Free Income": "Sales not subject to GST (exports, basic food)",
        "GST Free Expenses": "Purchases not subject to GST (bank fees, wages)",
        "BAS Excluded": "Not reported on BAS (movements between asset and liability accounts)",
        "Input Taxed Sales": "Residential rent, financial supplies",
        "Capital Acquisition": "Fixed asset purchases, claim full input credit and report on BAS",
    }
    for i, t in enumerate(TAX_DEFAULTS):
        rr = r + 3 + i
        zebra = i % 2 == 0
        fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(rr, 22)
        ws.write_string(rr, 1, t, fmt)
        ws.write_string(rr, 2, tax_hints[t], fmt)

    sc.apply_page_setup(ws, sheet_title="Reference")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE,
                       workbook_kicker=WORKBOOK_KICKER,
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_register(wb, formats)
    build_reference(wb, formats)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
