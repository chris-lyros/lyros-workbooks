"""Supplier to Chart of Accounts mapping workbook.

Setup register: maps each supplier to its default expense account, GST
treatment, and approval rule. Used by bank-rule setup or a bookkeeper
cleaning up an existing file.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\1020 - Supplier to COA Mapping.xlsx")

WORKBOOK_ID = "1020"
WORKBOOK_TITLE = "Supplier to Chart of Accounts Mapping"
WORKBOOK_KICKER = "Supplier setup register"
TARGET_USER = "Bookkeeper setting up a new accounting-software file or cleaning rules on an existing file."
HOW_TO_USE = [
    "List each supplier you transact with in the Supplier column on the Mapping sheet.",
    "Pick the default account code, the GST treatment, and the approval rule per supplier.",
    "The Validation block at the bottom checks that every supplier has a complete mapping (no blanks).",
]

EXAMPLE_PROFILE = [
    ("USE CASE", "Accounting-software file cleanup or new file setup"),
    ("OUTPUT", "Mapping register ready to drive bank rules and supplier defaults"),
    ("APPROVALS", "Three approval tiers: Auto-approve (low risk), 1-approver, 2-approver"),
]

INPUTS_REQUIRED = [
    ("Supplier name", "Mapping tab"),
    ("Default account code and name", "Mapping tab (dropdown from Chart of Accounts)"),
    ("GST treatment", "Mapping tab (dropdown)"),
    ("Approval rule", "Mapping tab (dropdown)"),
]

GST_TREATMENTS = [
    "GST on Expenses (10%)",
    "GST Free Expenses",
    "BAS Excluded",
    "Capital Acquisition (GST on Capital)",
    "No GST (Input Taxed Sales)",
]
APPROVAL_RULES = [
    "Auto-approve (under $500)",
    "1 approver (up to $5k)",
    "2 approvers (up to $20k)",
    "Director approval required (over $20k)",
]

# Chart of accounts list (standard)
COA_LIST = [
    ("310", "Cost of Goods Sold"),
    ("320", "Purchases - Materials"),
    ("469", "Rent"),
    ("451", "Light Power Heating"),
    ("433", "Insurance"),
    ("461", "Marketing and Advertising"),
    ("463", "Office Expenses"),
    ("466", "Accounting and Legal Fees"),
    ("478", "Superannuation"),
    ("479", "Workers Compensation"),
    ("416", "Depreciation"),
    ("700", "Property Plant Equipment"),
]


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
    rng = d.make_rng("coa")
    roster = d.roster()[:15]
    rows = []
    for i, co in enumerate(roster):
        # Assign reasonable defaults based on industry
        if co.industry in ("Wholesale", "Manufacturing", "Retail"):
            acct = COA_LIST[0]  # COGS
            gst = GST_TREATMENTS[0]
            approval = APPROVAL_RULES[1]
        elif co.industry in ("Transport", "Marine services"):
            acct = COA_LIST[1]  # Purchases
            gst = GST_TREATMENTS[0]
            approval = APPROVAL_RULES[1]
        elif co.industry == "Property":
            acct = COA_LIST[2]  # Rent
            gst = GST_TREATMENTS[0]
            approval = APPROVAL_RULES[2]
        elif co.industry in ("Consulting", "Software", "Creative agency"):
            acct = COA_LIST[7]  # Accounting/legal
            gst = GST_TREATMENTS[0]
            approval = APPROVAL_RULES[1]
        else:
            acct = COA_LIST[6]  # Office
            gst = GST_TREATMENTS[0]
            approval = APPROVAL_RULES[0]
        rows.append((co.name, co.abn, acct, gst, approval))
    return {"rows": rows}


def build_coa_reference(wb, formats):
    ws = wb.add_worksheet("Chart of Accounts")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)
    LAST_COL = 4
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 8)
    ws.set_column(2, 2, 32)
    ws.set_column(3, 3, 16)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Reference table",
                       title="Chart of accounts",
                       last_col=LAST_COL + 2,
                       explanation=(
                           "List of the expense accounts available for supplier mapping. "
                           "Add or amend the codes here to match your own chart of accounts."
                       ))

    st.write_section_header(ws, formats, row=6, kicker="Reference", title="Expense accounts")
    st.write_header_row(ws, formats, row=8, headers=["Code", "Account name", "Tax default"], right_align_from=3)
    for i, (code, name) in enumerate(COA_LIST):
        row_0 = 9 + i
        ws.set_row(row_0, 22)
        zebra = i % 2 == 0
        text_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.write_string(row_0, 1, code, text_fmt)
        ws.write_string(row_0, 2, name, text_fmt)
        ws.write_string(row_0, 3, "GST on Expenses (10%)", text_fmt)

    sc.apply_page_setup(ws, sheet_title="Chart of Accounts")


def build_mapping(wb, formats, dat):
    ws = wb.add_worksheet("Mapping")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = 7
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 32)
    ws.set_column(2, 2, 18)
    ws.set_column(3, 3, 8)
    ws.set_column(4, 4, 26)
    ws.set_column(5, 5, 22)
    ws.set_column(6, 6, 26)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Setup register",
                       title="Supplier to chart of accounts mapping",
                       last_col=LAST_COL + 2,
                       explanation=(
                           "Each row maps one supplier to a default account, GST treatment, and "
                           "approval rule. The dropdowns on the Account Code, GST, and Approval "
                           "columns pull from the Chart of Accounts reference sheet and the standard "
                           "tax / approval lists. The Validation status flags any incomplete row."
                       ))

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Step 1   Map each supplier",
                            title="Supplier setup")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Supplier", "ABN", "Account code", "Account name",
                                 "GST treatment", "Approval rule", "Status"],
                        right_align_from=99)  # left align for this table

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (name, abn, (acct_code, acct_name), gst, approval) in enumerate(dat["rows"]):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, name, label_fmt)
        ws.write_string(r, 2, abn, label_fmt)
        ws.write_string(r, 3, acct_code, formats["input_text"])
        ws.write_string(r, 4, acct_name, formats["input_text"])
        ws.write_string(r, 5, gst, formats["input_text"])
        ws.write_string(r, 6, approval, formats["input_text"])
        # Status: OK if all four mappings present
        status_f = f"=IF(AND(D{r + 1}<>\"\",E{r + 1}<>\"\",F{r + 1}<>\"\",G{r + 1}<>\"\"),\"OK\",\"INCOMPLETE\")"
        ws.write_formula(r, 7, status_f, formats["check_status_neutral"])
        r += 1
    last_data_row_1b = r

    # Add data validations
    coa_codes = [c for c, _ in COA_LIST]
    ws.data_validation(first_data_row_1b - 1, 3, last_data_row_1b - 1, 3,
                       {"validate": "list", "source": coa_codes})
    ws.data_validation(first_data_row_1b - 1, 5, last_data_row_1b - 1, 5,
                       {"validate": "list", "source": GST_TREATMENTS})
    ws.data_validation(first_data_row_1b - 1, 6, last_data_row_1b - 1, 6,
                       {"validate": "list", "source": APPROVAL_RULES})

    # CF on Status column
    status_range = f"H{first_data_row_1b}:H{last_data_row_1b}"
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "OK", "format": formats["check_status_ok"]})
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "INCOMPLETE", "format": formats["check_status_flag"]})

    # Validation summary
    r += 2
    st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {"name": "All suppliers have a complete mapping (no blanks)",
             "left": f"=COUNTIF(H{first_data_row_1b}:H{last_data_row_1b},\"OK\")",
             "right": f"=COUNTA(B{first_data_row_1b}:B{last_data_row_1b})"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Mapping")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_mapping(wb, formats, dat)
    build_coa_reference(wb, formats)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
