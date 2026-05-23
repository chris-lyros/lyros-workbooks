"""Accounting File Review Checklist (2020).

Quarterly hygiene check on the integrity of an accounting-software file.
Catches the small things that compound: misconfigured tax defaults,
contacts without ABNs, archived account references, untidy tracking
categories. Findings log captures issues raised for follow-up.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\2020 - Accounting File Review Checklist.xlsx")

WORKBOOK_ID = "2020"
WORKBOOK_TITLE = "Accounting File Review Checklist"
WORKBOOK_KICKER = "Quarterly hygiene check on your accounting file"
TARGET_USER = (
    "External reviewer or fractional CFO running a quarterly hygiene "
    "review on a client accounting-software file."
)
HOW_TO_USE = [
    "Work through each review question on the Checklist sheet and record what you found.",
    "For any exception, log a finding on the Findings sheet with owner and target date.",
    "The Summary sheet rolls findings up by severity and tracks completion of remediation actions.",
]

EXAMPLE_PROFILE = [
    ("REVIEW SCOPE", "Quarterly file review for an SME on a single accounting-software platform"),
    ("REVIEWER", "External fractional CFO or principal accountant (not the regular bookkeeper)"),
    ("TYPICAL FINDINGS", "Five to ten low or medium findings per quarter, one to two high in first quarter"),
]

INPUTS_REQUIRED = [
    ("Found answer per review question", "Checklist tab"),
    ("Findings log: description, severity, owner, target date, status", "Findings tab"),
]

CHECK_RESULTS = ["Pass", "Exception", "N/A"]
SEVERITY_OPTIONS = ["Low", "Medium", "High"]
FINDING_STATUS = ["Open", "In progress", "Resolved", "Accepted"]

# (category, question, expected)
QUESTIONS = [
    ("Chart of accounts", "Active accounts only; no duplicates by name", "No duplicate account names"),
    ("Chart of accounts", "Every account has a tax default set", "All accounts have a tax default"),
    ("Chart of accounts", "Header structure aligns with reporting pack", "Headers match reporting headers"),
    ("Contacts", "All supplier and customer contacts have ABNs where required", "ABN populated for all GST-registered counterparties"),
    ("Contacts", "No duplicate contacts by ABN", "No duplicate ABNs"),
    ("Contacts", "Inactive contacts are archived, not deleted", "Inactive flagged, not removed"),
    ("Bank reconciliation", "Every bank account reconciled to statement at period-end", "Reconciled with zero unreconciled difference"),
    ("Bank reconciliation", "No unreconciled items older than 60 days", "No items older than 60 days"),
    ("GL postings", "No manual journals to bank or AR/AP control accounts", "No manual journals to control accounts"),
    ("GL postings", "Manual journals are dated within the open period only", "All manual journals dated in open period"),
    ("GL postings", "Suspense or holding accounts are empty at period-end", "Suspense and holding accounts cleared"),
    ("Bank rules", "Active rules point to active accounts only", "All bank rules point to active accounts"),
    ("Bank rules", "No rule has been auto-applied to an unexpected counterparty", "Rules respected within tolerance"),
    ("Tracking categories", "Tracking categories applied consistently across departments", "Tracking applied to all expected accounts"),
    ("Tracking categories", "No orphan transactions missing a tracking value", "All trackable transactions tagged"),
    ("GST and BAS", "BAS-shaped report ties to GST receivable and payable balances", "GL ties to BAS report"),
    ("GST and BAS", "Capital acquisitions coded correctly", "Fixed asset GST treated as capital"),
    ("Payroll", "STP filings up to date and reconciled to GL", "STP and GL aligned"),
    ("Payroll", "Leave balances reconciled to subsidiary", "Leave subsidiary matches GL"),
    ("Fixed assets", "Asset register reconciles to GL by account", "Subsidiary equals GL by class"),
    ("Year-end carryforward", "Retained earnings and opening balances rolled forward correctly", "Opening balances tied to prior year close"),
]


def build_checklist(wb, formats):
    ws = wb.add_worksheet("Checklist")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 7
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 20)   # Category
    ws.set_column(2, 2, 50)   # Question
    ws.set_column(3, 3, 38)   # Expected
    ws.set_column(4, 4, 36)   # Found
    ws.set_column(5, 5, 14)   # Result
    ws.set_column(6, 6, 12)   # Notes
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Hygiene review",
        title="Accounting file review checklist",
        last_col=LAST_COL + 1,
        explanation=(
            "Each row is one question. Compare the Expected column to what you actually "
            "find in the file and mark the Result. Any Exception triggers a finding to "
            "log on the Findings sheet for remediation. Run this quarterly."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Step 1   Walk every category",
                            title="Review questions")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Category", "Question", "Expected", "Found", "Result", "Linked finding"],
        right_align_from=99,
    )

    r = header_row + 1
    first_data_row_1b = r + 1
    rng = d.make_rng("review")
    for i, (cat, q, expected) in enumerate(QUESTIONS):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 26)
        ws.write_string(r, 1, cat, label_fmt)
        ws.write_string(r, 2, q, label_fmt)
        ws.write_string(r, 3, expected, label_fmt)
        # Sample: mostly Pass with a few Exceptions for the example data
        result = "Exception" if rng.random() < 0.18 else "Pass"
        sample_found = "As expected" if result == "Pass" else "Variance noted; see Findings"
        ws.write_string(r, 4, sample_found, formats["input_text"])
        ws.write_string(r, 5, result, formats["input_text"])
        ws.write_blank(r, 6, None, formats["input_text"])
        r += 1
    last_data_row_1b = r

    ws.data_validation(first_data_row_1b - 1, 5, last_data_row_1b - 1, 5,
                       {"validate": "list", "source": CHECK_RESULTS})

    result_range = f"F{first_data_row_1b}:F{last_data_row_1b}"
    ws.conditional_format(result_range, {"type": "text", "criteria": "containing", "value": "Pass", "format": formats["check_status_ok"]})
    ws.conditional_format(result_range, {"type": "text", "criteria": "containing", "value": "Exception", "format": formats["check_status_flag"]})
    ws.conditional_format(result_range, {"type": "text", "criteria": "containing", "value": "N/A", "format": formats["check_status_neutral"]})

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Every question has a result recorded",
             "left": f"=COUNTA(F{first_data_row_1b}:F{last_data_row_1b})",
             "right": f"=COUNTA(C{first_data_row_1b}:C{last_data_row_1b})"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Checklist")


def build_findings(wb, formats):
    ws = wb.add_worksheet("Findings")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 8
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 6)    # Ref
    ws.set_column(2, 2, 20)   # Category
    ws.set_column(3, 3, 44)   # Finding
    ws.set_column(4, 4, 12)   # Severity
    ws.set_column(5, 5, 16)   # Owner
    ws.set_column(6, 6, 14)   # Target date
    ws.set_column(7, 7, 14)   # Status
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats, kicker="Issues log",
        title="Findings raised during the review",
        last_col=LAST_COL + 1,
        explanation=(
            "One row per finding raised during the review. Each finding has a severity, "
            "an owner, and a target resolution date. The Status field is updated as the "
            "finding is worked through."
        ),
    )

    section_row = 6
    st.write_section_header(ws, formats, row=section_row,
                            kicker="Step 2   Log every exception",
                            title="Open findings")
    header_row = section_row + 2
    st.write_header_row(
        ws, formats, row=header_row,
        headers=["Ref", "Category", "Finding", "Severity", "Owner", "Target", "Status"],
        right_align_from=99,
    )

    SAMPLE_FINDINGS = [
        ("Contacts", "12 customer contacts missing ABN where invoice is over $82.50", "Medium", "Bookkeeper", "2026-06-15", "In progress"),
        ("Bank rules", "Two rules point to account 463 Office Expenses which has been archived", "Low", "Bookkeeper", "2026-06-01", "Open"),
        ("Chart of accounts", "Duplicate Marketing accounts (461 and 462) merged in P&L pack but separate in software", "Medium", "Finance lead", "2026-06-30", "Open"),
        ("GST and BAS", "Two capital purchases coded as GST on Expenses rather than Capital Acquisition", "High", "Finance lead", "2026-05-31", "Resolved"),
    ]

    date_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "font_color": b.ACCENT_SOFT,
        "align": "right", "valign": "vcenter", "bg_color": "#FFFEF7",
        "border": 1, "border_color": b.AMBER, "num_format": "yyyy-mm-dd",
    })

    r = header_row + 1
    first_data_row_1b = r + 1
    for i, (cat, finding, sev, owner, target, status) in enumerate(SAMPLE_FINDINGS):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 26)
        ws.write_string(r, 1, f"F{i + 1:03d}", label_fmt)
        ws.write_string(r, 2, cat, label_fmt)
        ws.write_string(r, 3, finding, label_fmt)
        ws.write_string(r, 4, sev, formats["input_text"])
        ws.write_string(r, 5, owner, formats["input_text"])
        from datetime import date as _date
        y, m, d_ = [int(x) for x in target.split("-")]
        ws.write_datetime(r, 6, _date(y, m, d_), date_fmt)
        ws.write_string(r, 7, status, formats["input_text"])
        r += 1
    # Leave 6 blank rows for additional findings
    for i in range(6):
        zebra = (len(SAMPLE_FINDINGS) + i) % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.set_row(r, 22)
        ws.write_string(r, 1, f"F{len(SAMPLE_FINDINGS) + i + 1:03d}", label_fmt)
        for c in range(2, 4):
            ws.write_blank(r, c, None, formats["input_text"])
        ws.write_blank(r, 4, None, formats["input_text"])
        ws.write_blank(r, 5, None, formats["input_text"])
        ws.write_blank(r, 6, None, formats["input_text"])
        ws.write_blank(r, 7, None, formats["input_text"])
        r += 1
    last_data_row_1b = r

    ws.data_validation(first_data_row_1b - 1, 4, last_data_row_1b - 1, 4,
                       {"validate": "list", "source": SEVERITY_OPTIONS})
    ws.data_validation(first_data_row_1b - 1, 7, last_data_row_1b - 1, 7,
                       {"validate": "list", "source": FINDING_STATUS})

    sev_range = f"E{first_data_row_1b}:E{last_data_row_1b}"
    ws.conditional_format(sev_range, {"type": "text", "criteria": "containing", "value": "High", "format": formats["check_status_flag"]})
    ws.conditional_format(sev_range, {"type": "text", "criteria": "containing", "value": "Medium", "format": formats["check_status_neutral"]})
    ws.conditional_format(sev_range, {"type": "text", "criteria": "containing", "value": "Low", "format": formats["check_status_ok"]})

    status_range = f"H{first_data_row_1b}:H{last_data_row_1b}"
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "Resolved", "format": formats["check_status_ok"]})
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "Accepted", "format": formats["check_status_ok"]})
    ws.conditional_format(status_range, {"type": "text", "criteria": "containing", "value": "Open", "format": formats["check_status_flag"]})

    r += 2
    st.write_checks_block(
        ws, formats, row=r,
        checks=[
            {"name": "Number of Exception results on Checklist equals findings logged",
             "left": "=COUNTIF(Checklist!F:F,\"Exception\")",
             "right": f"=COUNTA(D{first_data_row_1b}:D{last_data_row_1b})"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Findings")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE,
                       workbook_kicker=WORKBOOK_KICKER,
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_checklist(wb, formats)
    build_findings(wb, formats)
    sc.add_connect_data_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
