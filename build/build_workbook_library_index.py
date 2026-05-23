"""Workbook Library index (internal).

Single-sheet, filterable table mapping every numbered workbook in the
library to its band, title, short description, when-to-use copy, status,
and on-disk filename. Internal reference only, not customer-facing, so
no cover sheet or branded hero band is rendered.

Also emits a JSON sidecar with the same content for the website to
consume when rendering workbook cards.

Output:
    C:\\dev\\lyros-workbooks\\library\\Workbook Library.xlsx
    C:\\dev\\lyros-workbooks\\library\\workbook_library.json
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import xlsxwriter

sys.path.insert(0, str(Path(__file__).parent))
from _shared import branding as b


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\Workbook Library.xlsx")
JSON_PATH = OUTPUT_PATH.parent / "workbook_library.json"


# Each row is (number, band, title, short_desc, when_to_use, status).
# Status is one of {"built", "pending"}. Written in Lyros advisory voice:
# no contractions, no em-dashes, Australian English.
ROWS: list[tuple[str, str, str, str, str, str]] = [
    ("1010", "1000s Engagement and setup", "Chart of Accounts Setup",
     "Setup register for the chart of accounts your bookkeeper will load into your software.",
     "A clean chart of accounts is the foundation of everything downstream. Most "
     "businesses inherit a chart that has grown organically over years, with duplicate "
     "accounts, dead codes, and inconsistent naming. We use this workbook at the start "
     "of an engagement to design the chart from scratch around how the business actually "
     "reads its numbers. The header accounts mirror the management reporting pack so "
     "monthly reads come out tidy with no remapping needed. Where finance teams struggle "
     "is making the structural decisions in isolation, for example whether marketing "
     "salaries sit in wages or in marketing, or whether COGS gets sub-accounts per "
     "channel. Those choices ripple through every workbook in this library, so we will "
     "usually scope a one-hour call at this point to lock in the structure with you "
     "before you load it into the software.",
     "built"),

    ("1020", "1000s Engagement and setup", "Supplier to COA Mapping",
     "Maps each supplier to a default account, GST treatment, and approval rule.",
     "We use chart-of-accounts mapping to supplier codes to automate the rule for what "
     "account each incoming invoice maps to in the accounting software. Setting these "
     "rules up eliminates the need to consider each invoice individually, which saves "
     "processing time and improves coding consistency. A rules-based approach works "
     "extremely well as an Excel automation in its own right. If you want to enhance "
     "this further, you can layer an AI-driven automation that reviews how many "
     "transactions per supplier are actually allocated to each account on a rolling "
     "twelve-month basis. The output will often show that 85 per cent of transactions "
     "for a supplier go to one account and 15 per cent to another, or that the split is "
     "100 per cent to one account, or evenly 50-50. That review reduces how often you "
     "need to re-assess the mapping, which should still happen on a cadence suited to "
     "your business or at least annually as a minimum standard.",
     "built"),

    ("2010", "2000s Processes and risks", "Month-End Close Checklist",
     "Step-by-step close checklist with sign-off, ageing, and reconciliation gates.",
     "Month-end close drifts when the checklist lives in someone's head. We see closes "
     "that should take three days stretch to nine because two reconciliations were "
     "missed and the variance review only picked it up after the management pack went "
     "out. This workbook is the canonical close checklist with sign-off gates and "
     "ageing thresholds. Use it monthly, tick each step as you complete it, and the "
     "workbook flags any reconciliation that is still open past the sign-off threshold. "
     "Operating the checklist sits within most teams' skillset; the discipline tends to "
     "slip during busy periods, which is when the workbook earns its keep. Lyros runs "
     "this checklist remotely on the same cadence as your bookkeeper for clients who "
     "want the extra discipline without hiring an extra finance head.",
     "built"),

    ("2020", "2000s Processes and risks", "Accounting File Review Checklist",
     "Quarterly review of account balances, posting hygiene, and lookup mappings.",
     "A quarterly check on the integrity of your accounting file. This catches the "
     "small things that compound, like a misconfigured GST default on a new account, "
     "contacts saved without ABNs, or supplier rules pointing at accounts that have "
     "since been archived. None of these break the file on day one, but they corrupt "
     "the data the management reporting pack reads. Within reach of an experienced "
     "bookkeeper to run, but most owner-operators do not have the time or the "
     "audit-style mindset to work through it. We include this checklist quarterly as a "
     "standard line item in our retainer engagements.",
     "built"),

    ("4000", "4000s Financials and pervasive workings", "Management Reporting Pack",
     "Monthly P&L pack with headline, monthly trend, quarter comparison, wages, and working capital.",
     "The headline workbook for monthly reporting to a leadership team or board. It "
     "pulls together a one-page summary, a monthly P&L trend, quarter-over-quarter "
     "comparison, wages breakdown, and working capital position. The pack is designed "
     "to be readable in five minutes by a non-finance founder. Building this from "
     "scratch takes a strong finance lead about a day per month once the template is "
     "set up. We typically run this for clients on a fractional CFO basis, which "
     "includes preparing the commentary and walking the leadership team through the "
     "narrative. The commentary is the part that creates value rather than the "
     "spreadsheet itself.",
     "built"),

    ("4001", "4000s Financials and pervasive workings", "Board Reporting Pack",
     "One-page board read-out plus detail, cash position, and commentary.",
     "Where the management reporting pack is built for an internal leadership team, the "
     "board pack is built for a board of directors who want governance-grade reporting "
     "with a longer narrative. Adds a one-page summary with commentary, a detail tab, a "
     "cash position tab, and a structured commentary block. Use this when you have an "
     "investor, advisory board, or governance-focused board member who needs reporting "
     "with a real narrative arc. Most finance leads can produce the figures; the "
     "commentary takes more practice, and our fractional CFO service includes "
     "preparation of the commentary as part of the engagement.",
     "built"),

    ("4010", "4000s Financials and pervasive workings", "Budget vs Actual",
     "Side-by-side actuals and budget with monthly and YTD variance.",
     "A side-by-side actuals and budget view with monthly and year-to-date variance. "
     "Most businesses build something like this in their accounting software's "
     "reporting module, but the limitations show quickly: monthly drill-down is "
     "shallow, you cannot easily add context, and reforecasting requires rebuilding the "
     "budget. This template separates the budget table from the actual table on the "
     "Data sheet, so you can reforecast without losing your original budget. When "
     "variance hits material thresholds the value is in explaining why, not in "
     "producing the variance number itself. That commentary is the typical scope "
     "boundary between in-house bookkeeping and a fractional CFO.",
     "built"),

    ("4020", "4000s Financials and pervasive workings", "Monthly Variance and YTD Bridge",
     "Walk-the-bridge from prior-month and YTD variance drivers to current actuals.",
     "Builds on the budget vs actual workbook to show what actually drove the variance "
     "month by month. Useful when leadership ask why we are behind and the answer is "
     "not one large item but five small ones compounded. Walks the bridge from "
     "prior-month variance drivers to current-month actuals, then collapses the same "
     "logic for year-to-date. This is squarely fractional CFO territory. Most finance "
     "leads can describe variance; few decompose it into clean driver categories on a "
     "monthly cadence with the same definitions each time.",
     "built"),

    ("4030", "4000s Financials and pervasive workings", "Departmental Variance Rollup",
     "Department-level P&L rollup with variance to budget per cost centre.",
     "For businesses with multiple departments or cost centres, this rolls each "
     "department's P&L up into a consolidated variance view. Most accounting software "
     "supports this natively if you have tracking categories or classes configured, but "
     "the standard output is narrow. This workbook flips it horizontally so you can see "
     "all departments side by side. Most useful where department heads are accountable "
     "to their own P&L and want to see their own line items alongside the group "
     "consolidation in the same view.",
     "built"),

    ("4040", "4000s Financials and pervasive workings", "Budgeting Scenario Flex",
     "Driver-based scenario flex (base, upside, downside) for forward 12 months.",
     "Driver-based forecast with base, upside, and downside scenarios across the next "
     "twelve months. Building this from scratch is harder than it looks. You need to "
     "identify which drivers actually matter, decide on sensitivity ranges that are not "
     "made up, and structure the workbook so changing one driver does not break the "
     "others. We use this in a fractional CFO context when a business is considering a "
     "hire, a property move, a price change, or a capital raise. The workbook is the "
     "working file; the value is in the assumptions and the conversation that surfaces "
     "them.",
     "built"),

    ("4050", "4000s Financials and pervasive workings", "Consolidation from Trial Balances",
     "Per-entity trial balances with eliminations and combined consolidated P&L and BS.",
     "Where a group has multiple entities and the accounting software does not "
     "consolidate them automatically (or does it poorly), this workbook accepts a trial "
     "balance per entity and produces a combined consolidated profit and loss and "
     "balance sheet with intercompany eliminations. Most useful for groups using "
     "different accounting platforms across entities, or where one entity is on a "
     "different chart of accounts. The technical skill required is journal-level "
     "accounting: you need to know what eliminates and what does not. This is beyond "
     "most owner-operators and firmly in firm-of-account or fractional CFO scope.",
     "built"),

    ("5100", "5000s Assets", "13 Week Rolling Cash Flow",
     "Forward 13-week cash flow with inflows, outflows, and opening and closing balances.",
     "Forward thirteen-week cash flow forecast with inflows, outflows, and opening and "
     "closing balances. The standard tool for businesses under cash pressure, but "
     "useful for every business at month-end. The skill required is not the workbook "
     "itself, it is being honest about what is actually likely to be paid by week and "
     "what is going to slip. Owner-operators tend to be optimistic about collections; "
     "bookkeepers tend to underestimate the upside. We run this workbook with clients "
     "on a weekly cadence when cash is tight and on a monthly cadence when it is not. "
     "The conversation around the file is usually more valuable than the file itself.",
     "built"),

    ("5110", "5000s Assets", "Cash Position by Bank",
     "Daily cash position split by bank account with reconciliation to the GL.",
     "Daily cash position split by bank account with reconciliation back to the general "
     "ledger. Useful where a business operates multiple bank accounts for separation of "
     "duties, or where there is a tax set-aside account that should not be touched for "
     "operations. Operationally straightforward once the bank feeds are configured; the "
     "workbook is mostly an aggregator and a reconciliation. We use this with clients "
     "who have outgrown the single-account simplicity but have not yet hired a "
     "full-time finance person who can hold the picture in their head.",
     "built"),

    ("5120", "5000s Assets", "GST and BAS Cash Flow Timing",
     "Cash impact of upcoming BAS payments and refunds across the next four quarters.",
     "Models the cash impact of upcoming BAS payments and refunds across the next four "
     "quarters. The naive view of BAS is that it is paid quarterly out of operating "
     "cash; the sophisticated view is that GST is held in trust during the quarter and "
     "the BAS payment is just settling up a liability that has been accruing all "
     "along. This workbook treats it the second way, which means it tells you whether "
     "you are short on cash to actually pay the BAS rather than just telling you that "
     "the BAS is due. We use this with clients who have surprised themselves with a "
     "BAS bill in the past.",
     "built"),

    ("5150", "5000s Assets", "Working Capital Cycle",
     "DSO, DPO, DIO, and cash conversion cycle trend with rolling 12-month chart.",
     "Calculates days sales outstanding, days payable outstanding, days inventory "
     "outstanding, and the resulting cash conversion cycle over a rolling twelve-month "
     "window. A standard finance-team workbook, but rarely run with the rigour needed "
     "to actually drive change. The trick is to compare the cycle against the cash on "
     "hand and ask whether tightening the cycle by ten days would meaningfully reduce "
     "the working capital tied up in operations. Most owner-operators understand the "
     "concept; few have the data discipline to run it monthly. We run this quarterly "
     "with clients on a retainer.",
     "built"),

    ("5200", "5000s Assets", "Aged Receivables Chase List",
     "Customer-level ageing buckets with follow-up tier and priority for the credit team.",
     "A customer-level ageing bucket with follow-up tier and priority for the credit "
     "team. The accounting software will produce an ageing report; what it does not "
     "produce is a prioritised follow-up list that respects the relationship with each "
     "customer. This workbook adds a tier (gentle reminder, second reminder, escalate, "
     "legal) and a priority based on amount and ageing. Operationally within reach of "
     "a strong bookkeeper; the scope creep is when collections start to affect customer "
     "relationships, at which point a CFO-level conversation usually helps determine "
     "where the line sits.",
     "built"),

    ("5210", "5000s Assets", "Aged Receivables Analysis",
     "Heatmap of customer ageing exposure, concentration, and top-debtor drill-down.",
     "Where the chase list is operational, the analysis is strategic. A heatmap of "
     "customer ageing exposure, concentration risk, and drill-down on the top debtors. "
     "Use this monthly to spot a debtor that has crept into the 60-day bucket twice in "
     "a row, which is usually a leading indicator of a payment problem worth flagging "
     "early. Most leadership teams look at total receivables; few look at the "
     "distribution of receivables and where the risk is concentrated. The analysis "
     "tends to sit in fractional CFO scope.",
     "built"),

    ("5300", "5000s Assets", "Depreciation Roll Forward",
     "Asset register with WDV opening, additions, disposals, depreciation, and closing balances.",
     "Asset register with written-down-value opening, additions, disposals, "
     "depreciation, and closing balances. The accounting software handles depreciation "
     "journal entries but the register itself often lives in a spreadsheet because the "
     "software cannot model things like fixed-asset disposals with partial-year "
     "depreciation cleanly. This workbook is the working register. Within reach of a "
     "bookkeeper with fixed-asset experience; we tend to step in at year-end for the "
     "depreciation schedule reconciliation that supports the income tax return.",
     "built"),

    ("6100", "6000s Liabilities and tax", "GST and BAS Lodgement Tracker",
     "Quarterly lodgement and payment register with due dates, amounts, and status.",
     "A quarterly lodgement and payment register with due dates, amounts paid, and "
     "lodgement status. Most useful for businesses with multiple GST-registered "
     "entities, or where the BAS is prepared by an external accountant and the "
     "operator wants to track lodgement and payment dates centrally across the group. "
     "Operationally simple; the workbook is a tracker rather than a model. We use it "
     "with multi-entity clients to give the operator one place to see what has been "
     "lodged and paid.",
     "built"),

    ("6200", "6000s Liabilities and tax", "FBT Exposure Quick Check",
     "Common FBT triggers screened against transactional data with exposure estimate.",
     "Common fringe benefits tax triggers screened against transactional data with an "
     "exposure estimate. FBT is one of those taxes where the worst outcome is "
     "discovering a year later that you have been triggering it without realising. The "
     "workbook is a screening tool, not an FBT return. Where it flags exposure, the "
     "next step is a conversation with a tax accountant. Useful for businesses with "
     "company vehicles, entertainment spending, or employee-paid expenses that could "
     "fall in scope.",
     "built"),

    ("6300", "6000s Liabilities and tax", "Wages and Leave Liability",
     "Period-end leave balance and accrued wages liability per employee.",
     "Period-end leave balance and accrued wages liability per employee. The "
     "accounting software calculates leave balances, but the period-end accrued wages "
     "number rarely appears anywhere cleanly. This workbook is the working file for "
     "the journal that recognises wages earned but not yet paid at month-end. "
     "Bookkeeper-level skill once the structure is set up; we tend to step in when "
     "the leave provision starts being material enough to the balance sheet to "
     "warrant a more formal review.",
     "built"),

    ("7000", "7000s Revenue", "Top Customers by Revenue",
     "Ranked customer revenue with year-on-year and concentration share.",
     "Ranked customer revenue with year-on-year change and concentration share. The "
     "accounting software shows top customers by revenue; what it does not show is "
     "concentration risk. If sixty per cent of your revenue comes from three "
     "customers, that is a finding worth surfacing to a board. This workbook "
     "calculates and visualises that concentration. Within reach of any finance lead "
     "to run; the value is in the conversation about concentration risk and customer "
     "diversification that should follow.",
     "built"),

    ("7010", "7000s Revenue", "Revenue and Margin Bridge",
     "Account-level revenue decomposition between prior and current quarter.",
     "Decomposes the change in revenue and gross margin between a prior quarter and "
     "the current quarter by individual revenue account. Useful when revenue moves and "
     "the leadership team wants to know whether the movement was price, volume, mix, "
     "or a new line item. The technical work is the decomposition logic; what takes "
     "practice is choosing which decomposition matters for your business model. "
     "Fractional CFO scope for most businesses; for product-led businesses with strong "
     "internal data operations, an in-house FP&A function might own this work.",
     "built"),

    ("8000", "8000s Expenses", "Top Supplier Review",
     "Ranked supplier spend with ABN-based dedup and year-on-year change.",
     "Ranked supplier spend with ABN-based deduplication and year-on-year change. Most "
     "accounting software treats every supplier contact as a separate entity even when "
     "the same legal entity is set up twice (one for invoices, one for purchase "
     "orders, for example). This workbook deduplicates on ABN, which produces a more "
     "honest view of where the money is going. Useful when negotiating with key "
     "suppliers or running an annual supplier review. Within reach of a finance lead; "
     "the cost savings tend to emerge from the conversation that follows, not from the "
     "workbook itself.",
     "built"),

    ("8100", "8000s Expenses", "Payroll to Revenue Ratio",
     "Payroll cost as a share of revenue over time, by department.",
     "Payroll cost as a share of revenue over time, broken down by department. One of "
     "the simplest leading indicators of operating health: if the ratio creeps up "
     "while revenue is flat, you are heading into trouble. The numerator is "
     "straightforward; the denominator is too if revenue is recognised cleanly. The "
     "skill is in interpreting it. Does a payroll ratio of thirty-eight per cent mean "
     "you are overpaying staff, or that your pricing is too low? That is a fractional "
     "CFO conversation, not a workbook conclusion.",
     "built"),
]


HEADERS = [
    "Number", "Band", "Workbook", "Description",
    "When would you use this?", "Status", "Filename",
]


def _estimate_row_height(text: str, chars_per_line: int = 95) -> int:
    if not text:
        return 18
    lines = max(2, -(-len(text) // chars_per_line))
    return max(40, lines * 15 + 6)


def build_workbook():
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))

    header_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": b.WHITE, "bg_color": b.GREY_BLACK,
        "align": "left", "valign": "vcenter", "indent": 1,
        "border": 1, "border_color": b.BORDER_GREY,
    })
    cell_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10,
        "font_color": b.GREY_BLACK, "bg_color": b.WHITE,
        "align": "left", "valign": "top", "text_wrap": True,
        "border": 1, "border_color": b.BORDER_GREY, "indent": 1,
    })
    when_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10,
        "font_color": b.GREY_BLACK, "bg_color": b.WHITE,
        "align": "left", "valign": "top", "text_wrap": True,
        "border": 1, "border_color": b.BORDER_GREY, "indent": 1,
    })
    status_built_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": "#1F5B3F", "bg_color": b.TINT_GOOD,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": b.BORDER_GREY,
    })
    status_pending_fmt = wb.add_format({
        "font_name": b.FONT_FAMILY, "font_size": 10, "bold": True,
        "font_color": b.TEXT_DIM, "bg_color": b.WHITE,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": b.BORDER_GREY,
    })

    ws = wb.add_worksheet("Index")

    ws.set_column(0, 0, 8)    # Number
    ws.set_column(1, 1, 30)   # Band
    ws.set_column(2, 2, 32)   # Workbook
    ws.set_column(3, 3, 44)   # Description
    ws.set_column(4, 4, 80)   # When would you use this
    ws.set_column(5, 5, 12)   # Status
    ws.set_column(6, 6, 48)   # Filename

    # Header row
    ws.set_row(0, 28)
    for c, h in enumerate(HEADERS):
        ws.write_string(0, c, h, header_fmt)

    # Data rows
    sorted_rows = sorted(ROWS, key=lambda r: r[0])
    for i, (num, band, title, desc, when_to_use, status) in enumerate(sorted_rows):
        r = i + 1
        ws.set_row(r, _estimate_row_height(when_to_use))
        ws.write_string(r, 0, num, cell_fmt)
        ws.write_string(r, 1, band, cell_fmt)
        ws.write_string(r, 2, title, cell_fmt)
        ws.write_string(r, 3, desc, cell_fmt)
        ws.write_string(r, 4, when_to_use, when_fmt)
        status_label = "Available" if status == "built" else "Pending"
        ws.write_string(
            r, 5, status_label,
            status_built_fmt if status == "built" else status_pending_fmt,
        )
        ws.write_string(r, 6, f"{num} - {title}.xlsx", cell_fmt)

    last_row = len(sorted_rows)
    ws.autofilter(0, 0, last_row, len(HEADERS) - 1)
    ws.freeze_panes(1, 0)
    ws.set_zoom(100)

    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


def write_json_sidecar():
    items = []
    for num, band, title, desc, when_to_use, status in sorted(ROWS, key=lambda r: r[0]):
        items.append({
            "number": num,
            "band": band,
            "title": title,
            "description": desc,
            "when_to_use": when_to_use,
            "status": status,
            "filename": f"{num} - {title}.xlsx",
        })
    JSON_PATH.write_text(
        json.dumps({"items": items}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"Wrote: {JSON_PATH}")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    build_workbook()
    write_json_sidecar()


if __name__ == "__main__":
    main()
