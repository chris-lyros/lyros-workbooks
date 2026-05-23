"""Build the Group Consolidation from Trial Balances utility workbook.

Different Data shape than the rest of batch 1: the user pastes a Trial
Balance per entity (account code, account name, debit, credit, account type),
the workbook eliminates intercompany balances via a user-edited Eliminations
table, and produces a consolidated P&L and Balance Sheet.

Output:
  C:\\dev\\lyros-workbooks\\library\\utility
        \\lyros_lib_utility_consolidation_from_tbs.xlsx
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\4050 - Consolidation from Trial Balances.xlsx")

WORKBOOK_ID = "4050"
WORKBOOK_TITLE = "Group Consolidation from Trial Balances"
WORKBOOK_KICKER = "Multi-entity consolidation utility"
TARGET_USER = (
    "CFO or finance lead of a small group with two to five entities preparing a "
    "consolidated profit and loss and balance sheet."
)
HOW_TO_USE = [
    "Open the Entity TBs sheet and paste each entity's Trial Balance into its column block (Debit / Credit per account).",
    "Open the Eliminations sheet and enter the intercompany journals that net to zero at the group level (e.g. intercompany loans, sales between entities).",
    "The Consolidated P&L and Consolidated Balance Sheet recalculate automatically and tie-out checks confirm the balance sheet balances and intercompany pairs net.",
]

EXAMPLE_PROFILE = [
    ("STRUCTURE", "Group with three entities: Parent, Subsidiary 1, Subsidiary 2"),
    ("INTERCOMPANY", "Loans between Parent and Subsidiaries, plus internal sales"),
    ("REPORTING DATE", "Most recent month-end"),
    ("OUTPUT", "Consolidated P&L and Balance Sheet ready for board or audit"),
]

INPUTS_REQUIRED = [
    ("Trial Balance per entity", "Entity TBs tab (paste from each entity's TB export from your accounting software)"),
    ("Account to Report line mapping", "Entity TBs tab (dropdown in Report line column)"),
    ("Intercompany eliminations", "Eliminations tab (manual journals that net to zero)"),
]


# Layout: 5 entity columns supported
ENTITIES = ["Parent", "Subsidiary 1", "Subsidiary 2"]
NUM_ENTITIES = len(ENTITIES)

# Trial balance accounts. Mix of P&L and BS accounts.
TB_ACCOUNTS = [
    # code, name, report_line, bs_or_pl, default_dr_value, default_cr_value (per entity)
    # Revenue accounts (CR balance)
    ("200", "Sales - Wholesale",       "Revenue",       "P&L"),
    ("210", "Sales - Retail",          "Revenue",       "P&L"),
    ("220", "Sales - Online",          "Revenue",       "P&L"),
    ("290", "Intercompany Sales",      "IC Revenue",    "P&L"),  # eliminated
    # COGS (DR balance)
    ("310", "Cost of Goods Sold",      "Cost of sales", "P&L"),
    # Expenses
    ("477", "Wages and Salaries",      "Wages",         "P&L"),
    ("469", "Rent",                    "Other opex",    "P&L"),
    ("466", "Accounting Fees",         "Other opex",    "P&L"),
    ("416", "Depreciation",            "D&A",           "P&L"),
    # Balance sheet assets (DR)
    ("610", "Cash and equivalents",    "Cash",          "BS"),
    ("620", "Trade Receivables",       "Receivables",   "BS"),
    ("625", "Intercompany Receivable", "IC Asset",      "BS"),  # eliminated
    ("630", "Inventory",               "Inventory",     "BS"),
    ("700", "Property Plant Equipment","PPE",           "BS"),
    # Balance sheet liabilities (CR)
    ("800", "Trade Payables",          "Payables",      "BS"),
    ("810", "Intercompany Payable",    "IC Liability",  "BS"),  # eliminated
    ("820", "Loans",                   "Loans",         "BS"),
    # Equity (CR)
    ("900", "Share Capital",           "Equity",        "BS"),
    ("910", "Retained Earnings",       "Equity",        "BS"),
]

REPORT_LINES = [
    "Revenue", "IC Revenue", "Cost of sales", "Wages", "Other opex", "D&A",
    "Cash", "Receivables", "IC Asset", "Inventory", "PPE",
    "Payables", "IC Liability", "Loans", "Equity",
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


# Entity TBs sheet layout (0-based columns)
TB_COL_CODE = 1   # B
TB_COL_NAME = 2   # C
TB_COL_LINE = 3   # D
TB_COL_TYPE = 4   # E
# Per entity: Debit, Credit pair
# Parent: F (DR), G (CR)
# Sub 1: H (DR), I (CR)
# Sub 2: J (DR), K (CR)
TB_COL_FIRST_ENTITY = 5

# Total columns at the end: combined DR, combined CR, net
TB_COL_COMBINED_DR = TB_COL_FIRST_ENTITY + NUM_ENTITIES * 2          # = 11 (L)
TB_COL_COMBINED_CR = TB_COL_COMBINED_DR + 1                          # = 12 (M)
TB_COL_NET = TB_COL_COMBINED_CR + 1                                  # = 13 (N)

TB_ROW_HEADER = 9
TB_ROW_FIRST = 10
TB_ROW_LAST = TB_ROW_FIRST + len(TB_ACCOUNTS) - 1

# Eliminations sheet layout
ELIM_ROW_HEADER = 8
ELIM_ROW_FIRST = 9
ELIM_NUM_ROWS = 5  # 5 elimination journals supported
ELIM_ROW_LAST = ELIM_ROW_FIRST + ELIM_NUM_ROWS - 1


def _synthetic_data() -> dict:
    """Build synthetic TBs per entity. Each TB is balanced (debits = credits)."""
    rng = d.make_rng("cons-v1")

    # Build TBs as dict of {account_code: (dr, cr)} per entity
    entity_tbs: list[dict[str, tuple[float, float]]] = []
    for ent_idx in range(NUM_ENTITIES):
        tb: dict[str, tuple[float, float]] = {}
        scale = [1.0, 0.55, 0.30][ent_idx]  # Parent is largest, subs are smaller

        # Synthetic figures keyed to scale
        revenue_total = round(2_400_000 * scale, 0)
        cogs_total = round(revenue_total * 0.56, 0)
        wages_total = round(revenue_total * 0.20, 0)
        rent_total = round(revenue_total * 0.04, 0)
        acct_total = round(revenue_total * 0.02, 0)
        dep_total = round(revenue_total * 0.03, 0)
        ic_sales = round(revenue_total * 0.05, 0) if ent_idx == 0 else (round(revenue_total * 0.08, 0) if ent_idx == 1 else 0)

        # P&L (Revenue: CR, expenses: DR)
        tb["200"] = (0, round(revenue_total * 0.50, 0))
        tb["210"] = (0, round(revenue_total * 0.35, 0))
        tb["220"] = (0, round(revenue_total * 0.15, 0))
        tb["290"] = (0, ic_sales)
        tb["310"] = (cogs_total, 0)
        tb["477"] = (wages_total, 0)
        tb["469"] = (rent_total, 0)
        tb["466"] = (acct_total, 0)
        tb["416"] = (dep_total, 0)

        # NPAT for the year (revenue - all expenses)
        npat = (revenue_total + ic_sales) - cogs_total - wages_total - rent_total - acct_total - dep_total

        # BS (assets: DR, liabilities: CR, equity: CR)
        cash = round(450_000 * scale + rng.uniform(-40_000, 40_000), 0)
        receivables = round(280_000 * scale + rng.uniform(-20_000, 20_000), 0)
        ic_asset = round(220_000, 0) if ent_idx == 0 else 0  # Parent has IC receivable
        inventory = round(190_000 * scale + rng.uniform(-15_000, 15_000), 0)
        ppe = round(800_000 * scale + rng.uniform(-30_000, 30_000), 0)

        payables = round(190_000 * scale + rng.uniform(-15_000, 15_000), 0)
        ic_liability = round(110_000, 0) if ent_idx == 1 else (round(110_000, 0) if ent_idx == 2 else 0)
        loans = round(400_000 * scale + rng.uniform(-20_000, 20_000), 0)

        # Balance the sheet: equity = total assets - total liabilities - NPAT
        total_assets = cash + receivables + ic_asset + inventory + ppe
        total_liab = payables + ic_liability + loans
        # Retained earnings rolls in NPAT; share capital is the residual to balance
        retained = round(300_000 * scale, 0)
        share_capital = total_assets - total_liab - retained - npat

        tb["610"] = (cash, 0)
        tb["620"] = (receivables, 0)
        tb["625"] = (ic_asset, 0)
        tb["630"] = (inventory, 0)
        tb["700"] = (ppe, 0)
        tb["800"] = (0, payables)
        tb["810"] = (0, ic_liability)
        tb["820"] = (0, loans)
        tb["900"] = (0, share_capital)
        tb["910"] = (0, retained)

        entity_tbs.append(tb)

    # Eliminations: a few standard intercompany journals
    eliminations = [
        ("Eliminate IC sales (Parent → Sub1)", "IC Revenue", "Cost of sales", 192_000),
        ("Eliminate IC sales (Parent → Sub2)", "IC Revenue", "Cost of sales", 36_000),
        ("Eliminate IC receivable/payable (Parent ↔ Sub1)", "IC Asset", "IC Liability", 110_000),
        ("Eliminate IC receivable/payable (Parent ↔ Sub2)", "IC Asset", "IC Liability", 110_000),
        ("", "", "", 0),  # blank row for user to extend
    ]

    return {"entity_tbs": entity_tbs, "eliminations": eliminations}


def build_entity_tbs(wb, formats, dat):
    ws = wb.add_worksheet("Entity TBs")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = TB_COL_NET + 1
    ws.set_column(0, 0, 2)
    ws.set_column(TB_COL_CODE, TB_COL_CODE, 8)
    ws.set_column(TB_COL_NAME, TB_COL_NAME, 26)
    ws.set_column(TB_COL_LINE, TB_COL_LINE, 14)
    ws.set_column(TB_COL_TYPE, TB_COL_TYPE, 6)
    # 3 entities × 2 columns each (DR, CR) at width 11
    ws.set_column(TB_COL_FIRST_ENTITY, TB_COL_FIRST_ENTITY + NUM_ENTITIES * 2 - 1, 11)
    ws.set_column(TB_COL_COMBINED_DR, TB_COL_NET, 12)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Single source of truth",
        title="Trial balance per entity",
        last_col=LAST_COL + 1,
        explanation=(
            "Paste each entity's Trial Balance into its column block. Debits and Credits "
            "are entered separately to mirror a TB export from your accounting software. The Combined columns sum "
            "across all entities; the Net column is Combined DR minus Combined CR. The "
            "Consolidated sheets read from the Net column filtered by Report line."
        ),
    )

    st.write_section_header(
        ws, formats, row=6,
        kicker="Step 1   Paste each entity's TB",
        title="Trial balances",
    )

    # Two-row header: entity name spans DR/CR pair on row 8, then DR/CR labels on row 9
    # First, write the top header row at row 7 (0-based) which is 8 in 1-based
    ws.set_row(TB_ROW_HEADER - 2, 22)
    ws.set_row(TB_ROW_HEADER - 1, 22)
    # Static labels
    for col_idx, label in [(TB_COL_CODE, "Code"), (TB_COL_NAME, "Account name"),
                            (TB_COL_LINE, "Report line"), (TB_COL_TYPE, "Type")]:
        ws.merge_range(TB_ROW_HEADER - 2, col_idx, TB_ROW_HEADER - 1, col_idx, label, formats["th"])

    # Per-entity DR/CR header spans
    for i, ent in enumerate(ENTITIES):
        dr_col = TB_COL_FIRST_ENTITY + i * 2
        cr_col = dr_col + 1
        ws.merge_range(TB_ROW_HEADER - 2, dr_col, TB_ROW_HEADER - 2, cr_col, ent, formats["th"])
        ws.write_string(TB_ROW_HEADER - 1, dr_col, "Debit", formats["th_right"])
        ws.write_string(TB_ROW_HEADER - 1, cr_col, "Credit", formats["th_right"])

    # Combined columns
    ws.merge_range(TB_ROW_HEADER - 2, TB_COL_COMBINED_DR, TB_ROW_HEADER - 2, TB_COL_NET, "Combined", formats["th"])
    ws.write_string(TB_ROW_HEADER - 1, TB_COL_COMBINED_DR, "Total DR", formats["th_right"])
    ws.write_string(TB_ROW_HEADER - 1, TB_COL_COMBINED_CR, "Total CR", formats["th_right"])
    ws.write_string(TB_ROW_HEADER - 1, TB_COL_NET, "Net", formats["th_right"])

    # Data rows
    for i, (code, name, report_line, bs_or_pl) in enumerate(TB_ACCOUNTS):
        row_0 = TB_ROW_FIRST - 1 + i
        zebra = i % 2 == 0
        text_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.set_row(row_0, 20)
        ws.write_string(row_0, TB_COL_CODE, code, text_fmt)
        ws.write_string(row_0, TB_COL_NAME, name, text_fmt)
        ws.write_string(row_0, TB_COL_LINE, report_line, formats["input_text"])
        ws.write_string(row_0, TB_COL_TYPE, bs_or_pl, text_fmt)

        for ent_idx, ent in enumerate(ENTITIES):
            dr_col = TB_COL_FIRST_ENTITY + ent_idx * 2
            cr_col = dr_col + 1
            dr, cr = dat["entity_tbs"][ent_idx].get(code, (0, 0))
            ws.write_number(row_0, dr_col, dr, formats["input_value"])
            ws.write_number(row_0, cr_col, cr, formats["input_value"])

        # Combined DR / CR / Net formulas
        dr_cells = [
            f"${col_letter(TB_COL_FIRST_ENTITY + k * 2)}${TB_ROW_FIRST + i}"
            for k in range(NUM_ENTITIES)
        ]
        cr_cells = [
            f"${col_letter(TB_COL_FIRST_ENTITY + k * 2 + 1)}${TB_ROW_FIRST + i}"
            for k in range(NUM_ENTITIES)
        ]
        ws.write_formula(row_0, TB_COL_COMBINED_DR, f"={'+'.join(dr_cells)}", num_fmt)
        ws.write_formula(row_0, TB_COL_COMBINED_CR, f"={'+'.join(cr_cells)}", num_fmt)
        ws.write_formula(row_0, TB_COL_NET,
                         f"=${col_letter(TB_COL_COMBINED_DR)}${TB_ROW_FIRST + i}"
                         f"-${col_letter(TB_COL_COMBINED_CR)}${TB_ROW_FIRST + i}",
                         num_fmt)

    ws.data_validation(TB_ROW_FIRST - 1, TB_COL_LINE, TB_ROW_LAST - 1, TB_COL_LINE,
                       {"validate": "list", "source": REPORT_LINES})

    # Sanity-check row at the bottom: each entity's DR total should equal its CR total
    sc_row = TB_ROW_LAST + 1
    ws.set_row(sc_row, 22)
    ws.write_string(sc_row, TB_COL_NAME, "Entity TB DR vs CR check", formats["td_bold_left"])
    for ent_idx, ent in enumerate(ENTITIES):
        dr_col = TB_COL_FIRST_ENTITY + ent_idx * 2
        cr_col = dr_col + 1
        dr_total_f = f"=SUM({col_letter(dr_col)}{TB_ROW_FIRST}:{col_letter(dr_col)}{TB_ROW_LAST})"
        cr_total_f = f"=SUM({col_letter(cr_col)}{TB_ROW_FIRST}:{col_letter(cr_col)}{TB_ROW_LAST})"
        ws.write_formula(sc_row, dr_col, dr_total_f, formats["td_bold_right"])
        ws.write_formula(sc_row, cr_col, cr_total_f, formats["td_bold_right"])

    sc.apply_page_setup(ws, sheet_title="Entity TBs")
    return ws


def build_eliminations(wb, formats, dat):
    ws = wb.add_worksheet("Eliminations")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)

    LAST_COL = 6
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 44)
    ws.set_column(2, 2, 16)
    ws.set_column(3, 3, 16)
    ws.set_column(4, 4, 14)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Intercompany journals",
        title="Eliminations",
        last_col=LAST_COL + 2,
        explanation=(
            "Each row records an intercompany journal that nets to zero at the group "
            "level. Pick the two Report lines that get reduced (the asset and the liability, "
            "or the revenue and the cost) and enter the amount once. The Consolidated sheets "
            "subtract these from each side."
        ),
    )

    st.write_section_header(
        ws, formats, row=6,
        kicker="Step 2   Enter your intercompany eliminations",
        title="Elimination journals",
    )

    st.write_header_row(
        ws, formats, row=ELIM_ROW_HEADER - 1,
        headers=["Description", "Report line A (reduces)", "Report line B (reduces)", "Amount"],
        right_align_from=4,
    )

    for i, (desc, line_a, line_b, amount) in enumerate(dat["eliminations"][:ELIM_NUM_ROWS]):
        row_0 = ELIM_ROW_FIRST - 1 + i
        zebra = i % 2 == 0
        ws.set_row(row_0, 22)
        ws.write_string(row_0, 1, desc, formats["input_text"])
        ws.write_string(row_0, 2, line_a, formats["input_text"])
        ws.write_string(row_0, 3, line_b, formats["input_text"])
        ws.write_number(row_0, 4, amount, formats["input_value"])

    ws.data_validation(ELIM_ROW_FIRST - 1, 2, ELIM_ROW_LAST - 1, 3,
                       {"validate": "list", "source": REPORT_LINES})

    # Sum row
    total_row = ELIM_ROW_LAST + 1
    ws.set_row(total_row, 24)
    ws.write_string(total_row, 1, "Total eliminations", formats["total_left"])
    ws.write_blank(total_row, 2, None, formats["total_left"])
    ws.write_blank(total_row, 3, None, formats["total_left"])
    ws.write_formula(total_row, 4, f"=SUM(E{ELIM_ROW_FIRST}:E{ELIM_ROW_LAST})", formats["total_right"])

    sc.apply_page_setup(ws, sheet_title="Eliminations")
    return ws


def _net_for_line(report_line: str) -> str:
    """SUMIFS of Net column on Entity TBs by Report line."""
    return (
        f"=SUMIFS('Entity TBs'!${col_letter(TB_COL_NET)}${TB_ROW_FIRST}"
        f":${col_letter(TB_COL_NET)}${TB_ROW_LAST},"
        f"'Entity TBs'!${col_letter(TB_COL_LINE)}${TB_ROW_FIRST}"
        f":${col_letter(TB_COL_LINE)}${TB_ROW_LAST},\"{report_line}\")"
    )


def _elim_for_line(report_line: str) -> str:
    """Sum of elimination amounts where either side matches this report line."""
    return (
        f"=SUMIFS('Eliminations'!$E${ELIM_ROW_FIRST}:$E${ELIM_ROW_LAST},"
        f"'Eliminations'!$C${ELIM_ROW_FIRST}:$C${ELIM_ROW_LAST},\"{report_line}\")"
        f"+SUMIFS('Eliminations'!$E${ELIM_ROW_FIRST}:$E${ELIM_ROW_LAST},"
        f"'Eliminations'!$D${ELIM_ROW_FIRST}:$D${ELIM_ROW_LAST},\"{report_line}\")"
    )


def build_consolidated_pl(wb, formats, dat):
    ws = wb.add_worksheet("Consolidated P&L")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 6
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 30)
    for c in range(2, LAST_COL):
        ws.set_column(c, c, 16)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Group result",
        title="Consolidated profit and loss",
        last_col=LAST_COL + 2,
        explanation=(
            "Sum of all entity profit and loss accounts from the Entity TBs sheet, less "
            "intercompany eliminations from the Eliminations sheet. Revenue is shown as "
            "the credit balance; expenses are shown positive. Net profit is the group result."
        ),
    )

    st.write_section_header(ws, formats, row=6, kicker="Drawn from Entity TBs and Eliminations",
                            title="Group P&L for the period")

    header_row = 8
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Line item", "Pre-elimination", "Eliminations", "Consolidated"],
                        right_align_from=2)

    # P&L lines: Revenue (CR balance, shown positive after sign flip), Cost of sales (DR positive),
    # Wages, Other opex, D&A. Then Gross profit and Net profit derivations.
    # On the TB sheet the Net = DR - CR. For revenue accounts (CR balance), Net is NEGATIVE.
    # We flip the sign here.
    pl_lines = [
        ("Revenue", "Revenue", -1, False),
        ("Intercompany revenue (eliminated)", "IC Revenue", -1, False),
        ("Cost of sales", "Cost of sales", 1, False),
        ("Wages", "Wages", 1, False),
        ("Other opex", "Other opex", 1, False),
        ("Depreciation and amortisation", "D&A", 1, False),
    ]

    r = header_row + 1
    first_data_row = r + 1
    for i, (label, rl, sign, bold) in enumerate(pl_lines):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        pre_elim = f"={sign}*({_net_for_line(rl)[1:]})"
        elim = f"={_elim_for_line(rl)[1:]}"
        cons = f"=C{r + 1}-D{r + 1}"
        ws.write_string(r, 1, label, label_fmt)
        ws.write_formula(r, 2, pre_elim, num_fmt)
        ws.write_formula(r, 3, elim, num_fmt)
        ws.write_formula(r, 4, cons, num_fmt)
        r += 1

    # Gross profit = Revenue (net of IC) - Cost of sales
    # Row references: rows for Revenue (first_data_row), IC revenue (first_data_row+1), COGS (first_data_row+2)
    gp_row = r
    ws.set_row(r, 24)
    ws.write_string(r, 1, "Gross profit", formats["td_bold_left"])
    ws.write_formula(r, 2, f"=C{first_data_row}+C{first_data_row + 1}-C{first_data_row + 2}", formats["td_bold_right"])
    ws.write_formula(r, 3, f"=D{first_data_row}+D{first_data_row + 1}-D{first_data_row + 2}", formats["td_bold_right"])
    ws.write_formula(r, 4, f"=E{first_data_row}+E{first_data_row + 1}-E{first_data_row + 2}", formats["td_bold_right"])
    r += 1

    # Net profit = GP - Wages - Opex - D&A
    np_row = r
    ws.set_row(r, 24)
    ws.write_string(r, 1, "Net profit", formats["total_left"])
    ws.write_formula(r, 2, f"=C{gp_row}-C{first_data_row + 3}-C{first_data_row + 4}-C{first_data_row + 5}", formats["total_right"])
    ws.write_formula(r, 3, f"=D{gp_row}-D{first_data_row + 3}-D{first_data_row + 4}-D{first_data_row + 5}", formats["total_right"])
    ws.write_formula(r, 4, f"=E{gp_row}-E{first_data_row + 3}-E{first_data_row + 4}-E{first_data_row + 5}", formats["total_right"])
    r += 2

    # Tie-out checks
    st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {
                "name": "Eliminations column for IC Revenue equals the IC Revenue eliminations entered",
                "left": f"=D{first_data_row + 1}",
                "right": _elim_for_line("IC Revenue"),
            },
            {
                "name": "Consolidated Revenue net of IC ties to TB Revenue less IC Revenue elimination",
                "left": f"=E{first_data_row}+E{first_data_row + 1}",
                "right": f"=(-1)*({_net_for_line('Revenue')[1:]})+(-1)*({_net_for_line('IC Revenue')[1:]})-{_elim_for_line('IC Revenue')[1:]}",
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Consolidated P&L")
    return ws


def build_consolidated_bs(wb, formats, dat):
    ws = wb.add_worksheet("Consolidated BS")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    LAST_COL = 6
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 30)
    for c in range(2, LAST_COL):
        ws.set_column(c, c, 16)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(
        ws, formats,
        kicker="Group position",
        title="Consolidated balance sheet",
        last_col=LAST_COL + 2,
        explanation=(
            "Sum of each entity's balance sheet from the Entity TBs sheet, less the "
            "intercompany asset and liability eliminations from the Eliminations sheet. "
            "Total assets must equal Total liabilities plus Equity."
        ),
    )

    st.write_section_header(ws, formats, row=6, kicker="Drawn from Entity TBs and Eliminations",
                            title="Group balance sheet at period end")

    header_row = 8
    st.write_header_row(ws, formats, row=header_row,
                        headers=["Line item", "Pre-elimination", "Eliminations", "Consolidated"],
                        right_align_from=2)

    # BS lines:
    # Assets: Cash, Receivables, IC Asset, Inventory, PPE (DR balance → Net is positive)
    # Liabilities: Payables, IC Liability, Loans (CR balance → Net is negative; flip sign)
    # Equity: Share Capital, Retained Earnings (CR balance → Net is negative; flip sign)
    bs_lines = [
        ("ASSETS",                "section", None, 1),
        ("Cash and equivalents",  "line", "Cash",        1),
        ("Trade receivables",     "line", "Receivables", 1),
        ("Intercompany receivable", "line", "IC Asset",  1),
        ("Inventory",             "line", "Inventory",   1),
        ("Property, plant, equipment", "line", "PPE",    1),
        ("Total assets",          "subtotal", None,      None),
        ("LIABILITIES",           "section", None, None),
        ("Trade payables",        "line", "Payables",    -1),
        ("Intercompany payable",  "line", "IC Liability", -1),
        ("Loans",                 "line", "Loans",       -1),
        ("Total liabilities",     "subtotal", None,      None),
        ("EQUITY",                "section", None, None),
        ("Equity (share capital + retained earnings)", "line", "Equity", -1),
        ("Total equity",          "subtotal", None,      None),
    ]

    r = header_row + 1
    asset_rows = []
    liab_rows = []
    eq_rows = []
    total_assets_row = None
    total_liab_row = None
    total_eq_row = None
    current_section = None

    for entry in bs_lines:
        label = entry[0]
        kind = entry[1]
        rl = entry[2]
        sign = entry[3]

        ws.set_row(r, 22)
        if kind == "section":
            current_section = label
            ws.write_string(r, 1, label, formats["section_kicker"])
        elif kind == "line":
            pre = f"={sign}*({_net_for_line(rl)[1:]})"
            elim = f"={_elim_for_line(rl)[1:]}"
            cons = f"=C{r + 1}-D{r + 1}"
            ws.write_string(r, 1, label, formats["td"])
            ws.write_formula(r, 2, pre, formats["td_right"])
            ws.write_formula(r, 3, elim, formats["td_right"])
            ws.write_formula(r, 4, cons, formats["td_right"])
            if current_section == "ASSETS":
                asset_rows.append(r + 1)
            elif current_section == "LIABILITIES":
                liab_rows.append(r + 1)
            elif current_section == "EQUITY":
                eq_rows.append(r + 1)
        elif kind == "subtotal":
            ws.write_string(r, 1, label, formats["total_left"])
            if current_section == "ASSETS":
                rows = asset_rows
                total_assets_row = r + 1
            elif current_section == "LIABILITIES":
                rows = liab_rows
                total_liab_row = r + 1
            else:
                rows = eq_rows
                total_eq_row = r + 1
            for col_letter_x in ("C", "D", "E"):
                if rows:
                    sum_formula = f"={'+'.join(f'{col_letter_x}{ri}' for ri in rows)}"
                else:
                    sum_formula = "=0"
                col_idx = {"C": 2, "D": 3, "E": 4}[col_letter_x]
                ws.write_formula(r, col_idx, sum_formula, formats["total_right"])
        r += 1

    # Balance check
    r += 1
    st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {
                "name": "Consolidated Total Assets equals Consolidated Total Liabilities plus Equity",
                "left": f"=E{total_assets_row}",
                "right": f"=E{total_liab_row}+E{total_eq_row}",
            },
            {
                "name": "Intercompany receivable eliminated equals intercompany payable eliminated",
                "left": _elim_for_line("IC Asset"),
                "right": _elim_for_line("IC Liability"),
            },
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Consolidated BS")
    return ws


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)

    sc.add_cover_sheet(
        wb, formats,
        workbook_title=WORKBOOK_TITLE,
        workbook_kicker=WORKBOOK_KICKER,
        workbook_id=WORKBOOK_ID,
        version=b.WORKBOOK_VERSION,
        build_date=date.today(),
        how_to_use=HOW_TO_USE,
        target_user=TARGET_USER,
        example_profile=EXAMPLE_PROFILE,
        inputs_required=INPUTS_REQUIRED,
    )
    build_consolidated_pl(wb, formats, dat)
    build_consolidated_bs(wb, formats, dat)
    build_entity_tbs(wb, formats, dat)
    build_eliminations(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
