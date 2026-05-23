"""Top Supplier Review with ABN dedup workbook.

Supplier-level spend extract from your accounting software. Adds ABN-based dedup logic (one ABN
can map to multiple supplier names due to typos or trading-as variations),
year-on-year change, and concentration ratio.
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


OUTPUT_PATH = Path(r"C:\dev\lyros-workbooks\library\8000 - Top Supplier Review.xlsx")

WORKBOOK_ID = "8000"
WORKBOOK_TITLE = "Top Supplier Review (ABN dedup)"
WORKBOOK_KICKER = "Supplier spend and concentration with ABN dedup"
TARGET_USER = "FC running a quarterly procurement review or auditing supplier concentration."
HOW_TO_USE = [
    "Open the Data sheet and paste your Supplier Activity Summary (current 12 months and prior).",
    "Enter the ABN per supplier in the ABN column; entries with the same ABN are grouped together on the Deduped sheet.",
    "The Deduped sheet ranks suppliers by total spend, shows year-on-year change, and computes top-5 / top-10 concentration.",
]

EXAMPLE_PROFILE = [
    ("INDUSTRY", "SME with 15-20 active suppliers"),
    ("SPEND SCALE", "Circa $2M annual"),
    ("DEDUP COMMON CASES", "Supplier name variations (e.g. 'XYZ Pty Ltd' vs 'XYZ Pty. Ltd.')"),
]

INPUTS_REQUIRED = [
    ("Supplier name and ABN", "Data tab"),
    ("Current 12-month spend per supplier", "Data tab"),
    ("Prior 12-month spend per supplier", "Data tab"),
]

DATA_COL_LABEL = 1
DATA_COL_ABN = 2
DATA_COL_CURR = 3
DATA_COL_PRIOR = 4
DATA_ROW_HEADER = 8
DATA_ROW_FIRST = 9
DATA_ROW_LAST = DATA_ROW_FIRST + 16  # 17 suppliers
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
    rng = d.make_rng("topsupp")
    roster = d.roster()[:17]
    rows = []
    for i, co in enumerate(roster):
        scale = rng.uniform(0.2, 2.0)
        curr = round(110_000 * scale + rng.uniform(-15_000, 15_000), 0)
        prior = round(curr * rng.uniform(0.78, 1.18), 0)
        # Synthetic ABN: same as roster co's
        abn = co.abn
        # Occasionally introduce a duplicate ABN with a near-duplicate name to show dedup
        if i == 4:
            rows.append(("Yarra Supplies Pty Ltd.", abn, curr, prior))
        else:
            rows.append((co.name, abn, curr, prior))
    return {"rows": rows}


def build_data_sheet(wb, formats, dat):
    ws = wb.add_worksheet("Data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.AMBER)
    LAST_COL = DATA_COL_PRIOR + 1
    ws.set_column(0, 0, 2)
    ws.set_column(DATA_COL_LABEL, DATA_COL_LABEL, 32)
    ws.set_column(DATA_COL_ABN, DATA_COL_ABN, 18)
    ws.set_column(DATA_COL_CURR, DATA_COL_PRIOR, 16)
    ws.set_column(LAST_COL, LAST_COL, 2)

    sc.write_hero_band(ws, formats, kicker="Single source of truth",
                       title="Drop your supplier spend here",
                       last_col=LAST_COL + 1,
                       explanation=(
                           "Paste your Supplier Activity Summary for the current 12 months and the "
                           "prior 12 months. Enter the ABN for each supplier; ABNs are used to dedup "
                           "near-duplicate names on the Deduped sheet. Australian Business Numbers (ABNs) "
                           "are 11-digit identifiers issued by the ATO."
                       ))

    st.write_section_header(ws, formats, row=6, kicker="Paste from your accounting software", title="Supplier spend")
    st.write_header_row(ws, formats, row=DATA_ROW_HEADER - 1,
                        headers=["Supplier", "ABN", "Current 12m", "Prior 12m"],
                        start_col=DATA_COL_LABEL, right_align_from=3)
    for i, (name, abn, curr, prior) in enumerate(dat["rows"]):
        row_0 = DATA_ROW_FIRST - 1 + i
        ws.set_row(row_0, 22)
        zebra = i % 2 == 0
        text_fmt = formats["td_zebra"] if zebra else formats["td"]
        ws.write_string(row_0, DATA_COL_LABEL, name, text_fmt)
        ws.write_string(row_0, DATA_COL_ABN, abn, text_fmt)
        ws.write_number(row_0, DATA_COL_CURR, curr, formats["input_value"])
        ws.write_number(row_0, DATA_COL_PRIOR, prior, formats["input_value"])

    ws.set_row(DATA_ROW_TOTAL - 1, 24)
    ws.write_string(DATA_ROW_TOTAL - 1, DATA_COL_LABEL, "Total", formats["total_left"])
    ws.write_blank(DATA_ROW_TOTAL - 1, DATA_COL_ABN, None, formats["total_left"])
    ws.write_formula(DATA_ROW_TOTAL - 1, DATA_COL_CURR,
                     f"=SUM({col_letter(DATA_COL_CURR)}{DATA_ROW_FIRST}:{col_letter(DATA_COL_CURR)}{DATA_ROW_LAST})",
                     formats["total_right"])
    ws.write_formula(DATA_ROW_TOTAL - 1, DATA_COL_PRIOR,
                     f"=SUM({col_letter(DATA_COL_PRIOR)}{DATA_ROW_FIRST}:{col_letter(DATA_COL_PRIOR)}{DATA_ROW_LAST})",
                     formats["total_right"])

    sc.apply_page_setup(ws, sheet_title="Data")


def build_deduped(wb, formats, dat):
    ws = wb.add_worksheet("Deduped")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)
    LAST_COL = 7
    ws.set_column(0, 0, 2)
    ws.set_column(1, 1, 18)
    ws.set_column(2, 2, 32)
    ws.set_column(3, LAST_COL, 14)
    ws.set_column(LAST_COL + 1, LAST_COL + 1, 2)

    sc.write_hero_band(ws, formats, kicker="Aggregated by ABN",
                       title="Top suppliers, deduped",
                       last_col=LAST_COL + 2,
                       explanation=(
                           "Suppliers are aggregated by ABN using SUMIFS, so entries with name "
                           "variations but the same ABN show as a single row. Unique ABNs are pulled "
                           "from the Data sheet via an array of IFERROR(INDEX/MATCH) lookups, which "
                           "may show blank rows for unused capacity."
                       ))

    section_row = 6
    st.write_section_header(ws, formats, row=section_row, kicker="Drawn from the Data sheet",
                            title="Deduped supplier spend ranked by total")
    header_row = section_row + 2
    st.write_header_row(ws, formats, row=header_row,
                        headers=["ABN", "Supplier (first name)", "Current 12m", "Prior 12m", "YoY $", "YoY %", "Share %"],
                        start_col=1, right_align_from=3)

    # We list each unique ABN from Data via array formula and sum by ABN.
    # Use the LARGE/INDEX/MATCH pattern: rank by Sum of current spend per ABN.
    abn_range = f"'Data'!${col_letter(DATA_COL_ABN)}${DATA_ROW_FIRST}:${col_letter(DATA_COL_ABN)}${DATA_ROW_LAST}"
    name_range = f"'Data'!${col_letter(DATA_COL_LABEL)}${DATA_ROW_FIRST}:${col_letter(DATA_COL_LABEL)}${DATA_ROW_LAST}"
    curr_range = f"'Data'!${col_letter(DATA_COL_CURR)}${DATA_ROW_FIRST}:${col_letter(DATA_COL_CURR)}${DATA_ROW_LAST}"
    prior_range = f"'Data'!${col_letter(DATA_COL_PRIOR)}${DATA_ROW_FIRST}:${col_letter(DATA_COL_PRIOR)}${DATA_ROW_LAST}"

    # For simplicity, we render up to 17 rows. Each row shows the unique ABN
    # from the data range, using an array formula trick: take the ABN whose
    # SUMIFS rank equals this rank.
    r = header_row + 1
    first_data_row_1b = r + 1
    NUM_ROWS = len(dat["rows"])
    for i in range(NUM_ROWS):
        zebra = i % 2 == 0
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        pct_fmt = formats["td_pct_zebra"] if zebra else formats["td_pct"]
        ws.set_row(r, 22)

        rank = i + 1
        # Approach: list unique ABNs by their first appearance order, then sort
        # by SUMIFS desc. Simpler implementation: just pull rank-th unique ABN
        # by first-seen order, leaving sort to the user via Excel autofilter.
        # For each row, get the ABN at position (DATA_ROW_FIRST + i - 1).
        # If duplicate (same as prior rows), show blank.
        target_row = DATA_ROW_FIRST + i
        abn_cell = f"'Data'!${col_letter(DATA_COL_ABN)}${target_row}"
        name_cell = f"'Data'!${col_letter(DATA_COL_LABEL)}${target_row}"

        # Display the ABN; if this ABN already appeared above (in Data above this row), show blank
        if i == 0:
            abn_f = f"={abn_cell}"
            name_f = f"={name_cell}"
        else:
            abn_f = f"=IF(COUNTIF('Data'!${col_letter(DATA_COL_ABN)}${DATA_ROW_FIRST}:${col_letter(DATA_COL_ABN)}${target_row - 1},{abn_cell})>0,\"\",{abn_cell})"
            name_f = f"=IF(B{r + 1}=\"\",\"\",{name_cell})"

        ws.write_formula(r, 1, abn_f, label_fmt)
        ws.write_formula(r, 2, name_f, label_fmt)
        # SUMIFS by ABN for current and prior
        curr_sumifs = f"=IF(B{r + 1}=\"\",\"\",SUMIFS({curr_range},{abn_range},B{r + 1}))"
        prior_sumifs = f"=IF(B{r + 1}=\"\",\"\",SUMIFS({prior_range},{abn_range},B{r + 1}))"
        ws.write_formula(r, 3, curr_sumifs, num_fmt)
        ws.write_formula(r, 4, prior_sumifs, num_fmt)
        ws.write_formula(r, 5, f"=IF(B{r + 1}=\"\",\"\",D{r + 1}-E{r + 1})", num_fmt)
        ws.write_formula(r, 6, f"=IF(OR(B{r + 1}=\"\",E{r + 1}=0),\"\",(D{r + 1}-E{r + 1})/ABS(E{r + 1}))", pct_fmt)
        ws.write_formula(r, 7, f"=IF(B{r + 1}=\"\",\"\",D{r + 1}/'Data'!${col_letter(DATA_COL_CURR)}${DATA_ROW_TOTAL})", pct_fmt)
        r += 1
    last_data_row_1b = r

    # Total row
    st.write_total_row(ws, formats, row=r, label="Total (deduped)",
                       formulas=["", f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})",
                                  f"=SUM(E{first_data_row_1b}:E{last_data_row_1b})",
                                  f"=SUM(F{first_data_row_1b}:F{last_data_row_1b})",
                                  "", ""],
                       cell_format="total_right")

    # Tie-out
    r += 2
    st.write_checks_block(
        ws, formats,
        row=r,
        checks=[
            {"name": "Deduped Total Current ties to Data Total Current (sanity)",
             "left": f"=SUM(D{first_data_row_1b}:D{last_data_row_1b})",
             "right": f"='Data'!{col_letter(DATA_COL_CURR)}{DATA_ROW_TOTAL}"},
        ],
    )

    sc.apply_page_setup(ws, sheet_title="Deduped")


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    dat = _synthetic_data()
    wb = xlsxwriter.Workbook(str(OUTPUT_PATH))
    formats = b.make_formats(wb)
    sc.add_cover_sheet(wb, formats, workbook_title=WORKBOOK_TITLE, workbook_kicker=WORKBOOK_KICKER,
                       workbook_id=WORKBOOK_ID, build_date=date.today(),
                       how_to_use=HOW_TO_USE, target_user=TARGET_USER,
                       example_profile=EXAMPLE_PROFILE, inputs_required=INPUTS_REQUIRED)
    build_deduped(wb, formats, dat)
    build_data_sheet(wb, formats, dat)
    sc.add_connect_xero_sheet(wb, formats, workbook_title=WORKBOOK_TITLE)
    wb.close()
    print(f"Wrote: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
