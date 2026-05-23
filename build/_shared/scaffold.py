"""Scaffold helpers shared across every Lyros library workbook (xlsxwriter).

Two required sheets:
- `Cover` — wordmark band, embedded Lyros logo PNG, kicker, title, three-line
  how-to, target user, optional example profile and inputs, synthetic-data
  disclaimer, full-width Book here CTA banner.
- `Connect your data` — last sheet, two paths to populate the workbook (manual
  vs Lyros-assisted) plus a "Book here" hyperlink to the bookings URL.

Page setup (A4 landscape, footer with bookings link) is applied to every
analytical sheet via `apply_page_setup`.
"""

from __future__ import annotations

from datetime import date
from typing import Any

import xlsxwriter

from . import branding as b


# ── Page setup ──────────────────────────────────────────────────────────────

def apply_page_setup(ws, *, sheet_title: str) -> None:
    """A4 landscape print, fixed footer with bookings, fit width."""
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.fit_to_pages(1, 0)
    ws.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.6)
    ws.center_horizontally()
    ws.repeat_rows(0, 4)  # repeat hero band on subsequent pages

    ws.set_header(
        f"&L&\"Arial\"&8&K707070Lyros Accounting"
        f"&C&\"Arial\"&8&K707070{sheet_title}"
        f"&R&\"Arial\"&8&K707070Page &P of &N",
        margin=0.2,
    )
    ws.set_footer(
        f"&L&\"Arial\"&8&K707070lyros.com.au"
        f"&C&\"Arial\"&8&K2D7A55Book a 15-min call: bookings.cloud.microsoft/book/LyrosAccounting"
        f"&R&\"Arial\"&8&K707070&D",
        margin=0.3,
    )


# ── Hero band (top 4 rows on each analytical sheet) ─────────────────────────

def write_hero_band(
    ws,
    formats: dict,
    *,
    kicker: str,
    title: str,
    last_col: int = 14,
    explanation: str = "",
) -> None:
    """Dark hero band at the top of an analytical sheet.

    Row 1: top buffer (dark)
    Row 2: kicker text in green on dark
    Row 3: title in white on dark
    Row 4: thin green accent line
    Row 5 (optional): plain-text explanation of what the sheet does

    Logo: small white logomark in the top-left corner of the band.
    """
    for col in range(last_col):
        ws.write_blank(0, col, None, formats["wordmark_band"])
        ws.write_blank(1, col, None, formats["wordmark_band"])
        ws.write_blank(2, col, None, formats["wordmark_band"])
        ws.write_blank(3, col, None, formats["accent_band"])

    # Logomark on the RIGHT side of the dark band. Kicker and title sit
    # on the left (columns B onwards). Logo is white-on-transparent, so it
    # stays entirely within the dark band rather than extending into the
    # white area below.
    try:
        ws.insert_image(
            0, last_col - 2, str(b.LOGOMARK_WHITE),
            {
                "x_scale": 0.20, "y_scale": 0.20,
                "x_offset": 4, "y_offset": 4,
                "object_position": 3,
            },
        )
    except Exception:
        pass

    # Kicker and title sit on the left, well clear of the right-side logo
    ws.merge_range(1, 1, 1, last_col - 3, kicker.upper(), formats["hero_kicker"])
    ws.merge_range(2, 1, 2, last_col - 3, title, formats["hero_title"])

    ws.set_row(0, 14)
    ws.set_row(1, 16)
    ws.set_row(2, 26)
    ws.set_row(3, 4)

    if explanation:
        # Size the row based on text length so wrapped text shows in full
        char_count = len(explanation)
        # Estimate ~3.5 chars per column-pixel-of-width at 9pt italic. Merged
        # range spans (last_col - 3) columns at ~11px wide → ~120 chars/line.
        chars_per_line = max(80, (last_col - 3) * 11)
        lines = max(2, -(-char_count // chars_per_line))  # ceil division
        height = max(40, lines * 16)
        ws.set_row(4, height)
        ws.merge_range(
            4, 1, 4, last_col - 2, explanation, formats["body_muted"],
        )


# ── Cover sheet ─────────────────────────────────────────────────────────────

def add_cover_sheet(
    wb: xlsxwriter.Workbook,
    formats: dict,
    *,
    workbook_title: str,
    workbook_kicker: str,
    workbook_id: str | None = None,
    version: str | None = None,
    build_date: date | None = None,
    how_to_use: list[str] | None = None,
    target_user: str = "",
    example_profile: list[tuple[str, str]] | None = None,
    inputs_required: list[tuple[str, str]] | None = None,
):
    """Create the Cover sheet as the first sheet of the workbook.

    example_profile: optional list of (label, value) pairs describing the
        synthetic example business (e.g. industry, revenue scale, margin).
    inputs_required: optional list of (input_name, used_on_tab) pairs naming
        the figures the user must replace and where each one is consumed.

    workbook_id, version, build_date are accepted for backwards compatibility
    with existing build scripts but are no longer rendered on the cover.
    """
    del workbook_id, version, build_date
    if how_to_use is None:
        how_to_use = [
            "Open the Data sheet and replace the sample figures with your own.",
            "Every analytical sheet recalculates automatically from the Data sheet.",
            "Read the Headline sheet for the summary and commentary.",
        ]

    ws = wb.add_worksheet("Cover")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.GREEN)

    # Column widths
    for col in range(13):
        ws.set_column(col, col, 11.5)
    ws.set_column(0, 0, 2)
    ws.set_column(12, 12, 2)

    LAST_COL = 12

    # Dark band: rows 1-4 dark fill
    for r in range(0, 4):
        for c in range(0, LAST_COL + 1):
            ws.write_blank(r, c, None, formats["wordmark_band"])
    ws.set_row(0, 18)
    ws.set_row(1, 26)
    ws.set_row(2, 26)
    ws.set_row(3, 18)

    # Wordmark anchored to the bottom-RIGHT of the dark band.
    try:
        ws.insert_image(
            0, LAST_COL - 5, str(b.LOGO_FULL_WHITE),
            {
                "x_scale": 0.11, "y_scale": 0.11,
                "x_offset": 4, "y_offset": 30,
                "object_position": 3,
            },
        )
    except Exception:
        ws.merge_range(1, 1, 2, LAST_COL - 1, "LYROS ACCOUNTING", formats["cover_brand_band_text"])

    # Accent line under the dark band
    for c in range(0, LAST_COL + 1):
        ws.write_blank(4, c, None, formats["accent_band"])
    ws.set_row(4, 4)
    ws.set_row(5, 8)

    # Row counter for everything below the band
    cur = 7

    # Kicker + title block
    ws.set_row(cur, 18)
    ws.write(cur, 1, workbook_kicker.upper(), formats["cover_kicker"])
    cur += 1
    ws.set_row(cur, 36)
    ws.merge_range(cur, 1, cur, LAST_COL - 1, workbook_title, formats["cover_title"])
    cur += 2

    # How to use
    ws.set_row(cur, 18)
    ws.write(cur, 1, "HOW TO USE", formats["cover_kicker"])
    cur += 1
    for i, line in enumerate(how_to_use[:3]):
        ws.set_row(cur, 24)
        ws.write(cur, 1, f"{i + 1}.", formats["body_muted"])
        ws.merge_range(cur, 2, cur, LAST_COL - 1, line, formats["cover_body"])
        cur += 1
    cur += 1

    # Target user
    if target_user:
        ws.set_row(cur, 18)
        ws.write(cur, 1, "DESIGNED FOR", formats["cover_kicker"])
        cur += 1
        ws.set_row(cur, 24)
        ws.merge_range(cur, 1, cur, LAST_COL - 1, target_user, formats["cover_body"])
        cur += 2

    # Example business profile (what kind of company the synthetic data
    # represents — helps the user gauge whether the example fits them).
    if example_profile:
        ws.set_row(cur, 18)
        ws.write(cur, 1, "EXAMPLE BUSINESS PROFILE", formats["cover_kicker"])
        cur += 1
        ws.set_row(cur, 24)
        ws.merge_range(
            cur, 1, cur, LAST_COL - 1,
            "Synthetic data inside this workbook represents the following business shape. "
            "Use it as a reference for what good looks like; your numbers will differ.",
            formats["body_muted"],
        )
        cur += 1
        for label, value in example_profile:
            ws.set_row(cur, 22)
            ws.write_string(cur, 1, label, formats["cover_kicker"])
            ws.merge_range(cur, 3, cur, LAST_COL - 1, "", formats["cover_meta"])
            ws.write_string(cur, 3, value, formats["cover_meta"])
            cur += 1
        cur += 1

    # Inputs the user needs to provide (with the tab each input drives).
    if inputs_required:
        ws.set_row(cur, 18)
        ws.write(cur, 1, "INPUTS YOU NEED TO PROVIDE", formats["cover_kicker"])
        cur += 1
        ws.set_row(cur, 24)
        ws.merge_range(
            cur, 1, cur, LAST_COL - 1,
            "These figures vary by company and cannot be exported directly from "
            "your accounting software. Replace the amber-bordered sample values "
            "on the tabs noted below.",
            formats["body_muted"],
        )
        cur += 1
        for input_name, used_on in inputs_required:
            ws.set_row(cur, 22)
            ws.write_string(cur, 1, input_name, formats["cover_kicker"])
            ws.merge_range(cur, 5, cur, LAST_COL - 1, "", formats["cover_meta"])
            ws.write_string(cur, 5, f"Used on: {used_on}", formats["cover_meta"])
            cur += 1
        cur += 1

    # Disclosure block: salesy + professional. Three lines stacked, each in a
    # different visual weight to read as a closing pitch rather than fine
    # print. Lyros voice (no contractions, no em-dashes, no exclamations).
    dr = cur
    ws.set_row(dr, 18)
    ws.merge_range(
        dr, 1, dr, LAST_COL - 1,
        "WHAT THIS IS",
        formats["cover_kicker"],
    )

    ws.set_row(dr + 1, 24)
    ws.merge_range(
        dr + 1, 1, dr + 1, LAST_COL - 1,
        "Free. Professionally designed. Pre-populated with synthetic data so you can "
        "see exactly what good monthly reporting looks like before investing your own time.",
        formats["cover_body"],
    )

    ws.set_row(dr + 2, 18)
    ws.merge_range(
        dr + 2, 1, dr + 2, LAST_COL - 1,
        "MAKE IT YOURS",
        formats["cover_kicker"],
    )

    ws.set_row(dr + 3, 38)
    ws.merge_range(
        dr + 3, 1, dr + 3, LAST_COL - 1,
        "Replace the figures on the Data sheet with your own to use this workbook as a "
        "template. Or invite Lyros as adviser on your accounting software and Lyros will "
        "populate the workbook accurately and walk you through it on a 15-minute call.",
        formats["cover_body"],
    )

    ws.set_row(dr + 4, 18)
    ws.merge_range(
        dr + 4, 1, dr + 4, LAST_COL - 1,
        "DISCLOSURE",
        formats["cover_kicker"],
    )

    ws.set_row(dr + 5, 34)
    ws.merge_range(
        dr + 5, 1, dr + 5, LAST_COL - 1,
        "This workbook is provided free of charge as a visual template. It has not been "
        "reviewed against any individual circumstances and is not financial advice. "
        "All figures and company names inside are synthetic. Review and tailor with your "
        "finance lead before relying on any output.",
        formats["cover_disclaimer"],
    )

    # Full-width Book here CTA banner (matches the Connect tab CTA).
    cta_row = dr + 7
    ws.set_row(cta_row, 28)
    ws.set_row(cta_row + 1, 28)
    ws.merge_range(cta_row, 1, cta_row + 1, LAST_COL - 1, "", formats["link_cta"])
    ws.write_url(
        cta_row, 1, b.BOOKINGS_URL, formats["link_cta"],
        "Book here   |   15-minute discovery call",
    )

    # Right accent rail
    for r in range(7, cta_row + 2):
        ws.write_blank(r, LAST_COL, None, formats["accent_soft_band"])

    ws.print_area(0, 0, cta_row + 3, LAST_COL)
    return ws


# ── Connect your data sheet ─────────────────────────────────────────────────

def add_connect_data_sheet(
    wb: xlsxwriter.Workbook,
    formats: dict,
    *,
    workbook_title: str,
):
    ws = wb.add_worksheet("Connect your data")
    ws.hide_gridlines(2)
    ws.set_tab_color(b.ACCENT_SOFT)

    LAST_COL = 12
    for col in range(LAST_COL + 1):
        ws.set_column(col, col, 11.5)
    ws.set_column(0, 0, 2)
    ws.set_column(LAST_COL, LAST_COL, 2)

    write_hero_band(
        ws, formats,
        kicker="Populate this workbook",
        title="Connect your accounting data",
        last_col=LAST_COL + 1,
    )

    row = 6
    options = [
        (
            "Option 1   Enter the data yourself",
            [
                f"Export the relevant report from your accounting software (e.g. {workbook_title.lower()} or trial balance).",
                "Paste the figures into the input cells on the Data sheet.",
                "All formulas, charts, and summaries update automatically.",
            ],
            "Best for owner-operated businesses willing to spend 30 minutes per month.",
        ),
        (
            "Option 2   Invite Lyros to populate it for you",
            [
                "Invite Lyros Accounting as an adviser on your accounting software (we will send instructions).",
                "We connect to your file, populate this workbook with your figures, and walk you through it on a 15-minute call.",
                "We can then maintain the workbook on the cadence you choose (monthly, quarterly, or ad-hoc).",
            ],
            "Best for finance leads who want the workbook used as a working document, not a one-off.",
        ),
    ]
    for label, steps, footer in options:
        ws.set_row(row, 28)
        ws.merge_range(row, 1, row, LAST_COL - 1, label, formats["section_h2"])
        row += 1
        for i, step in enumerate(steps):
            ws.set_row(row, 24)
            ws.write(row, 1, f"{i + 1}.", formats["body_muted"])
            ws.merge_range(row, 2, row, LAST_COL - 1, step, formats["cover_body"])
            row += 1
        ws.set_row(row, 22)
        ws.merge_range(row, 1, row, LAST_COL - 1, footer, formats["body_muted"])
        row += 2

    # CTA banner (merged cell containing a hyperlink labelled "Book here")
    cta_row = row + 1
    ws.set_row(cta_row, 24)
    ws.set_row(cta_row + 1, 24)
    ws.merge_range(cta_row, 1, cta_row + 1, LAST_COL - 1, "", formats["link_cta"])
    ws.write_url(
        cta_row, 1, b.BOOKINGS_URL, formats["link_cta"],
        "Book here   |   15-minute discovery call",
    )

    apply_page_setup(ws, sheet_title="Connect your data")
    ws.print_area(0, 0, cta_row + 2, LAST_COL)
    return ws


# Backwards-compatibility alias for build scripts still calling the old name.
add_connect_xero_sheet = add_connect_data_sheet
