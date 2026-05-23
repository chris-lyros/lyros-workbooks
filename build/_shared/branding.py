"""Lyros design tokens for xlsxwriter workbooks.

Single source of truth for color, font, border, and number formats applied
across the workbook library. Aligns with `lyros-design-system.md` and
`tokens.css` in the public site repo.

xlsxwriter uses Format objects bound to a Workbook. Call `make_formats(wb)`
once per workbook and pass the returned dict to sheet builders.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import xlsxwriter


# ── Core palette (matches tokens.css) ────────────────────────────────────────

BLACK = "#000000"
GREY_BLACK = "#1A1A1A"
WHITE = "#FFFFFF"
OFF_WHITE = "#F4F4F4"
PARAGRAPH = "#D9D9E0"
DIMMED = "#9999A5"
SUBTITLE = "#999999"
GREEN = "#3A9E6E"
ACCENT_SOFT = "#2D7A55"
BORDER_GREY = "#E5E5E5"
GRID_GREY = "#D0D0D0"
TEXT_DIM = "#707070"
TEXT_MUTED = "#A0A0A0"

# Soft conditional-formatting tints (much lighter than the status palette, so
# heatmaps stay legible at small sizes and feel like spreadsheet polish rather
# than dashboard alerts).
TINT_GOOD = "#E0F2E5"      # pale green
TINT_GOOD_MID = "#C2E0CC"  # pale green stronger
TINT_NEUTRAL = "#FFFFFF"   # white middle
TINT_BAD = "#FCE5E6"       # pale red
TINT_BAD_MID = "#F5C6C8"   # pale red stronger
TINT_WARN = "#FFF3D9"      # pale amber

# Status tones (used in legend dots only, not as heatmap fills)
RED = "#E5484D"
AMBER = "#F5A524"
BLUE = "#4DABF7"

# Sequential green spectrum for chart series (tints + shades of brand green only)
GREEN_SPECTRUM = [
    "#D8EBE0",  # 1 lightest
    "#B1D7C1",  # 2
    "#7CC09C",  # 3
    "#3A9E6E",  # 4 brand
    "#2D7A55",  # 5 accent soft
    "#1F5B3F",  # 6 darkest
]

CHART_PRIMARY = GREEN
CHART_SECONDARY = ACCENT_SOFT
CHART_TERTIARY = GREY_BLACK


# ── Constants ───────────────────────────────────────────────────────────────

FONT_FAMILY = "Arial"
WORKBOOK_VERSION = "1.0"
LYROS_ENTITY = "Lyros Pty Ltd  ABN 46 689 015 165  trading as Lyros Accounting"
BOOKINGS_URL = "https://bookings.cloud.microsoft/book/LyrosAccounting@lyros.com.au"
SITE_URL = "https://lyros.com.au"

# Asset paths (relative to public site repo). Build script resolves to absolute.
ASSETS_DIR = Path(r"C:\dev\lyros-workbooks\assets")
LOGO_FULL_WHITE = ASSETS_DIR / "logo-full-white.png"
LOGO_FULL_BLACK = ASSETS_DIR / "logo-full-black.png"
LOGOMARK_WHITE = ASSETS_DIR / "logomark-white.png"


# ── Number formats ──────────────────────────────────────────────────────────

NF_AUD = '_-"$"* #,##0_-;[Red]_-"$"* (#,##0)_-;_-"$"* "-"_-;_-@_-'
NF_AUD_K = '_-"$"* #,##0,"k"_-;[Red]_-"$"* (#,##0,"k")_-;_-"$"* "-"_-;_-@_-'
NF_INT = '#,##0;[Red](#,##0);"-"'
NF_PCT = "0.0%;[Red](0.0%);\"-\""
NF_PCT_INT = "0%;[Red](0%);\"-\""
NF_DATE = "yyyy-mm-dd"
NF_MONTH = "mmm yy"
NF_DAYS = '#,##0" days";[Red](#,##0)" days";"-"'


# ── Format factory ──────────────────────────────────────────────────────────

def _base(**overrides: Any) -> dict:
    base = {"font_name": FONT_FAMILY, "font_size": 10, "font_color": GREY_BLACK}
    base.update(overrides)
    return base


def make_formats(wb: xlsxwriter.Workbook) -> dict[str, xlsxwriter.workbook.Format]:
    """Create the workbook's Format registry. Reused by every sheet builder."""
    f: dict[str, Any] = {}

    # Cover and brand band
    f["wordmark_band"] = wb.add_format(_base(bg_color=GREY_BLACK))
    f["accent_band"] = wb.add_format(_base(bg_color=GREEN))
    f["accent_soft_band"] = wb.add_format(_base(bg_color=ACCENT_SOFT))
    f["cover_title"] = wb.add_format(_base(font_size=26, bold=True, font_color=GREY_BLACK, align="left", valign="vcenter"))
    f["cover_kicker"] = wb.add_format(_base(font_size=9, bold=True, font_color=GREEN, align="left", valign="vcenter"))
    f["cover_kicker_white"] = wb.add_format(_base(font_size=9, bold=True, font_color=GREEN, bg_color=GREY_BLACK, align="left", valign="vcenter"))
    f["cover_body"] = wb.add_format(_base(font_size=11, align="left", valign="vcenter", text_wrap=True))
    f["cover_meta"] = wb.add_format(_base(font_size=9, font_color=TEXT_DIM, align="left", valign="vcenter"))
    f["cover_disclaimer"] = wb.add_format(_base(font_size=9, italic=True, font_color=TEXT_DIM, align="left", valign="top", text_wrap=True))
    f["cover_brand_band_text"] = wb.add_format(_base(font_size=22, bold=True, font_color=WHITE, bg_color=GREY_BLACK, align="left", valign="vcenter", indent=2))

    # Section headers (inside analytical sheets)
    f["section_h2"] = wb.add_format(_base(font_size=14, bold=True, font_color=GREY_BLACK, align="left", valign="vcenter"))
    f["section_h3"] = wb.add_format(_base(font_size=11, bold=True, font_color=GREY_BLACK, align="left", valign="vcenter"))
    f["section_kicker"] = wb.add_format(_base(font_size=9, bold=True, font_color=GREEN, align="left", valign="vcenter"))

    # Sheet hero band (sits at top of every analytical sheet)
    f["hero_band"] = wb.add_format(_base(bg_color=GREY_BLACK, font_color=WHITE))
    f["hero_title"] = wb.add_format(_base(font_size=18, bold=True, font_color=WHITE, bg_color=GREY_BLACK, align="left", valign="vcenter", indent=1))
    f["hero_kicker"] = wb.add_format(_base(font_size=8, bold=True, font_color=GREEN, bg_color=GREY_BLACK, align="left", valign="vcenter", indent=1))

    # Table header row
    f["th"] = wb.add_format(_base(
        bold=True, font_color=WHITE, bg_color=GREY_BLACK,
        align="left", valign="vcenter",
        top=2, top_color=GREEN, bottom=2, bottom_color=GREEN,
        text_wrap=True,
    ))
    f["th_right"] = wb.add_format(_base(
        bold=True, font_color=WHITE, bg_color=GREY_BLACK,
        align="right", valign="vcenter",
        top=2, top_color=GREEN, bottom=2, bottom_color=GREEN,
        text_wrap=True,
    ))

    # Data cells (zebra and white variants)
    f["td"] = wb.add_format(_base(align="left", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=WHITE))
    f["td_zebra"] = wb.add_format(_base(align="left", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=OFF_WHITE))
    f["td_right"] = wb.add_format(_base(align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=WHITE, num_format=NF_AUD))
    f["td_right_zebra"] = wb.add_format(_base(align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=OFF_WHITE, num_format=NF_AUD))
    f["td_pct"] = wb.add_format(_base(align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=WHITE, num_format=NF_PCT))
    f["td_pct_zebra"] = wb.add_format(_base(align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=OFF_WHITE, num_format=NF_PCT))
    f["td_days"] = wb.add_format(_base(align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=WHITE, num_format=NF_DAYS))
    f["td_days_zebra"] = wb.add_format(_base(align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=OFF_WHITE, num_format=NF_DAYS))

    # Bold rows (subtotals): same as td but bold and no fill change
    f["td_bold_left"] = wb.add_format(_base(bold=True, align="left", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=WHITE))
    f["td_bold_right"] = wb.add_format(_base(bold=True, align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=WHITE, num_format=NF_AUD))
    f["td_bold_pct"] = wb.add_format(_base(bold=True, align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=WHITE, num_format=NF_PCT))

    # Strong total row (dark fill, white text)
    f["total_left"] = wb.add_format(_base(font_size=11, bold=True, font_color=WHITE, bg_color=GREY_BLACK, align="left", valign="vcenter", indent=1, top=2, top_color=GREEN))
    f["total_right"] = wb.add_format(_base(font_size=11, bold=True, font_color=WHITE, bg_color=GREY_BLACK, align="right", valign="vcenter", num_format=NF_AUD, top=2, top_color=GREEN))
    f["total_pct"] = wb.add_format(_base(font_size=11, bold=True, font_color=WHITE, bg_color=GREY_BLACK, align="right", valign="vcenter", num_format=NF_PCT, top=2, top_color=GREEN))
    f["total_days"] = wb.add_format(_base(font_size=11, bold=True, font_color=WHITE, bg_color=GREY_BLACK, align="right", valign="vcenter", num_format=NF_DAYS, top=2, top_color=GREEN))

    # KPI card cells
    f["kpi_label"] = wb.add_format(_base(font_size=8, bold=True, font_color=GREEN, bg_color=OFF_WHITE, align="left", valign="vcenter", indent=1))
    f["kpi_value"] = wb.add_format(_base(font_size=20, bold=True, font_color=GREY_BLACK, bg_color=OFF_WHITE, align="left", valign="vcenter", indent=1, num_format=NF_AUD))
    f["kpi_value_pct"] = wb.add_format(_base(font_size=20, bold=True, font_color=GREY_BLACK, bg_color=OFF_WHITE, align="left", valign="vcenter", indent=1, num_format=NF_PCT))
    f["kpi_change"] = wb.add_format(_base(font_size=9, font_color=TEXT_DIM, bg_color=OFF_WHITE, align="left", valign="vcenter", indent=1))

    # Input cells (Data sheet): users replace these values
    f["input_label"] = wb.add_format(_base(font_size=10, bold=True, align="left", valign="vcenter", indent=1, bg_color=OFF_WHITE, border=1, border_color=BORDER_GREY))
    f["input_value"] = wb.add_format(_base(font_size=10, font_color=ACCENT_SOFT, align="right", valign="vcenter", bg_color="#FFFEF7", border=1, border_color=AMBER, num_format=NF_AUD))
    f["input_pct"] = wb.add_format(_base(font_size=10, font_color=ACCENT_SOFT, align="right", valign="vcenter", bg_color="#FFFEF7", border=1, border_color=AMBER, num_format=NF_PCT))
    f["input_days"] = wb.add_format(_base(font_size=10, font_color=ACCENT_SOFT, align="right", valign="vcenter", bg_color="#FFFEF7", border=1, border_color=AMBER, num_format=NF_DAYS))
    f["input_text"] = wb.add_format(_base(font_size=10, font_color=ACCENT_SOFT, align="left", valign="vcenter", bg_color="#FFFEF7", border=1, border_color=AMBER, indent=1))

    # Helper rows
    f["body"] = wb.add_format(_base(font_size=10))
    f["body_muted"] = wb.add_format(_base(font_size=9, italic=True, font_color=TEXT_DIM, text_wrap=True, valign="top"))
    f["legend_label"] = wb.add_format(_base(font_size=9, font_color=TEXT_DIM, align="left", valign="vcenter"))
    f["legend_good"] = wb.add_format(_base(font_size=9, font_color=GREY_BLACK, bg_color=TINT_GOOD, align="center", valign="vcenter", border=1, border_color=BORDER_GREY))
    f["legend_neutral"] = wb.add_format(_base(font_size=9, font_color=TEXT_DIM, bg_color=WHITE, align="center", valign="vcenter", border=1, border_color=BORDER_GREY))
    f["legend_bad"] = wb.add_format(_base(font_size=9, font_color=GREY_BLACK, bg_color=TINT_BAD, align="center", valign="vcenter", border=1, border_color=BORDER_GREY))

    # Hyperlink ("Book here")
    f["link"] = wb.add_format(_base(font_color=ACCENT_SOFT, underline=1, bold=True))
    f["link_cta"] = wb.add_format(_base(font_size=12, font_color=WHITE, bg_color=ACCENT_SOFT, underline=1, bold=True, align="center", valign="vcenter"))

    # Reconciliation Checks sheet
    f["check_difference"] = wb.add_format(_base(align="right", valign="vcenter", border=1, border_color=BORDER_GREY, bg_color=WHITE, num_format=NF_AUD))
    f["check_status_ok"] = wb.add_format(_base(bold=True, font_color="#1F5B3F", bg_color=TINT_GOOD, align="center", valign="vcenter", border=1, border_color=BORDER_GREY))
    f["check_status_flag"] = wb.add_format(_base(bold=True, font_color="#8B1A1F", bg_color=TINT_BAD, align="center", valign="vcenter", border=1, border_color=BORDER_GREY))
    f["check_status_neutral"] = wb.add_format(_base(bold=True, font_color=TEXT_DIM, bg_color=WHITE, align="center", valign="vcenter", border=1, border_color=BORDER_GREY))

    return f
