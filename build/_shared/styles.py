"""Table, chart, sparkline and conditional-formatting helpers (xlsxwriter).

Chart styling: Lyros green primary, accent soft secondary, grey-black tertiary,
Arial titles, minimal gridlines, no chart border. Sparklines inline on trend
rows. Conditional formatting uses soft tints (pale green / white / pale red)
paired with explicit legend strips so the reader knows what each colour means.
"""

from __future__ import annotations

from typing import Iterable, Sequence

from . import branding as b


# ── Table helpers ───────────────────────────────────────────────────────────

def write_header_row(
    ws,
    formats: dict,
    *,
    row: int,
    headers: Sequence[str],
    start_col: int = 1,
    right_align_from: int | None = None,
    height: float = 26,
) -> None:
    """Write a styled header row."""
    ws.set_row(row, height)
    for i, h in enumerate(headers):
        col = start_col + i
        if right_align_from is not None and i + 1 >= right_align_from:
            ws.write(row, col, h, formats["th_right"])
        else:
            ws.write(row, col, h, formats["th"])


def write_data_row(
    ws,
    formats: dict,
    *,
    row: int,
    label: str,
    formulas: Sequence[str | None] = (),
    values: Sequence | None = None,
    start_col: int = 1,
    zebra: bool = False,
    bold: bool = False,
    cell_format: str = "td_right",
    height: float = 20,
) -> None:
    """Write a data row.

    Either `formulas` or `values` should be provided. Formulas take precedence.
    cell_format names a format from the registry (e.g. "td_right", "td_pct",
    "td_days") which already includes zebra fill via the _zebra variant.
    """
    ws.set_row(row, height)

    if zebra:
        label_fmt = formats["td_zebra"]
        if not bold:
            num_fmt_key = cell_format + "_zebra"
            num_fmt = formats.get(num_fmt_key, formats[cell_format])
        else:
            num_fmt = formats.get(f"td_bold_{cell_format.replace('td_', '')}", formats["td_bold_right"])
    else:
        label_fmt = formats["td"] if not bold else formats["td_bold_left"]
        if not bold:
            num_fmt = formats[cell_format]
        else:
            num_fmt = formats.get(f"td_bold_{cell_format.replace('td_', '')}", formats["td_bold_right"])

    ws.write(row, start_col, label, label_fmt)
    seq = formulas if formulas else (values or [])
    for i, item in enumerate(seq):
        col = start_col + 1 + i
        if formulas and item:
            ws.write_formula(row, col, item, num_fmt)
        elif values is not None:
            ws.write(row, col, item, num_fmt)
        else:
            ws.write_blank(row, col, None, num_fmt)


def write_total_row(
    ws,
    formats: dict,
    *,
    row: int,
    label: str,
    formulas: Sequence[str | None] = (),
    start_col: int = 1,
    cell_format: str = "total_right",
    height: float = 24,
) -> None:
    ws.set_row(row, height)
    ws.write(row, start_col, label, formats["total_left"])
    for i, formula in enumerate(formulas):
        col = start_col + 1 + i
        fmt = formats[cell_format]
        if formula:
            ws.write_formula(row, col, formula, fmt)
        else:
            ws.write_blank(row, col, None, fmt)


def write_section_header(ws, formats: dict, *, row: int, kicker: str, title: str) -> None:
    """Kicker on row, title below."""
    ws.set_row(row, 14)
    ws.write(row, 1, kicker.upper(), formats["section_kicker"])
    ws.set_row(row + 1, 26)
    ws.write(row + 1, 1, title, formats["section_h2"])


# ── KPI card (3 rows × N cols, soft fill block) ─────────────────────────────

def kpi_card(
    ws,
    formats: dict,
    *,
    anchor_row: int,
    anchor_col: int,
    width_cols: int,
    label: str,
    value_formula: str,
    change_formula: str | None = None,
    value_kind: str = "aud",  # aud | pct
) -> None:
    last_col = anchor_col + width_cols - 1

    ws.set_row(anchor_row, 18)
    ws.merge_range(anchor_row, anchor_col, anchor_row, last_col, label.upper(), formats["kpi_label"])

    ws.set_row(anchor_row + 1, 32)
    value_fmt = formats["kpi_value_pct"] if value_kind == "pct" else formats["kpi_value"]
    ws.merge_range(anchor_row + 1, anchor_col, anchor_row + 1, last_col, "", value_fmt)
    ws.write_formula(anchor_row + 1, anchor_col, value_formula, value_fmt)

    ws.set_row(anchor_row + 2, 18)
    ws.merge_range(anchor_row + 2, anchor_col, anchor_row + 2, last_col, "", formats["kpi_change"])
    if change_formula:
        ws.write_formula(anchor_row + 2, anchor_col, change_formula, formats["kpi_change"])


# ── Sparklines ──────────────────────────────────────────────────────────────

def add_sparkline_line(ws, *, anchor_cell: str, range_str: str, color: str = b.GREEN) -> None:
    """Add a single line sparkline at `anchor_cell` reading `range_str`.

    range_str is e.g. "Data!$C$8:$N$8" (absolute is fine).
    """
    ws.add_sparkline(anchor_cell, {
        "range": range_str,
        "type": "line",
        "series_color": color,
        "high_point": True,
        "low_point": True,
        "high_color": b.GREEN,
        "low_color": b.RED,
        "markers": False,
        "weight": 1.25,
    })


def add_sparkline_column(ws, *, anchor_cell: str, range_str: str, color: str = b.GREEN) -> None:
    """Add a single column sparkline at `anchor_cell`."""
    ws.add_sparkline(anchor_cell, {
        "range": range_str,
        "type": "column",
        "series_color": color,
        "negative_points": True,
        "negative_color": b.RED,
    })


# ── Conditional formatting with explicit legends ────────────────────────────

def add_three_color_scale(
    ws,
    cell_range: str,
    *,
    favourable_high: bool = True,
) -> None:
    """Soft 3-tone scale: pale green / white / pale red.

    favourable_high=True : high = good (green), low = bad (red). Revenue, GP, EBITDA.
    favourable_high=False: high = bad  (red),   low = good (green). Costs, ratios.
    """
    if favourable_high:
        ws.conditional_format(cell_range, {
            "type": "3_color_scale",
            "min_color": b.TINT_BAD,
            "mid_color": b.TINT_NEUTRAL,
            "max_color": b.TINT_GOOD,
            "min_type": "min",
            "mid_type": "percentile", "mid_value": 50,
            "max_type": "max",
        })
    else:
        ws.conditional_format(cell_range, {
            "type": "3_color_scale",
            "min_color": b.TINT_GOOD,
            "mid_color": b.TINT_NEUTRAL,
            "max_color": b.TINT_BAD,
            "min_type": "min",
            "mid_type": "percentile", "mid_value": 50,
            "max_type": "max",
        })


def write_cf_legend(
    ws,
    formats: dict,
    *,
    row: int,
    col: int,
    favourable_high: bool = True,
    metric_label: str = "Metric",
) -> None:
    """Write a three-cell legend strip explaining the heatmap meaning.

    Layout: <metric_label>:  [ Lower ][ Mid ][ Higher ]   meaning text
    """
    ws.set_row(row, 18)
    ws.write(row, col, f"{metric_label}:", formats["legend_label"])

    if favourable_high:
        ws.write(row, col + 1, "Lower", formats["legend_bad"])
        ws.write(row, col + 2, "Mid", formats["legend_neutral"])
        ws.write(row, col + 3, "Higher", formats["legend_good"])
        ws.write(row, col + 4, "Green = better, red = worse", formats["legend_label"])
    else:
        ws.write(row, col + 1, "Lower", formats["legend_good"])
        ws.write(row, col + 2, "Mid", formats["legend_neutral"])
        ws.write(row, col + 3, "Higher", formats["legend_bad"])
        ws.write(row, col + 4, "Green = better, red = worse", formats["legend_label"])


# ── Charts (Lyros-styled) ───────────────────────────────────────────────────

def _style_chart_common(chart, *, title: str, show_legend: bool = True) -> None:
    chart.set_title({
        "name": title,
        "name_font": {"name": b.FONT_FAMILY, "size": 12, "bold": True, "color": b.GREY_BLACK},
    })
    chart.set_x_axis({
        "num_font": {"name": b.FONT_FAMILY, "size": 9, "color": b.TEXT_DIM},
        "line": {"color": b.GRID_GREY},
        "major_gridlines": {"visible": False},
    })
    chart.set_y_axis({
        "num_font": {"name": b.FONT_FAMILY, "size": 9, "color": b.TEXT_DIM},
        "major_gridlines": {"visible": True, "line": {"color": b.OFF_WHITE, "width": 0.75}},
        "line": {"none": True},
    })
    if show_legend:
        chart.set_legend({
            "position": "bottom",
            "font": {"name": b.FONT_FAMILY, "size": 9, "color": b.TEXT_DIM},
        })
    else:
        chart.set_legend({"none": True})
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": b.WHITE}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": b.WHITE}})


def add_line_chart(
    wb,
    ws,
    *,
    title: str,
    anchor_cell: str,
    series: list[dict],
    cats_range: str,
    width: int = 720,
    height: int = 280,
) -> None:
    """Add a Lyros-branded line chart.

    series: list of dicts with keys: name, values, color (hex with #).
    Each `values` is an absolute range like "='P&L Monthly'!$C$10:$N$10".
    """
    chart = wb.add_chart({"type": "line"})
    for s in series:
        chart.add_series({
            "name": s["name"],
            "values": s["values"],
            "categories": cats_range,
            "line": {"color": s["color"], "width": 2.25},
            "marker": {
                "type": "circle", "size": 5,
                "border": {"color": s["color"]},
                "fill": {"color": s["color"]},
            },
            "smooth": False,
        })
    _style_chart_common(chart, title=title, show_legend=len(series) > 1)
    chart.set_size({"width": width, "height": height})
    ws.insert_chart(anchor_cell, chart)


def write_checks_block(
    ws,
    formats: dict,
    *,
    row: int,
    checks: list[dict],
) -> int:
    """Write a small tie-out reconciliation block at the bottom of a sheet.

    Each check is a dict with keys:
        name      str
        left      str   (Excel formula starting with '=')
        right     str   (Excel formula starting with '=')
        is_pct    bool  (default False, switches number format)

    Returns the row index after the block (so the caller can continue layout).
    """
    write_section_header(
        ws, formats, row=row,
        kicker="Reconciliation",
        title="Tie-out checks for this tab",
    )
    header_row = row + 2
    write_header_row(
        ws, formats, row=header_row,
        headers=["Check", "Left side", "Right side", "Difference", "Status"],
        right_align_from=2,
    )

    for i, chk in enumerate(checks):
        r = header_row + 1 + i
        ws.set_row(r, 22)
        zebra = i % 2 == 0
        is_pct = chk.get("is_pct", False)
        label_fmt = formats["td_zebra"] if zebra else formats["td"]
        if is_pct:
            num_fmt = formats["td_pct_zebra"] if zebra else formats["td_pct"]
        else:
            num_fmt = formats["td_right_zebra"] if zebra else formats["td_right"]
        ws.write_string(r, 1, chk["name"], label_fmt)
        ws.write_formula(r, 2, chk["left"], num_fmt)
        ws.write_formula(r, 3, chk["right"], num_fmt)
        left_cell = f"C{r + 1}"
        right_cell = f"D{r + 1}"
        ws.write_formula(r, 4, f"={left_cell}-{right_cell}", num_fmt)
        tol = "0.0001" if is_pct else "0.5"
        ws.write_formula(
            r, 5,
            f"=IF(ABS({left_cell}-{right_cell})<{tol},\"OK\",\"FLAG\")",
            formats["check_status_neutral"],
        )

    last_check_row_one = header_row + len(checks) + 1
    status_range = f"F{header_row + 2}:F{last_check_row_one}"
    ws.conditional_format(status_range, {
        "type": "text", "criteria": "containing", "value": "OK",
        "format": formats["check_status_ok"],
    })
    ws.conditional_format(status_range, {
        "type": "text", "criteria": "containing", "value": "FLAG",
        "format": formats["check_status_flag"],
    })
    return last_check_row_one


def add_column_chart(
    wb,
    ws,
    *,
    title: str,
    anchor_cell: str,
    series: list[dict],
    cats_range: str,
    width: int = 720,
    height: int = 280,
    show_legend: bool = False,
) -> None:
    chart = wb.add_chart({"type": "column"})
    for s in series:
        chart.add_series({
            "name": s["name"],
            "values": s["values"],
            "categories": cats_range,
            "fill": {"color": s["color"]},
            "border": {"none": True},
            "gap": 80,
        })
    _style_chart_common(chart, title=title, show_legend=show_legend)
    chart.set_size({"width": width, "height": height})
    ws.insert_chart(anchor_cell, chart)
