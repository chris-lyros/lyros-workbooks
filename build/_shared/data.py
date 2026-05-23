"""Synthetic Australian-shaped data shared across all Lyros library workbooks.

One roster, one history window, one set of conventions. Every published
workbook draws from this module so the library reads as a single product line.

Conventions:
- ABNs are 11-digit synthetic numbers prefixed `99` so they cannot collide
  with real ABNs.
- Company names are Australian place names paired with industry-typical
  suffixes (Pty Ltd, Co, Group).
- AUD figures sit in the SME range ($100k to $50M revenue annualised).
- History anchors to a build date 6 months before today so the workbook does
  not feel stale for at least 12 months from publication.

This module is deterministic: a fixed seed gives the same numbers every run.
That is intentional — published workbooks are reproducible artefacts.
"""

from __future__ import annotations

import random
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Iterable


SEED = 20260523
ANCHOR_BUILD_DATE = date(2026, 5, 23)
# Last reported month is 6 months before the anchor so the data ages cleanly.
LAST_REPORTED_MONTH = date(2025, 11, 30)


# ── Synthetic Australian companies (the canonical roster) ───────────────────
# Mix of industries so a single roster works across debtors, suppliers,
# payroll, and revenue contexts.

COMPANIES = [
    # name,                          industry,            abn,           state
    ("Yarra Supplies Pty Ltd",       "Wholesale",         "99 102 478 311", "VIC"),
    ("Southbank Trading Co",         "Retail",            "99 211 884 027", "VIC"),
    ("Latrobe Manufacturing Pty Ltd","Manufacturing",     "99 326 119 540", "VIC"),
    ("Brunswick Logistics",          "Transport",         "99 414 067 882", "VIC"),
    ("Geelong Foods Pty Ltd",        "Food production",   "99 568 922 134", "VIC"),
    ("Carlton Constructions",        "Construction",      "99 633 745 091", "VIC"),
    ("Docklands Digital Pty Ltd",    "Software",          "99 705 281 446", "VIC"),
    ("Fitzroy Media Group",          "Media",             "99 819 037 612", "VIC"),
    ("Williamstown Marine Co",       "Marine services",   "99 902 156 738", "VIC"),
    ("Northbridge Holdings Pty Ltd", "Investment",        "99 037 612 845", "NSW"),
    ("Parramatta Plumbing Pty Ltd",  "Trades",            "99 154 803 226", "NSW"),
    ("Bondi Brew Co",                "Hospitality",       "99 261 045 718", "NSW"),
    ("Pyrmont Print Services",       "Print",             "99 388 226 905", "NSW"),
    ("Surry Hills Studio Pty Ltd",   "Creative agency",   "99 492 803 117", "NSW"),
    ("Fortitude Valley Cafe Group",  "Hospitality",       "99 514 062 837", "QLD"),
    ("Toowoomba Agri Pty Ltd",       "Agriculture",       "99 627 184 309", "QLD"),
    ("Adelaide Hills Wine Co",       "Wine production",   "99 738 502 144", "SA"),
    ("Fremantle Freight Pty Ltd",    "Transport",         "99 845 270 661", "WA"),
    ("Hobart Heritage Pty Ltd",      "Property",          "99 956 813 224", "TAS"),
    ("Canberra Consulting Group",    "Consulting",        "99 077 614 850", "ACT"),
    ("Darwin Defence Services",      "Government",        "99 188 026 437", "NT"),
    ("Sunshine Coast Fitness Co",    "Health services",   "99 290 478 016", "QLD"),
    ("Newcastle Steel Pty Ltd",      "Manufacturing",     "99 314 882 750", "NSW"),
    ("Ballarat Building Supplies",   "Wholesale",         "99 425 996 308", "VIC"),
    ("Mornington Marine Pty Ltd",    "Marine services",   "99 539 071 642", "VIC"),
]


@dataclass(frozen=True)
class Company:
    name: str
    industry: str
    abn: str
    state: str


def roster() -> list[Company]:
    """Return the canonical company roster as Company objects."""
    return [Company(*row) for row in COMPANIES]


# ── Date helpers ────────────────────────────────────────────────────────────

def month_end(d: date) -> date:
    """Return the last day of the month containing d."""
    if d.month == 12:
        return date(d.year, 12, 31)
    nxt = date(d.year, d.month + 1, 1)
    return nxt - timedelta(days=1)


def months_back(anchor: date, n: int) -> list[date]:
    """Return n month-end dates ending at the month of anchor (oldest first)."""
    dates: list[date] = []
    y, m = anchor.year, anchor.month
    for _ in range(n):
        dates.append(month_end(date(y, m, 1)))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return list(reversed(dates))


def weeks_forward(start: date, n: int) -> list[date]:
    """Return n week-start dates beginning at start (Monday-aligned)."""
    monday = start - timedelta(days=start.weekday())
    return [monday + timedelta(weeks=i) for i in range(n)]


# ── AUD figure helpers ──────────────────────────────────────────────────────

def aud(rng: random.Random, low: float, high: float, *, ndigits: int = 0) -> float:
    """Return a synthetic AUD figure in [low, high], rounded to ndigits."""
    return round(rng.uniform(low, high), ndigits)


def make_rng(salt: str = "") -> random.Random:
    """Return a deterministic RNG keyed off SEED and salt.

    Use a distinct salt per logical data series so different sheets do not
    accidentally produce identical numbers.
    """
    return random.Random(f"{SEED}-{salt}")


# ── Series shaped like a real SME P&L ──────────────────────────────────────

def revenue_series(months: int, rng: random.Random, base: float = 280_000) -> list[float]:
    """Twelve-or-N months of revenue with mild seasonality and trend."""
    series: list[float] = []
    for i in range(months):
        trend = 1.0 + (i / max(months - 1, 1)) * 0.18
        season = 1.0 + 0.08 * (1 if i % 3 == 0 else -1 if i % 3 == 2 else 0)
        noise = rng.uniform(-0.05, 0.05)
        val = base * trend * season * (1 + noise)
        series.append(round(val, 0))
    return series


def cogs_series(revenue: Iterable[float], rng: random.Random, gm_pct: float = 0.42) -> list[float]:
    """COGS as a percentage of revenue with small monthly variance."""
    out: list[float] = []
    for r in revenue:
        v = r * (1 - gm_pct + rng.uniform(-0.03, 0.03))
        out.append(round(v, 0))
    return out


def opex_series(months: int, rng: random.Random, base: float = 110_000) -> list[float]:
    """Opex band that drifts upward modestly."""
    out: list[float] = []
    for i in range(months):
        trend = 1.0 + (i / max(months - 1, 1)) * 0.08
        noise = rng.uniform(-0.04, 0.04)
        out.append(round(base * trend * (1 + noise), 0))
    return out


def wages_series(revenue: Iterable[float], rng: random.Random, ratio: float = 0.22) -> list[float]:
    """Wages as a ratio of revenue with mild fluctuation."""
    return [round(r * (ratio + rng.uniform(-0.015, 0.015)), 0) for r in revenue]
