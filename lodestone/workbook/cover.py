"""Cover and Sources tabs."""

from __future__ import annotations

from lodestone.workbook import styles as s
from lodestone.workbook.context import ModelContext

TABS = (
    ("Inputs", "Price deck, FX, price-case selectors, discount rates"),
    ("Market", "yfinance snapshot: prices, shares, balance sheet, TTM metrics"),
    ("MtMilligan … Xavantina", "One life-of-mine NAV model per asset"),
    ("Corporate_NAV", "NAV roll-up, NAV per share, P/NAV and 2027–28E financial summary"),
    ("Comps", "Peer trading multiples (written by scripts/build_comps.py)"),
    ("Precedents", "Copper and gold M&A precedents: premiums, P/NAV, value per reserve ounce"),
    ("Sources", "Every filing behind the inputs, with verification status"),
)


def build_cover(ctx: ModelContext, extra_tabs: tuple[tuple[str, str], ...] = ()) -> None:
    ws = ctx.wb.active
    ws.title = "Cover"
    s.title(ws, "Project Lodestone — Centerra Gold / Ero Copper",
            "Hypothetical buy-side merger analysis. Public information only. Not investment advice. "
            "Not affiliated with any bank. See DISCLAIMER.md.")
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 90
    s.header_row(ws, 4, 2, ["Tab", "Contents"])
    row = 5
    for name, text in TABS + extra_tabs:
        s.put(ws, row, 2, name, s.BOLD)
        s.put(ws, row, 3, text)
        row += 1
    row += 1
    s.header_row(ws, row, 2, ["Colour legend", "Meaning"])
    for font, label, meaning in (
        (s.INPUT, "Blue text", "Hardcoded input (from a CSV under data/)"),
        (s.FORMULA, "Black text", "Formula"),
        (s.LINK, "Green text", "Link to another tab"),
    ):
        row += 1
        s.put(ws, row, 2, label, font)
        s.put(ws, row, 3, meaning)
    row += 1
    cell = s.put(ws, row, 2, "Yellow fill")
    cell.fill = s.KEY_FILL
    s.put(ws, row, 3, "Key assumption or scenario lever")
    row += 2
    s.put(ws, row, 2, "Rebuild", s.BOLD)
    s.put(ws, row, 3, "make model  (fetch data → build workbook → comps → scenarios → recalculate)")
    row += 1
    s.put(ws, row, 2, "Market snapshot", s.BOLD)
    s.put(ws, row, 3, ctx.a.macro["timestamp_utc"])


def build_sources(ctx: ModelContext) -> None:
    ws = ctx.wb.create_sheet("Sources")
    s.title(ws, "Sources", "Every input traces to one of these documents. Status flags anything estimated "
            "or calibrated rather than read directly from a filing.")
    sources = ctx.a.sources
    widths = (16, 18, 60, 22, 14, 50, 50, 40)
    s.header_row(ws, 4, 2, list(sources.columns))
    for j, width in enumerate(widths):
        ws.column_dimensions[s.col(2 + j)].width = width
    for i, record in enumerate(sources.itertuples(index=False)):
        for j, value in enumerate(record):
            s.put(ws, 5 + i, 2 + j, value)
    row = 7 + len(sources)
    s.section(ws, row, "Asset assumption status (data/asset_assumption_notes.csv)", 9)
    notes = ctx.a.asset_notes
    s.header_row(ws, row + 1, 2, list(notes.columns))
    for i, record in enumerate(notes.itertuples(index=False)):
        for j, value in enumerate(record):
            s.put(ws, row + 2 + i, 2 + j, value)
