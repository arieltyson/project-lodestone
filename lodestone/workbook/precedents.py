"""Precedent transactions tab (data/precedents.csv)."""

from __future__ import annotations

from lodestone.workbook import styles as s
from lodestone.workbook.context import ModelContext

SHEET = "Precedents"
COLUMNS = (
    ("announced", "Announced", 11, None),
    ("acquirer", "Acquirer", 22, None),
    ("target", "Target", 22, None),
    ("primary_metal", "Metal", 12, None),
    ("consideration", "Consideration", 14, None),
    ("equity_value_usd_m", "Equity value (US$M)", 12, s.USD_0),
    ("premium_spot", "Premium", 9, s.PCT),
    ("p_nav", "P/NAV at ann.", 10, s.MULT),
    ("target_reserves_moz_aueq", "Reserves (Moz AuEq)", 11, s.USD_2),
    ("ev_per_oz", "Value / reserve oz (US$)", 12, s.USD_0),
    ("status", "Status", 50, None),
)


def build(ctx: ModelContext) -> None:
    deals = ctx.a._csv("precedents.csv")
    ws = ctx.wb.create_sheet(SHEET)
    s.title(ws, "Precedent transactions — copper and gold M&A, 2021–2025",
            "Premium to the target's undisturbed price at announcement, from company press releases. "
            "Blank = not disclosed or not yet verified.")
    s.header_row(ws, 4, 2, [label for _, label, _, _ in COLUMNS])
    for j, (_, _, width, _) in enumerate(COLUMNS):
        ws.column_dimensions[s.col(2 + j)].width = width
    first = 5
    for i, deal in deals.iterrows():
        r = first + i
        for j, (field, _, _, fmt) in enumerate(COLUMNS):
            if field == "ev_per_oz":
                value, font = f'=IF(J{r}>0,G{r}/J{r},"n/a")', s.FORMULA
            else:
                value = deal[field]
                value = None if value != value else (value.item() if hasattr(value, "item") else value)
                font = s.INPUT
            s.put(ws, r, 2 + j, value, font, fmt)
        s.put(ws, r, 2 + len(COLUMNS), deal["source_url"], s.NOTE)
    last = first + len(deals) - 1
    r = last + 2
    for label, fn, suffix in (("Mean", "AVERAGE({})", "MEAN"), ("Median", "MEDIAN({})", "MED"),
                              ("25th percentile", "QUARTILE({},1)", "Q1"),
                              ("75th percentile", "QUARTILE({},3)", "Q3")):
        s.put(ws, r, 2, label, s.BOLD)
        for c, name, fmt in ((8, "PREM", s.PCT), (9, "PNAV", s.MULT), (11, "EVOZ", s.USD_0)):
            L = s.col(c)
            cell = s.put(ws, r, c, f'=IFERROR({fn.format(f"{L}{first}:{L}{last}")},"n/a")', s.BOLD, fmt)
            cell.border = s.TOTAL_BORDER
            s.define(ctx.wb, f"PREC_{name}_{suffix}", SHEET, c, r)
        r += 1
    r += 1
    s.put(ws, r, 2, "Benchmarks", s.BOLD)
    for note in (
        "FactSet: gold targets acquired in 2025 averaged ~0.73x consensus P/NAV (13 deals); ten-year average ~0.8x.",
        "FactSet: Harmony / MAC Copper implied 1.26x P/NAV and Anglo / Teck 1.22x, the highest since BHP's bid for Anglo.",
        "Source: insight.factset.com/metals-mining-2025-ma-and-equity-capital-markets-insights",
    ):
        r += 1
        s.put(ws, r, 2, note, s.NOTE)
