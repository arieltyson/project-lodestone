"""Valuation football field for Ero, per share in C$."""

from __future__ import annotations

from lodestone.workbook import styles as s
from lodestone.workbook.context import ModelContext
from lodestone.workbook.merger import SCEN_FIRST_ROW, SCENARIO_OUTPUTS, SCENARIOS

SHEET = "Football_Field"
LB_PER_TONNE = 2204.62


def _scen(output: str, row: int) -> str:
    return f"{SCENARIOS}!${s.col(6 + SCENARIO_OUTPUTS.index(output))}${row}"


def build(ctx: ModelContext) -> None:
    ws = ctx.wb.create_sheet(SHEET)
    s.title(ws, "Valuation football field — Ero Copper (C$ per share)",
            "Each bar is a low–high range. The offer price line comes from the Merger tab.")
    ws.column_dimensions["B"].width = 44
    for c in "CDE":
        ws.column_dimensions[c].width = 12
    ws.column_dimensions["F"].width = 80
    m = lambda field: ctx.refs[f"mkt.ERO.TO.{field}"]
    per_share_ev = lambda ev: f"(({ev})-DEAL_NET_DEBT)/ERO_SHARES/FX_SPOT"
    deck, spot = SCEN_FIRST_ROW, SCEN_FIRST_ROW + 1
    reserves_lb = f"Comps!$L$6*{LB_PER_TONNE}"
    methods = (
        ("52-week trading range", "WK52", f"={m('week52_low')}", f"={m('week52_high')}", "Market tab"),
        ("Analyst price targets", "TARGETS", f"={m('target_low')}", f"={m('target_high')}",
         "Low and high street targets (yfinance)"),
        ("Trading comps: EV / EBITDA (TTM)", "EVEBITDA", "", "",
         "Peer 25th–75th percentile x Ero TTM EBITDA, less net debt"),
        ("Trading comps: P / CF (TTM)", "PCF",
         f"=COMPS_PCF_Q1*{m('cfo_ttm_usd_m')}/ERO_SHARES/FX_SPOT",
         f"=COMPS_PCF_Q3*{m('cfo_ttm_usd_m')}/ERO_SHARES/FX_SPOT", "Peer 25th–75th percentile x Ero TTM operating cash flow"),
        ("Trading comps: EV / lb CuEq reserve", "EVLB",
         f"={per_share_ev(f'COMPS_EVLB_Q1*{reserves_lb}')}", f"={per_share_ev(f'COMPS_EVLB_Q3*{reserves_lb}')}",
         "Peer 25th–75th percentile x Ero P&P reserves (CuEq), less net debt"),
        ("Precedent premiums", "PREC", "=DEAL_UNAFFECTED*(1+PREC_PREM_Q1)", "=DEAL_UNAFFECTED*(1+PREC_PREM_Q3)",
         "25th–75th percentile premium applied to the unaffected price"),
        ("NAV at Centerra's P/NAV (deck to spot)", "NAVREL",
         f"=MIN({_scen('ERO_NAVPS_CAD', deck)}*{_scen('CG_PNAV', deck)},{_scen('ERO_NAVPS_CAD', spot)}*{_scen('CG_PNAV', spot)})",
         f"=MAX({_scen('ERO_NAVPS_CAD', deck)}*{_scen('CG_PNAV', deck)},{_scen('ERO_NAVPS_CAD', spot)}*{_scen('CG_PNAV', spot)})",
         "Relative value in a share deal: Ero NAVPS x the multiple Centerra's own stock trades at"),
        ("NAV at 1.0x (deck to spot)", "NAV",
         f"=MIN({_scen('ERO_NAVPS_CAD', deck)},{_scen('ERO_NAVPS_CAD', spot)})",
         f"=MAX({_scen('ERO_NAVPS_CAD', deck)},{_scen('ERO_NAVPS_CAD', spot)})",
         "Reserve-only NAV; excludes resources outside the mine plan"),
    )
    s.header_row(ws, 4, 2, ["Method", "Low", "High", "Offer", "Basis"])
    ebitda = m("ebitda_ttm_usd_m")
    for i, (label, name, low, high, note) in enumerate(methods):
        r = 5 + i
        if name == "EVEBITDA":
            low = f"={per_share_ev(f'COMPS_EV_EBITDA_Q1*{ebitda}')}"
            high = f"={per_share_ev(f'COMPS_EV_EBITDA_Q3*{ebitda}')}"
        s.put(ws, r, 2, label)
        s.put(ws, r, 3, low, s.FORMULA, s.USD_2)
        s.put(ws, r, 4, high, s.FORMULA, s.USD_2)
        s.put(ws, r, 5, "=DEAL_OFFER_CAD", s.LINK, s.USD_2)
        s.put(ws, r, 6, note, s.NOTE)
        s.define(ctx.wb, f"FF_{name}_LOW", SHEET, 3, r)
        s.define(ctx.wb, f"FF_{name}_HIGH", SHEET, 4, r)
    r = 6 + len(methods)
    s.put(ws, r, 2, "Offer price (C$)", s.BOLD)
    cell = s.put(ws, r, 3, "=DEAL_OFFER_CAD", s.LINK, s.USD_2)
    cell.fill = s.KEY_FILL
    s.put(ws, r + 1, 2, "Unaffected price (C$)", s.BOLD)
    s.put(ws, r + 1, 3, "=DEAL_UNAFFECTED", s.LINK, s.USD_2)
