"""Peer trading multiples computed from the market snapshot.

The same metrics are written into the Comps tab as live formulas by scripts/build_comps.py;
this module produces the CSV copy used for diffs, the deck and the dashboard.
"""

from __future__ import annotations

import pandas as pd

from lodestone.assumptions import Assumptions

LB_PER_TONNE = 2204.62
GRAMS_PER_OZ = 31.1035


def reserves_cueq_mt(a: Assumptions, company: str) -> float:
    """Remaining P&P reserves at 1 Jan 2027, copper-equivalent at the long-term deck."""
    deck = a.price_deck.iloc[-1]
    t_cu_per_oz = deck["gold_usd_oz"] / (deck["copper_usd_lb"] * LB_PER_TONNE)
    assets = a.assets[a.assets["company"] == company]
    ore = assets["reserve_mt"] - assets["depleted_mt"]
    cu_mt = (ore * assets["cu_grade_pct"]).sum()
    au_moz = (ore * assets["au_gpt"] / GRAMS_PER_OZ).sum()
    return float(cu_mt + au_moz * t_cu_per_oz)


def comps_table(a: Assumptions, pnav: dict[str, float] | None = None) -> pd.DataFrame:
    universe = a._csv("universe.csv").set_index("ticker")
    peers = a._csv("peer_inputs.csv").set_index("ticker")
    m = a.market.join(universe[["name", "role"]])
    out = pd.DataFrame(index=m.index)
    out["name"] = m["name"]
    out["role"] = m["role"]
    out["price_cad"] = m["price_local"]
    out["market_cap_usd_m"] = m["market_cap_usd_m"]
    out["ev_usd_m"] = m["ev_usd_m"]
    out["ev_ebitda"] = m["ev_usd_m"] / m["ebitda_ttm_usd_m"]
    out["p_cf"] = m["market_cap_usd_m"] / m["cfo_ttm_usd_m"]
    out["price_to_target"] = m["price_local"] / m["target_mean"]

    reserves = peers["pp_reserves_cueq_mt"].reindex(m.index)
    for ticker, role in universe["role"].items():
        if role in ("acquirer", "target"):
            company = a.corporate.index[a.corporate["ticker"] == ticker][0]
            reserves[ticker] = reserves_cueq_mt(a, company)
    out["reserves_cueq_mt"] = reserves
    out["ev_per_lb_cueq"] = m["ev_usd_m"] / (reserves * LB_PER_TONNE)

    navps = peers["consensus_navps_cad"].reindex(m.index)
    out["p_nav"] = m["price_local"] / navps
    for ticker, value in (pnav or {}).items():
        out.loc[ticker, "p_nav"] = value
    return out


def peer_stats(table: pd.DataFrame, metrics=("ev_ebitda", "p_cf", "ev_per_lb_cueq", "price_to_target")) -> pd.DataFrame:
    """Quartiles across peers only (acquirer and target excluded), matching Excel QUARTILE."""
    peers = table[table["role"] == "peer"]
    rows = {}
    for metric in metrics:
        series = peers[metric].dropna()
        rows[metric] = {
            "mean": series.mean(),
            "median": series.median(),
            "q1": series.quantile(0.25),
            "q3": series.quantile(0.75),
        }
    return pd.DataFrame(rows)
