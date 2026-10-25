"""Project Lodestone deal dashboard.

Flex the copper and gold prices, CAD/USD, offer premium, consideration mix and synergies,
and watch Ero's NAV, the football field and Centerra's accretion move. All math lives in
lodestone/valuation.py, which is reconciled to the Excel model by the test suite.

Run locally:
    streamlit run app/dashboard.py
"""

from __future__ import annotations

import sys
from dataclasses import replace
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from lodestone.assumptions import Assumptions  # noqa: E402
from lodestone.comps import comps_table, peer_stats  # noqa: E402
from lodestone.valuation import (  # noqa: E402
    DealTerms, Prices, eps_sensitivity, football_field, navps_sensitivity, run,
)

# Categorical pair validated for CVD separation and contrast on a light surface.
CG_COLOR, ERO_COLOR = "#2F6FB0", "#B5652B"
GOOD, BAD, NEUTRAL = "#2F6FB0", "#B23A3A", "#EDEDED"
INK, MUTED, GRID = "#1E2933", "#5F6B76", "#E6E9EC"
METHOD_LABELS = {
    "WK52": "52-week trading range", "TARGETS": "Analyst price targets", "EVEBITDA": "Comps: EV / EBITDA (TTM)",
    "PCF": "Comps: P / CF (TTM)", "EVLB": "Comps: EV / lb CuEq reserve", "PREC": "Precedent premiums",
    "NAVREL": "NAV at Centerra's P/NAV", "NAV": "NAV at 1.0x (deck to spot)",
}

st.set_page_config(page_title="Project Lodestone", page_icon="⛏️", layout="wide")


@st.cache_resource
def assumptions() -> Assumptions:
    return Assumptions()


def layout(fig: go.Figure, title: str, height: int = 420) -> go.Figure:
    fig.update_layout(
        title=dict(text=title, font=dict(size=15, color=INK), x=0), height=height, margin=dict(l=10, r=10, t=50, b=10),
        plot_bgcolor="white", paper_bgcolor="white", font=dict(color=INK), hoverlabel=dict(bgcolor="white"),
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1),
    )
    fig.update_xaxes(gridcolor=GRID, zeroline=False, color=MUTED)
    fig.update_yaxes(gridcolor=GRID, zeroline=False, color=MUTED)
    return fig


def pct(v: float) -> str:
    return f"{v:+.1%}"


a = assumptions()
base_terms = DealTerms.from_assumptions(a)
deck = Prices.deck(a)

# --- Sidebar controls ---------------------------------------------------------------------
st.sidebar.title("Deal levers")
price_mode = st.sidebar.radio("Commodity prices", ["Analyst deck", "Flat prices"], horizontal=True,
                              help="Analyst deck steps down from 2027 to a long-term price. Flat holds one price every year.")
copper = st.sidebar.slider("Copper (US\\$/lb)", 3.00, 8.00, float(deck.copper[-1]), 0.05,
                           disabled=price_mode == "Analyst deck")
gold = st.sidebar.slider("Gold (US\\$/oz)", 2000, 5000, int(deck.gold[-1]), 50, disabled=price_mode == "Analyst deck")
fx = st.sidebar.slider("CAD/USD (US\\$ per C\\$)", 0.65, 0.80, round(float(a.macro["cadusd"]), 3), 0.005, format="%.3f")
st.sidebar.divider()
premium = st.sidebar.slider("Offer premium", 0.0, 0.60, float(base_terms.premium), 0.01, format="%.2f")
all_stock = st.sidebar.toggle("All-stock offer", value=base_terms.cash_pct == 0)
cash_pct = 0.0 if all_stock else st.sidebar.slider("Cash portion of offer", 0.0, 0.60, float(base_terms.cash_pct) or 0.25,
                                                    0.05, format="%.2f")
synergies = st.sidebar.slider("Run-rate synergies (US\\$M, pre-tax)", 0, 300, int(base_terms.synergies), 5)
st.sidebar.caption(f"Market snapshot {a.macro['timestamp_utc']} · spot Cu US\\${a.macro['copper_usd_lb']:.2f}/lb · "
                   f"Au US\\${a.macro['gold_usd_oz']:,.0f}/oz")

prices = replace(deck, cadusd=fx) if price_mode == "Analyst deck" else Prices.flat(a, copper, gold, fx)
terms = replace(base_terms, premium=premium, cash_pct=cash_pct, synergies=float(synergies))
model = run(a, prices, terms)
cg, ero, deal = model.companies["CG"], model.companies["ERO"], model.deal

# --- Header -----------------------------------------------------------------------------
st.title("Project Lodestone — Centerra Gold / Ero Copper")
st.caption("Hypothetical merger analysis built from public information. Not investment advice. "
           "Not affiliated with any bank or with either company.")
m = st.columns(5)
m[0].metric("Offer per Ero share", f"C${deal.offer_cad:,.2f}", f"{premium:.0%} premium", delta_color="off")
m[1].metric("2027E EPS accretion", pct(deal.years[2027]["acc_eps"]))
m[2].metric("2028E CFPS accretion", pct(deal.years[2028]["acc_cfps"]))
m[3].metric("NAVPS accretion", pct(deal.acc_navps))
m[4].metric("Pro forma leverage", f"{deal.leverage:.1f}x", "net debt / 2027E EBITDA", delta_color="off")

nav_tab, comps_tab, deal_tab = st.tabs(["NAV", "Comps & Football Field", "Deal Structuring"])

# --- NAV tab ----------------------------------------------------------------------------
with nav_tab:
    rows = []
    for co_name, co in (("Centerra", cg), ("Ero", ero)):
        for key, res in co.assets.items():
            rows.append({"Company": co_name, "Asset": a.assets.loc[key, "display_name"], "NAV (US$M)": res.nav,
                         "Reserve life (yrs)": res.life, "Discount rate": float(a.discount_rates[key]),
                         "AISC (LOM, by-product)": res.aisc_lom})
    assets_df = pd.DataFrame(rows).sort_values("NAV (US$M)")
    left, right = st.columns([3, 2])
    with left:
        fig = go.Figure()
        for co_name, color in (("Centerra", CG_COLOR), ("Ero", ERO_COLOR)):
            sub = assets_df[assets_df["Company"] == co_name]
            fig.add_bar(y=sub["Asset"], x=sub["NAV (US$M)"], orientation="h", name=co_name, marker_color=color,
                        text=[f"{v:,.0f}" for v in sub["NAV (US$M)"]], textposition="outside",
                        hovertemplate="%{y}: US$%{x:,.0f}M<extra></extra>")
        st.plotly_chart(layout(fig, "Attributable asset NAV (US$M)"), width="stretch")
    with right:
        corp = pd.DataFrame({
            "Centerra": [cg.asset_nav, cg.ga, cg.other, cg.cash, cg.debt, cg.nav, cg.navps_cad, cg.price_cad, cg.pnav],
            "Ero": [ero.asset_nav, ero.ga, ero.other, ero.cash, ero.debt, ero.nav, ero.navps_cad, ero.price_cad, ero.pnav],
        }, index=["Asset NAV (US$M)", "Corporate G&A (US$M)", "Other assets (US$M)", "Cash (US$M)", "Debt (US$M)",
                  "NAV (US$M)", "NAV per share (C$)", "Share price (C$)", "P/NAV (x)"])
        st.subheader("Corporate NAV")
        st.dataframe(corp.style.format("{:,.2f}"), width="stretch")
    with st.expander("Asset table"):
        st.dataframe(assets_df.style.format({"NAV (US$M)": "{:,.0f}", "Discount rate": "{:.1%}",
                                             "AISC (LOM, by-product)": "{:,.2f}"}), width="stretch",
                     hide_index=True)

    years = a.years[:16]
    fig = go.Figure()
    lb_per_oz = float(deck.gold[-1] / deck.copper[-1])
    for co_name, co, color in (("Centerra", cg, CG_COLOR), ("Ero", ero, ERO_COLOR)):
        aueq = sum((r.cashflows["au_oz"] + r.cashflows["cu_lb"] * 1000 / lb_per_oz) * r.ownership
                   for r in co.assets.values())
        fig.add_bar(x=[str(y) for y in years], y=aueq.loc[years], name=co_name, marker_color=color,
                    hovertemplate=f"{co_name} %{{x}}: %{{y:,.0f}} koz<extra></extra>")
    fig.update_layout(barmode="stack", bargap=0.25)
    st.plotly_chart(layout(fig, "Pro forma attributable production (koz gold-equivalent)", 360),
                    width="stretch")

# --- Comps & football field ---------------------------------------------------------------
with comps_tab:
    ff = football_field(a, model)
    fig = go.Figure()
    labels = [METHOD_LABELS[k] for k in ff.index][::-1]
    lows, highs = ff["low"].to_numpy()[::-1], ff["high"].to_numpy()[::-1]
    fig.add_bar(y=labels, x=highs - lows, base=lows, orientation="h", marker_color=ERO_COLOR, name="Range",
                customdata=list(zip(lows, highs)),
                hovertemplate="%{y}<br>C$%{customdata[0]:.2f} – C$%{customdata[1]:.2f}<extra></extra>")
    fig.add_vline(x=deal.offer_cad, line=dict(color=INK, dash="dash", width=2),
                  annotation_text=f"Offer C${deal.offer_cad:.2f}", annotation_position="top right",
                  annotation=dict(yshift=22))
    fig.add_vline(x=ero.price_cad, line=dict(color=MUTED, dash="dot", width=1.5),
                  annotation_text=f"Unaffected C${ero.price_cad:.2f}", annotation_position="top left",
                  annotation=dict(yshift=22))
    fig.update_layout(showlegend=False, bargap=0.35)
    fig.update_xaxes(tickprefix="C$")
    fig = layout(fig, "Valuation football field — Ero (C$ per share)", 480)
    fig.update_layout(margin=dict(t=80))
    st.plotly_chart(fig, width="stretch")
    with st.expander("Football field data"):
        st.dataframe(ff.rename(index=METHOD_LABELS).style.format("C${:,.2f}"), width="stretch")

    table = comps_table(a, {"CG.TO": cg.pnav, "ERO.TO": ero.pnav})
    table["p_nav"] = pd.to_numeric(table["p_nav"], errors="coerce")
    stats = peer_stats(table)
    st.subheader("Trading comparables")
    shown = table.rename(columns={
        "name": "Company", "role": "Role", "price_cad": "Price (C$)", "market_cap_usd_m": "Mkt cap (US$M)",
        "ev_usd_m": "EV (US$M)", "ev_ebitda": "EV/EBITDA", "p_cf": "P/CF", "p_nav": "P/NAV",
        "price_to_target": "Price/target", "reserves_cueq_mt": "Reserves (Mt CuEq)", "ev_per_lb_cueq": "EV/lb CuEq",
    })
    numeric = [c for c in shown.columns if c not in ("Company", "Role")]
    shown[numeric] = shown[numeric].map(lambda v: "n/a" if pd.isna(v) else f"{v:,.2f}")
    st.dataframe(shown, width="stretch")
    st.caption("Peer statistics (peers only): " + " · ".join(
        f"{metric}: median {stats.loc['median', metric]:.2f}" for metric in ("ev_ebitda", "p_cf", "ev_per_lb_cueq")
    ) + ". Peer reserves are estimates; P/NAV is shown only where a NAV is modelled.")

# --- Deal structuring ---------------------------------------------------------------------
with deal_tab:
    left, right = st.columns(2)
    with left:
        cats = ["EPS 2027E", "EPS 2028E", "CFPS 2027E", "CFPS 2028E", "NAVPS"]
        vals = [deal.years[2027]["acc_eps"], deal.years[2028]["acc_eps"], deal.years[2027]["acc_cfps"],
                deal.years[2028]["acc_cfps"], deal.acc_navps]
        fig = go.Figure(go.Bar(x=cats, y=vals, marker_color=[GOOD if v >= 0 else BAD for v in vals],
                               text=[pct(v) for v in vals], textposition="outside",
                               hovertemplate="%{x}: %{y:+.1%}<extra></extra>"))
        fig.update_yaxes(tickformat="+.0%")
        st.plotly_chart(layout(fig, "Accretion / (dilution) to Centerra shareholders"), width="stretch")
    with right:
        terms_df = pd.DataFrame({"Value": [
            f"C${deal.offer_cad:,.2f}", f"{deal.exchange_ratio:.4f}x", f"US${deal.equity_usd:,.0f}M",
            f"US${deal.ev_usd:,.0f}M", f"{deal.pnav:.2f}x", f"{deal.ev_ebitda:.1f}x", f"US${deal.cash_usd:,.0f}M",
            f"US${deal.new_debt:,.0f}M", f"{deal.own_cg:.1%} / {1 - deal.own_cg:.1%}", f"US${deal.syn_npv:,.0f}M",
        ]}, index=["Offer price", "Exchange ratio", "Equity value", "Enterprise value", "Implied P/NAV",
                   "EV / 2027E EBITDA", "Cash consideration", "New debt", "Ownership (Centerra / Ero)",
                   "Synergy NPV (after tax, net of costs)"])
        st.subheader("Transaction summary")
        st.dataframe(terms_df, width="stretch")

    st.subheader("Sensitivities")
    s1, s2 = st.columns(2)
    premiums = (0.15, 0.20, 0.25, 0.30, 0.35, 0.40)
    coppers = (4.00, 4.50, 5.00, 5.50, 6.00, 6.50)
    navps = navps_sensitivity(a, premiums, coppers, terms, prices)
    eps = eps_sensitivity(model, (0, 50, 100, 155, 200), (0.0, 0.15, 0.25, 0.35, 0.50))

    def heatmap(grid: pd.DataFrame, title: str, x_label: str, y_label: str, x_fmt, y_fmt) -> go.Figure:
        z = grid.to_numpy(float)
        bound = max(abs(z.min()), abs(z.max())) or 1
        fig = go.Figure(go.Heatmap(
            z=z, x=[x_fmt(c) for c in grid.columns], y=[y_fmt(i) for i in grid.index], zmin=-bound, zmax=bound,
            colorscale=[[0, BAD], [0.5, NEUTRAL], [1, GOOD]], text=[[pct(v) for v in row] for row in z],
            texttemplate="%{text}", hovertemplate=f"{y_label} %{{y}}, {x_label} %{{x}}: %{{text}}<extra></extra>",
            showscale=False, xgap=2, ygap=2,
        ))
        fig.update_xaxes(title=x_label, showgrid=False, type="category")
        fig.update_yaxes(title=y_label, showgrid=False, autorange="reversed", type="category")
        return layout(fig, title, 380)

    with s1:
        st.plotly_chart(heatmap(navps, "NAVPS accretion: premium vs flat copper", "Copper (US$/lb)", "Premium",
                                lambda c: f"${c:.2f}", lambda p: f"{p:.0%}"), width="stretch")
    with s2:
        st.plotly_chart(heatmap(eps, "2028E EPS accretion: synergies vs cash portion", "Cash portion",
                                "Synergies (US$M)", lambda c: f"{c:.0%}", lambda s: f"${s:,.0f}M"),
                        width="stretch")
    with st.expander("Sensitivity data"):
        st.dataframe(navps.style.format("{:+.1%}"), width="stretch")
        st.dataframe(eps.style.format("{:+.1%}"), width="stretch")
    st.caption("NAVPS sensitivity holds gold at the selected price case and varies copper as a flat price.")
