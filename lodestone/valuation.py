"""Python mirror of the Excel model's NAV, merger and football-field math.

Kept free of any UI code so the dashboard, the tests and ad-hoc analysis all share one
engine. tests/test_reconciliation.py checks every output here against the recalculated
workbook within 0.5%.

Conventions match the workbook: US$M, real terms, costs negative, mid-year discounting
from 1 January of the first model year.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace

import numpy as np
import pandas as pd

from lodestone.assumptions import COMPANIES, Assumptions
from lodestone.comps import LB_PER_TONNE, comps_table, peer_stats, reserves_cueq_mt

GRAMS_PER_OZ = 31.1035
FIN_YEARS = (2027, 2028)


# --- Prices -----------------------------------------------------------------------------

@dataclass(frozen=True)
class Prices:
    """Copper (US$/lb) and gold (US$/oz) by model year, plus spot CAD/USD."""

    copper: np.ndarray
    gold: np.ndarray
    cadusd: float

    @classmethod
    def deck(cls, a: Assumptions, cadusd: float | None = None) -> "Prices":
        return cls(a.price_deck["copper_usd_lb"].to_numpy(float), a.price_deck["gold_usd_oz"].to_numpy(float),
                   float(cadusd if cadusd is not None else a.macro["cadusd"]))

    @classmethod
    def flat(cls, a: Assumptions, copper: float, gold: float, cadusd: float | None = None) -> "Prices":
        n = len(a.years)
        return cls(np.full(n, float(copper)), np.full(n, float(gold)),
                   float(cadusd if cadusd is not None else a.macro["cadusd"]))

    @classmethod
    def spot(cls, a: Assumptions) -> "Prices":
        return cls.flat(a, a.macro["copper_usd_lb"], a.macro["gold_usd_oz"])

    def with_copper(self, copper: float) -> "Prices":
        return replace(self, copper=np.full_like(self.copper, float(copper)))


# --- Asset NAV --------------------------------------------------------------------------

def asset_cashflows(a: Assumptions, asset: str, prices: Prices) -> pd.DataFrame:
    """Annual life-of-mine model for one asset (100% basis), one row per year."""
    p = a.assets.loc[asset]
    years = np.array(a.years)
    rate = float(a.discount_rates[asset])

    capacity = np.where(years < p["start_year"], 0.0,
                        np.where(years >= p["expansion_year"], p["expansion_mtpa"], p["throughput_mtpa"]))
    ore = np.zeros(len(years))
    remaining = p["reserve_mt"] - p["depleted_mt"]
    for i, cap in enumerate(capacity):
        ore[i] = max(0.0, min(cap, remaining))
        remaining -= ore[i]

    cu_lb = ore * p["cu_grade_pct"] * p["cu_recovery"] * p["cu_payable"] * LB_PER_TONNE
    au_oz = ore * p["au_gpt"] * p["au_recovery"] * p["au_payable"] * 1000 / GRAMS_PER_OZ
    cu_real = prices.copper * (1 - p["stream_cu_pct"]) + prices.copper * p["stream_cu_pct"] * p["stream_cu_spot_pct"]
    au_real = prices.gold * (1 - p["stream_au_pct"]) + p["stream_au_pct"] * p["stream_au_price"]
    cu_rev = cu_lb * cu_real
    au_rev = au_oz * au_real / 1000
    rev = cu_rev + au_rev
    tcrc = -cu_lb * p["tcrc_usd_lb"]
    royalty = -rev * p["nsr_royalty"]
    opex = -ore * p["opex_usd_t"]
    ebitda = rev + tcrc + royalty + opex
    sust = -np.where(ore > 0, p["sustaining_usd_m"], 0.0)
    in_growth = (years >= p["growth_start"]) & (years < p["growth_start"] + p["growth_years"])
    growth = -np.where(in_growth, p["growth_capex_usd_m"] / p["growth_years"], 0.0)
    prev_ore = np.concatenate([[0.0], ore[:-1]])
    closure = -np.where((ore == 0) & (prev_ore > 0), p["closure_usd_m"], 0.0)
    closure[0] = 0.0
    tax = -np.maximum(0.0, (ebitda + sust + growth) * p["tax_rate"])
    fcf = ebitda + sust + growth + closure + tax
    t = years - a.first_year + 1
    df = 1 / (1 + rate) ** (t - 0.5)
    return pd.DataFrame({
        "year": years, "ore": ore, "cu_lb": cu_lb, "au_oz": au_oz, "cu_rev": cu_rev, "au_rev": au_rev, "rev": rev,
        "tcrc": tcrc, "royalty": royalty, "opex": opex, "ebitda": ebitda, "sust": sust, "growth": growth,
        "closure": closure, "tax": tax, "fcf": fcf, "df": df, "pv": fcf * df,
    }).set_index("year")


def _aisc(cf: pd.DataFrame, gold: bool) -> float:
    cost = -(cf["tcrc"] + cf["royalty"] + cf["opex"] + cf["sust"]).sum()
    if gold:
        oz = cf["au_oz"].sum()
        return (cost - cf["cu_rev"].sum()) / oz * 1000 if oz else 0.0
    lb = cf["cu_lb"].sum()
    return (cost - cf["au_rev"].sum()) / lb if lb else 0.0


@dataclass
class AssetResult:
    asset: str
    cashflows: pd.DataFrame
    ownership: float
    npv: float
    nav: float
    life: int
    aisc_y1: float
    aisc_lom: float


def value_asset(a: Assumptions, asset: str, prices: Prices) -> AssetResult:
    cf = asset_cashflows(a, asset, prices)
    p = a.assets.loc[asset]
    gold = p["primary_metal"].startswith("gold")
    npv = float(cf["pv"].sum())
    return AssetResult(
        asset=asset, cashflows=cf, ownership=float(p["ownership"]), npv=npv, nav=npv * float(p["ownership"]),
        life=int((cf["ore"] > 0).sum()), aisc_y1=_aisc(cf.iloc[[0]], gold), aisc_lom=_aisc(cf, gold),
    )


# --- Corporate NAV and financials -------------------------------------------------------

@dataclass
class CompanyResult:
    company: str
    assets: dict[str, AssetResult]
    asset_nav: float
    ga: float
    other: float
    hedges: float
    cash: float
    debt: float
    nav: float
    shares: float
    navps_usd: float
    navps_cad: float
    price_cad: float
    pnav: float
    fin: dict[int, dict[str, float]] = field(default_factory=dict)


def value_company(a: Assumptions, company: str, prices: Prices) -> CompanyResult:
    corp = a.corporate.loc[company]
    mkt = a.market.loc[corp["ticker"]]
    assets = {k: value_asset(a, k, prices) for k in a.assets_of(company)}
    asset_nav = sum(r.nav for r in assets.values())
    r, n = float(a.settings["ga_discount_rate"]), float(a.settings["ga_capitalisation_years"])
    ga = -corp["ga_usd_m"] * (1 - (1 + r) ** -n) / r
    cash, debt = float(mkt["cash_usd_m"]), -float(mkt["debt_usd_m"])
    nav = asset_nav + ga + corp["other_assets_usd_m"] + corp["hedge_mtm_usd_m"] + cash + debt
    shares = float(mkt["shares_out"]) / 1e6
    navps_usd = nav / shares
    navps_cad = navps_usd / prices.cadusd
    price = float(mkt["price_local"])
    result = CompanyResult(company, assets, asset_nav, ga, float(corp["other_assets_usd_m"]),
                           float(corp["hedge_mtm_usd_m"]), cash, debt, nav, shares, navps_usd, navps_cad, price,
                           price / navps_cad)
    for year in FIN_YEARS:
        attrib = lambda key: sum(res.cashflows.loc[year, key] * res.ownership for res in assets.values())
        ebitda = attrib("ebitda") - corp["ga_usd_m"]
        dna = -corp[f"dna_{year}_usd_m"]
        interest = debt * corp["interest_rate"]
        ebt = ebitda + dna + interest
        ni = ebt - max(0.0, ebt) * corp["book_tax_rate"]
        cfo = ebitda + interest + attrib("tax")
        result.fin[year] = {
            "cu_lb": attrib("cu_lb"), "au_oz": attrib("au_oz"), "rev": attrib("rev"), "ebitda": ebitda,
            "ni": ni, "eps": ni / shares, "cfo": cfo, "cfps": cfo / shares, "cash_tax": attrib("tax"),
        }
    return result


# --- Deal -------------------------------------------------------------------------------

@dataclass(frozen=True)
class DealTerms:
    premium: float
    cash_pct: float  # cash share of the offer actually applied (0 for all-stock)
    synergies: float  # pre-tax run-rate, US$M
    min_cash: float
    debt_rate: float
    cash_yield: float
    txn_costs: float
    tax: float
    phase_y1: float
    phase_y2: float
    integration_x: float
    syn_years: float
    syn_rate: float

    @classmethod
    def from_assumptions(cls, a: Assumptions, **overrides) -> "DealTerms":
        d = a._csv("deal_assumptions.csv").set_index("key")["value"].astype(float)
        terms = cls(
            premium=d["offer_premium"],
            cash_pct=0.0 if d["consideration_case"] == 1 else d["cash_pct"],
            synergies=float(a._csv("synergies.csv")["run_rate_usd_m"].sum()),
            min_cash=d["min_cash_usd_m"], debt_rate=d["new_debt_rate"], cash_yield=d["cash_yield"],
            txn_costs=d["transaction_costs_usd_m"], tax=d["deal_tax_rate"], phase_y1=d["synergy_phase_in_y1"],
            phase_y2=d["synergy_phase_in_y2"], integration_x=d["integration_cost_multiple"],
            syn_years=d["synergy_npv_years"], syn_rate=d["synergy_discount_rate"],
        )
        return replace(terms, **overrides)


def synergy_npv(t: DealTerms) -> float:
    r, n = t.syn_rate, t.syn_years
    annuity = (1 - (1 + r) ** -n) / r * (1 + r) ** 0.5
    shortfall = t.synergies * (1 - t.phase_y1) * (1 - t.tax) / (1 + r) ** 0.5
    return t.synergies * (1 - t.tax) * annuity - shortfall - t.synergies * t.integration_x


@dataclass
class DealResult:
    offer_cad: float
    exchange_ratio: float
    new_shares: float
    equity_usd: float
    cash_usd: float
    stock_usd: float
    net_debt_assumed: float
    ev_usd: float
    pnav: float
    ev_ebitda: float
    cash_need: float
    cash_used: float
    new_debt: float
    pf_shares: float
    own_cg: float
    syn_npv: float
    pf_nav: float
    pf_navps: float
    acc_navps: float
    acc_navps_ex_syn: float
    pf_net_debt: float
    leverage: float
    years: dict[int, dict[str, float]]


def merge(cg: CompanyResult, ero: CompanyResult, t: DealTerms, cadusd: float) -> DealResult:
    offer = ero.price_cad * (1 + t.premium)
    er = offer * (1 - t.cash_pct) / cg.price_cad
    new_shares = ero.shares * er
    equity = offer * ero.shares * cadusd
    cash_usd = equity * t.cash_pct
    net_debt = -(ero.debt + ero.cash)
    need = cash_usd + t.txn_costs
    used = min(need, max(0.0, cg.cash + ero.cash - t.min_cash))
    new_debt = need - used
    pf_shares = cg.shares + new_shares
    syn_npv = synergy_npv(t)
    pf_nav = cg.nav + ero.nav + syn_npv - need
    pf_navps = pf_nav / pf_shares
    years = {}
    for year, phase in zip(FIN_YEARS, (t.phase_y1, t.phase_y2)):
        syn = t.synergies * phase
        fin_cost = -new_debt * t.debt_rate * (1 - t.tax) - used * t.cash_yield * (1 - t.tax)
        ni = cg.fin[year]["ni"] + ero.fin[year]["ni"] + syn * (1 - t.tax) + fin_cost
        cfo = cg.fin[year]["cfo"] + ero.fin[year]["cfo"] + syn * (1 - t.tax) + fin_cost
        years[year] = {
            "syn": syn, "pf_ni": ni, "pf_eps": ni / pf_shares, "acc_eps": ni / pf_shares / cg.fin[year]["eps"] - 1,
            "pf_cfo": cfo, "pf_cfps": cfo / pf_shares, "acc_cfps": cfo / pf_shares / cg.fin[year]["cfps"] - 1,
            "pf_ebitda": cg.fin[year]["ebitda"] + ero.fin[year]["ebitda"] + syn,
        }
    pf_net_debt = -(cg.cash + ero.cash + cg.debt + ero.debt) + need
    return DealResult(
        offer_cad=offer, exchange_ratio=er, new_shares=new_shares, equity_usd=equity, cash_usd=cash_usd,
        stock_usd=equity - cash_usd, net_debt_assumed=net_debt, ev_usd=equity + net_debt, pnav=offer / ero.navps_cad,
        ev_ebitda=(equity + net_debt) / ero.fin[2027]["ebitda"], cash_need=need, cash_used=used, new_debt=new_debt,
        pf_shares=pf_shares, own_cg=cg.shares / pf_shares, syn_npv=syn_npv, pf_nav=pf_nav, pf_navps=pf_navps,
        acc_navps=pf_navps / cg.navps_usd - 1,
        acc_navps_ex_syn=(cg.nav + ero.nav - need) / pf_shares / cg.navps_usd - 1,
        pf_net_debt=pf_net_debt, leverage=pf_net_debt / years[2027]["pf_ebitda"], years=years,
    )


# --- Whole-model entry point ------------------------------------------------------------

@dataclass
class ModelResult:
    prices: Prices
    terms: DealTerms
    companies: dict[str, CompanyResult]
    deal: DealResult


def run(a: Assumptions | None = None, prices: Prices | None = None, terms: DealTerms | None = None) -> ModelResult:
    a = a or Assumptions()
    prices = prices or Prices.deck(a)
    terms = terms or DealTerms.from_assumptions(a)
    companies = {co: value_company(a, co, prices) for co in COMPANIES}
    return ModelResult(prices, terms, companies, merge(companies["CG"], companies["ERO"], terms, prices.cadusd))


# --- Sensitivities and football field ---------------------------------------------------

def navps_sensitivity(a: Assumptions, premiums, coppers, terms: DealTerms | None = None,
                      base: Prices | None = None) -> pd.DataFrame:
    """NAVPS accretion by premium (rows) and flat copper price (columns); gold stays on `base`."""
    base = base or Prices.deck(a)
    terms = terms or DealTerms.from_assumptions(a)
    grid = pd.DataFrame(index=list(premiums), columns=list(coppers), dtype=float)
    for cu in coppers:
        prices = base.with_copper(cu)
        cg, ero = value_company(a, "CG", prices), value_company(a, "ERO", prices)
        for prem in premiums:
            grid.loc[prem, cu] = merge(cg, ero, replace(terms, premium=prem), prices.cadusd).acc_navps
    return grid


def eps_sensitivity(model: ModelResult, synergies, cash_pcts, year: int = 2028) -> pd.DataFrame:
    cg, ero = model.companies["CG"], model.companies["ERO"]
    grid = pd.DataFrame(index=list(synergies), columns=list(cash_pcts), dtype=float)
    for syn in synergies:
        for c in cash_pcts:
            terms = replace(model.terms, synergies=syn, cash_pct=c)
            grid.loc[syn, c] = merge(cg, ero, terms, model.prices.cadusd).years[year]["acc_eps"]
    return grid


def football_field(a: Assumptions, model: ModelResult, spot: ModelResult | None = None) -> pd.DataFrame:
    """Low/high value per Ero share (C$) by method, mirroring the Football_Field tab."""
    spot = spot or run(a, Prices.spot(a), model.terms)
    fx = model.prices.cadusd
    mkt = a.market.loc["ERO.TO"]
    ero = model.companies["ERO"]
    stats = peer_stats(comps_table(a))
    net_debt = model.deal.net_debt_assumed
    per_share = lambda ev: (ev - net_debt) / ero.shares / fx
    reserves_lb = reserves_cueq_mt(a, "ERO") * LB_PER_TONNE
    prem = a._csv("precedents.csv")["premium_spot"].dropna()
    rel = [model.companies["ERO"].navps_cad * model.companies["CG"].pnav,
           spot.companies["ERO"].navps_cad * spot.companies["CG"].pnav]
    nav = [model.companies["ERO"].navps_cad, spot.companies["ERO"].navps_cad]
    rows = {
        "WK52": (mkt["week52_low"], mkt["week52_high"]),
        "TARGETS": (mkt["target_low"], mkt["target_high"]),
        "EVEBITDA": tuple(per_share(stats.loc[q, "ev_ebitda"] * mkt["ebitda_ttm_usd_m"]) for q in ("q1", "q3")),
        "PCF": tuple(stats.loc[q, "p_cf"] * mkt["cfo_ttm_usd_m"] / ero.shares / fx for q in ("q1", "q3")),
        "EVLB": tuple(per_share(stats.loc[q, "ev_per_lb_cueq"] * reserves_lb) for q in ("q1", "q3")),
        "PREC": (ero.price_cad * (1 + prem.quantile(0.25)), ero.price_cad * (1 + prem.quantile(0.75))),
        "NAVREL": (min(rel), max(rel)),
        "NAV": (min(nav), max(nav)),
    }
    return pd.DataFrame(rows, index=["low", "high"]).T.astype(float)
