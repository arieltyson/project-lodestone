"""Behavioural checks on the valuation engine that do not depend on the workbook."""

from __future__ import annotations

from dataclasses import replace

import pytest

from lodestone.assumptions import Assumptions
from lodestone.valuation import DealTerms, Prices, asset_cashflows, merge, run, value_company


@pytest.fixture(scope="module")
def a():
    return Assumptions()


@pytest.mark.parametrize("asset", ["Oksut", "Tucuma", "Xavantina", "Caraiba", "Goldfield", "MtMilligan"])
def test_mine_plan_never_exceeds_reserves(a, asset):
    cf = asset_cashflows(a, asset, Prices.deck(a))
    p = a.assets.loc[asset]
    assert cf["ore"].sum() == pytest.approx(min(p["reserve_mt"] - p["depleted_mt"], cf["ore"].sum()))
    assert cf["ore"].sum() <= p["reserve_mt"] - p["depleted_mt"] + 1e-9


def test_reserves_exhausted_within_horizon(a):
    for asset, p in a.assets.iterrows():
        cf = asset_cashflows(a, asset, Prices.deck(a))
        assert cf["ore"].sum() == pytest.approx(p["reserve_mt"] - p["depleted_mt"]), asset
        assert cf["closure"].sum() == pytest.approx(-p["closure_usd_m"]), asset


def test_stream_lowers_mount_milligan_revenue(a):
    streamed = asset_cashflows(a, "MtMilligan", Prices.deck(a))
    a_free = Assumptions()
    a_free.assets.loc["MtMilligan", ["stream_au_pct", "stream_cu_pct"]] = 0.0
    unstreamed = asset_cashflows(a_free, "MtMilligan", Prices.deck(a_free))
    assert unstreamed["rev"].sum() > streamed["rev"].sum()


def test_nav_rises_with_copper(a):
    low = value_company(a, "ERO", Prices.flat(a, 4.0, 3000)).nav
    high = value_company(a, "ERO", Prices.flat(a, 6.0, 3000)).nav
    assert high > low


def test_all_stock_deal_uses_no_cash_for_consideration(a):
    model = run(a, terms=DealTerms.from_assumptions(a, cash_pct=0.0))
    assert model.deal.cash_usd == 0
    assert model.deal.cash_need == pytest.approx(model.terms.txn_costs)


def test_eps_accretion_increases_with_synergies(a):
    model = run(a)
    cg, ero = model.companies["CG"], model.companies["ERO"]
    low = merge(cg, ero, replace(model.terms, synergies=0), model.prices.cadusd).years[2028]["acc_eps"]
    high = merge(cg, ero, replace(model.terms, synergies=200), model.prices.cadusd).years[2028]["acc_eps"]
    assert high > low


def test_higher_premium_is_more_dilutive_to_navps(a):
    model = run(a)
    cg, ero = model.companies["CG"], model.companies["ERO"]
    acc = [merge(cg, ero, replace(model.terms, premium=p), model.prices.cadusd).acc_navps for p in (0.1, 0.3, 0.5)]
    assert acc[0] > acc[1] > acc[2]
