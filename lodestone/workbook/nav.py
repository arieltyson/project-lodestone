"""Inputs, Market, mine-by-mine NAV tabs and the corporate NAV roll-up."""

from __future__ import annotations

from lodestone.assumptions import COMPANIES
from lodestone.workbook import styles as s
from lodestone.workbook.context import FIRST_YEAR_COL, ModelContext

LB_PER_TONNE = 2204.62
GRAMS_PER_OZ = 31.1035

# --- Inputs tab -----------------------------------------------------------------------

INPUTS = "Inputs"
ROW_CU_SPOT, ROW_AU_SPOT, ROW_FX_SPOT = 5, 6, 7
ROW_CU_CASE, ROW_AU_CASE, ROW_CU_CUSTOM, ROW_AU_CUSTOM = 10, 11, 12, 13
ROW_DECK_HEADER, ROW_CU_DECK, ROW_AU_DECK, ROW_CU_ACTIVE, ROW_AU_ACTIVE = 15, 16, 17, 18, 19
ROW_DISC_HEADER = 21
ROW_GA_YEARS, ROW_GA_RATE = 30, 31


def build_inputs(ctx: ModelContext) -> None:
    ws = ctx.wb.create_sheet(INPUTS)
    a = ctx.a
    s.title(ws, "Inputs", "Price deck, FX, price-case selectors and discount rates. Edit blue cells "
            "(or the CSVs under data/ and rebuild).")
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 60

    s.section(ws, 4, f"Spot market (yfinance snapshot {a.macro['timestamp_utc']})", 5)
    for row, label, key, unit, fmt in (
        (ROW_CU_SPOT, "Copper spot (COMEX HG)", "copper_usd_lb", "US$/lb", s.USD_2),
        (ROW_AU_SPOT, "Gold spot (COMEX GC)", "gold_usd_oz", "US$/oz", s.USD_0),
        (ROW_FX_SPOT, "CAD/USD spot", "cadusd", "US$ per C$", s.NUM_3),
    ):
        s.put(ws, row, 2, label)
        s.put(ws, row, 3, round(float(a.macro[key]), 4), s.INPUT, fmt)
        s.put(ws, row, 4, unit, s.NOTE)
        s.put(ws, row, 5, "Source: MKT-YF (data/market/latest_macro.csv)", s.NOTE)

    s.section(ws, 9, "Price case selectors (1 = analyst deck, 2 = spot held flat, 3 = custom flat)", 5)
    for row, label, value, fmt in (
        (ROW_CU_CASE, "Copper price case", 1, "0"),
        (ROW_AU_CASE, "Gold price case", 1, "0"),
        (ROW_CU_CUSTOM, "Custom copper price (US$/lb)", 4.50, s.USD_2),
        (ROW_AU_CUSTOM, "Custom gold price (US$/oz)", 3000, s.USD_0),
    ):
        s.put(ws, row, 2, label)
        cell = s.put(ws, row, 3, value, s.INPUT, fmt)
        cell.fill = s.KEY_FILL

    s.header_row(ws, ROW_DECK_HEADER, 2, ["Price deck", "", "", "Source"])
    s.header_row(ws, ROW_DECK_HEADER, FIRST_YEAR_COL, a.years)
    labels = {
        ROW_CU_DECK: ("Copper deck (US$/lb)", "data/price_deck.csv; last deck year held flat"),
        ROW_AU_DECK: ("Gold deck (US$/oz)", "data/price_deck.csv; last deck year held flat"),
        ROW_CU_ACTIVE: ("Copper price used in model (US$/lb)", "CHOOSE(case, deck, spot, custom)"),
        ROW_AU_ACTIVE: ("Gold price used in model (US$/oz)", "CHOOSE(case, deck, spot, custom)"),
    }
    for row, (label, note) in labels.items():
        s.put(ws, row, 2, label, s.BOLD if row >= ROW_CU_ACTIVE else s.FORMULA)
        s.put(ws, row, 5, note, s.NOTE)
    for year in a.years:
        c = ctx.year_col(year)
        L = s.col(c)
        s.put(ws, ROW_CU_DECK, c, float(a.price_deck.loc[year, "copper_usd_lb"]), s.INPUT, s.USD_2)
        s.put(ws, ROW_AU_DECK, c, float(a.price_deck.loc[year, "gold_usd_oz"]), s.INPUT, s.USD_0)
        s.put(ws, ROW_CU_ACTIVE, c,
              f"=CHOOSE($C${ROW_CU_CASE},{L}{ROW_CU_DECK},$C${ROW_CU_SPOT},$C${ROW_CU_CUSTOM})",
              fmt=s.USD_2)
        s.put(ws, ROW_AU_ACTIVE, c,
              f"=CHOOSE($C${ROW_AU_CASE},{L}{ROW_AU_DECK},$C${ROW_AU_SPOT},$C${ROW_AU_CUSTOM})",
              fmt=s.USD_0)
        ws.column_dimensions[L].width = 9

    s.header_row(ws, ROW_DISC_HEADER, 2, ["Discount rate (real)", "Rate", "", "Rationale"])
    for i, asset in enumerate(a.assets.index):
        row = ROW_DISC_HEADER + 1 + i
        s.put(ws, row, 2, a.assets.loc[asset, "display_name"])
        cell = s.put(ws, row, 3, float(a.discount_rates[asset]), s.INPUT, s.PCT)
        cell.fill = s.KEY_FILL
        s.put(ws, row, 5, a.discount_rationale[asset], s.NOTE)
        ctx.refs[f"disc.{asset}"] = s.absref(INPUTS, 3, row)

    s.section(ws, ROW_GA_YEARS - 1, "Corporate", 5)
    s.put(ws, ROW_GA_YEARS, 2, "G&A capitalisation (years)")
    s.put(ws, ROW_GA_YEARS, 3, int(a.settings["ga_capitalisation_years"]), s.INPUT, "0")
    s.put(ws, ROW_GA_RATE, 2, "G&A discount rate")
    s.put(ws, ROW_GA_RATE, 3, float(a.settings["ga_discount_rate"]), s.INPUT, s.PCT)
    ws.freeze_panes = "C4"

    ctx.refs["fx_spot"] = s.absref(INPUTS, 3, ROW_FX_SPOT)
    ctx.refs["cu_case"] = s.absref(INPUTS, 3, ROW_CU_CASE)
    ctx.refs["au_case"] = s.absref(INPUTS, 3, ROW_AU_CASE)
    ctx.refs["cu_custom"] = s.absref(INPUTS, 3, ROW_CU_CUSTOM)
    ctx.refs["au_custom"] = s.absref(INPUTS, 3, ROW_AU_CUSTOM)
    ctx.refs["ga_years"] = s.absref(INPUTS, 3, ROW_GA_YEARS)
    ctx.refs["ga_rate"] = s.absref(INPUTS, 3, ROW_GA_RATE)


# --- Market tab -----------------------------------------------------------------------

MARKET = "Market"
MARKET_FIELDS = (
    ("name", "Company", None),
    ("price_local", "Price (C$)", s.USD_2),
    ("shares_m", "Shares (M)", s.USD_M),
    ("market_cap_usd_m", "Mkt cap (US$M)", s.USD_0),
    ("cash_usd_m", "Cash (US$M)", s.USD_0),
    ("debt_usd_m", "Debt (US$M)", s.USD_0),
    ("ev_usd_m", "EV (US$M)", s.USD_0),
    ("ebitda_ttm_usd_m", "EBITDA TTM (US$M)", s.USD_0),
    ("cfo_ttm_usd_m", "CFO TTM (US$M)", s.USD_0),
    ("eps_ttm_fin", "EPS TTM (US$)", s.USD_2),
    ("week52_low", "52w low (C$)", s.USD_2),
    ("week52_high", "52w high (C$)", s.USD_2),
    ("target_low", "Target low (C$)", s.USD_2),
    ("target_mean", "Target mean (C$)", s.USD_2),
    ("target_high", "Target high (C$)", s.USD_2),
)


def build_market(ctx: ModelContext) -> None:
    ws = ctx.wb.create_sheet(MARKET)
    a = ctx.a
    s.title(ws, "Market data", f"yfinance snapshot {a.macro['timestamp_utc']} (scripts/fetch_market_data.py). "
            "Balance sheet and TTM figures converted to US$M.")
    universe = a._csv("universe.csv").set_index("ticker")
    market = a.market.copy()
    market["shares_m"] = market["shares_out"] / 1e6
    market["name"] = universe["name"].reindex(market.index)

    s.header_row(ws, 4, 2, ["Ticker"] + [label for _, label, _ in MARKET_FIELDS])
    ws.column_dimensions["B"].width = 10
    for i, ticker in enumerate(market.index):
        row = 5 + i
        s.put(ws, row, 2, ticker, s.BOLD)
        for j, (field, _, fmt) in enumerate(MARKET_FIELDS):
            value = market.loc[ticker, field]
            if hasattr(value, "item"):
                value = value.item()
            s.put(ws, row, 3 + j, value, s.INPUT, fmt)
            ctx.refs[f"mkt.{ticker}.{field}"] = s.absref(MARKET, 3 + j, row)
        ctx.refs[f"mkt_row.{ticker}"] = str(row)
    ctx.refs["mkt.first_row"] = "5"
    ctx.refs["mkt.last_row"] = str(4 + len(market))
    ws.column_dimensions["C"].width = 18
    for j in range(len(MARKET_FIELDS)):
        ws.column_dimensions[s.col(4 + j)].width = 13
    s.put(ws, 6 + len(market), 2, "Source: MKT-YF. Values are hardcoded inputs refreshed by the pipeline.", s.NOTE)


# --- Asset tabs -----------------------------------------------------------------------

ASSET_INPUTS = (
    ("ownership", "Ownership", "%", s.PCT),
    ("reserve_mt", "P&P reserve (at effective date)", "Mt", s.NUM_3),
    ("depleted_mt", "Depletion to 1 Jan 2027", "Mt", s.NUM_3),
    ("cu_grade_pct", "Copper grade", "% Cu", s.PCT_2),
    ("au_gpt", "Gold grade", "g/t Au", s.USD_2),
    ("cu_recovery", "Copper recovery", "%", s.PCT),
    ("au_recovery", "Gold recovery", "%", s.PCT),
    ("cu_payable", "Copper payable", "%", s.PCT),
    ("au_payable", "Gold payable", "%", s.PCT),
    ("throughput_mtpa", "Throughput", "Mtpa", s.USD_2),
    ("expansion_mtpa", "Throughput after expansion", "Mtpa", s.USD_2),
    ("expansion_year", "Expansion year", "year", s.YEAR),
    ("start_year", "First production year", "year", s.YEAR),
    ("opex_usd_t", "Site operating cost", "US$/t milled", s.USD_2),
    ("tcrc_usd_lb", "TC/RC and freight", "US$/lb Cu", s.USD_2),
    ("sustaining_usd_m", "Sustaining capex", "US$M/yr", s.USD_M),
    ("growth_capex_usd_m", "Growth capex (total)", "US$M", s.USD_M),
    ("growth_start", "Growth capex start", "year", s.YEAR),
    ("growth_years", "Growth capex years", "years", "0"),
    ("closure_usd_m", "Closure cost", "US$M", s.USD_M),
    ("nsr_royalty", "Royalties (NSR)", "% revenue", s.PCT_2),
    ("tax_rate", "Cash tax rate", "%", s.PCT),
    ("stream_au_pct", "Gold streamed", "% payable Au", s.PCT_2),
    ("stream_au_price", "Stream gold price received", "US$/oz", s.USD_0),
    ("stream_cu_pct", "Copper streamed", "% payable Cu", s.PCT_2),
    ("stream_cu_spot_pct", "Stream copper price received", "% of spot", s.PCT),
)
INPUT_ROW0 = 5
ROW_DISC = INPUT_ROW0 + len(ASSET_INPUTS)

# Time-series rows (row number, key, label, unit, number format, LOM total in column D?)
TS_HEADER = ROW_DISC + 3
TS = {}
_ts_rows = (
    ("t", "Period", "", "0", False),
    ("cu_price", "Copper price", "US$/lb", s.USD_2, False),
    ("au_price", "Gold price", "US$/oz", s.USD_0, False),
    ("capacity", "Mill capacity", "Mt", s.USD_2, False),
    ("open_res", "Opening reserve", "Mt", s.USD_M, False),
    ("ore", "Ore processed", "Mt", s.USD_2, True),
    ("close_res", "Closing reserve", "Mt", s.USD_M, False),
    ("cu_lb", "Payable copper", "Mlb", s.USD_M, True),
    ("au_oz", "Payable gold", "koz", s.USD_M, True),
    ("cu_real", "Realised copper price (after stream)", "US$/lb", s.USD_2, False),
    ("au_real", "Realised gold price (after stream)", "US$/oz", s.USD_0, False),
    ("cu_rev", "Copper revenue", "US$M", s.USD_M, True),
    ("au_rev", "Gold revenue", "US$M", s.USD_M, True),
    ("rev", "Total revenue", "US$M", s.USD_M, True),
    ("tcrc", "TC/RC and freight", "US$M", s.USD_M, True),
    ("royalty", "Royalties", "US$M", s.USD_M, True),
    ("opex", "Site operating costs", "US$M", s.USD_M, True),
    ("ebitda", "Mine EBITDA", "US$M", s.USD_M, True),
    ("sust", "Sustaining capex", "US$M", s.USD_M, True),
    ("growth", "Growth capex", "US$M", s.USD_M, True),
    ("closure", "Closure and reclamation", "US$M", s.USD_M, True),
    ("tax", "Cash taxes", "US$M", s.USD_M, True),
    ("fcf", "Unlevered free cash flow", "US$M", s.USD_M, True),
    ("df", "Discount factor (mid-year)", "", "0.0000", False),
    ("pv", "PV of free cash flow", "US$M", s.USD_M, True),
)
for _i, (_key, *_rest) in enumerate(_ts_rows):
    TS[_key] = TS_HEADER + 1 + _i
SUMMARY0 = TS_HEADER + len(_ts_rows) + 3
SUMMARY = {
    "npv": SUMMARY0,
    "nav": SUMMARY0 + 1,
    "life": SUMMARY0 + 2,
    "aisc_y1": SUMMARY0 + 3,
    "aisc_lom": SUMMARY0 + 4,
}


def _asset_formula(key: str, L: str, P: str | None, inp: dict[str, str], first: bool) -> str:
    r = TS
    if key == "t":
        return f"={L}${TS_HEADER}-{inp['first_year']}+1"
    if key == "capacity":
        return (f"=IF({L}${TS_HEADER}<{inp['start_year']},0,IF({L}${TS_HEADER}>={inp['expansion_year']},"
                f"{inp['expansion_mtpa']},{inp['throughput_mtpa']}))")
    if key == "open_res":
        return f"={inp['reserve_mt']}-{inp['depleted_mt']}" if first else f"={P}{r['close_res']}"
    if key == "ore":
        return f"=MAX(0,MIN({L}{r['capacity']},{L}{r['open_res']}))"
    if key == "close_res":
        return f"={L}{r['open_res']}-{L}{r['ore']}"
    if key == "cu_lb":
        return (f"={L}{r['ore']}*{inp['cu_grade_pct']}*{inp['cu_recovery']}*{inp['cu_payable']}"
                f"*{LB_PER_TONNE}")
    if key == "au_oz":
        return (f"={L}{r['ore']}*{inp['au_gpt']}*{inp['au_recovery']}*{inp['au_payable']}"
                f"*1000/{GRAMS_PER_OZ}")
    if key == "cu_real":
        return (f"={L}{r['cu_price']}*(1-{inp['stream_cu_pct']})"
                f"+{L}{r['cu_price']}*{inp['stream_cu_pct']}*{inp['stream_cu_spot_pct']}")
    if key == "au_real":
        return (f"={L}{r['au_price']}*(1-{inp['stream_au_pct']})"
                f"+{inp['stream_au_pct']}*{inp['stream_au_price']}")
    if key == "cu_rev":
        return f"={L}{r['cu_lb']}*{L}{r['cu_real']}"
    if key == "au_rev":
        return f"={L}{r['au_oz']}*{L}{r['au_real']}/1000"
    if key == "rev":
        return f"={L}{r['cu_rev']}+{L}{r['au_rev']}"
    if key == "tcrc":
        return f"=-{L}{r['cu_lb']}*{inp['tcrc_usd_lb']}"
    if key == "royalty":
        return f"=-{L}{r['rev']}*{inp['nsr_royalty']}"
    if key == "opex":
        return f"=-{L}{r['ore']}*{inp['opex_usd_t']}"
    if key == "ebitda":
        return f"=SUM({L}{r['rev']}:{L}{r['opex']})"
    if key == "sust":
        return f"=-IF({L}{r['ore']}>0,{inp['sustaining_usd_m']},0)"
    if key == "growth":
        return (f"=-IF(AND({L}${TS_HEADER}>={inp['growth_start']},{L}${TS_HEADER}<{inp['growth_start']}"
                f"+{inp['growth_years']}),{inp['growth_capex_usd_m']}/{inp['growth_years']},0)")
    if key == "closure":
        return "=0" if first else f"=-IF(AND({L}{r['ore']}=0,{P}{r['ore']}>0),{inp['closure_usd_m']},0)"
    if key == "tax":
        return f"=-MAX(0,({L}{r['ebitda']}+{L}{r['sust']}+{L}{r['growth']})*{inp['tax_rate']})"
    if key == "fcf":
        return f"=SUM({L}{r['ebitda']}:{L}{r['tax']})"
    if key == "df":
        return f"=1/(1+{inp['disc']})^({L}{r['t']}-0.5)"
    if key == "pv":
        return f"={L}{r['fcf']}*{L}{r['df']}"
    raise KeyError(key)


def build_asset(ctx: ModelContext, asset: str) -> None:
    a = ctx.a
    row = a.assets.loc[asset]
    ws = ctx.wb.create_sheet(asset)
    gold = row["primary_metal"].startswith("gold")
    s.title(ws, f"{row['display_name']} — asset NAV ({row['company']}, {row['country']})",
            "Reserve-constrained life-of-mine model. Blue = input (data/assets.csv), "
            "green = link, black = formula. US$M real unless noted.")
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 3

    notes = a.asset_notes[a.asset_notes["asset"] == asset]
    s.header_row(ws, INPUT_ROW0 - 1, 2, ["Input", "Value", "Unit"])
    inp: dict[str, str] = {"first_year": str(a.first_year)}
    for i, (field, label, unit, fmt) in enumerate(ASSET_INPUTS):
        r = INPUT_ROW0 + i
        s.put(ws, r, 2, label)
        value = row[field]
        value = value.item() if hasattr(value, "item") else value
        s.put(ws, r, 3, value, s.INPUT, fmt)
        s.put(ws, r, 4, unit, s.NOTE)
        inp[field] = f"$C${r}"
    s.put(ws, ROW_DISC, 2, "Discount rate (real)", s.BOLD)
    s.put(ws, ROW_DISC, 3, f"={ctx.refs[f'disc.{asset}']}", s.LINK, s.PCT)
    s.put(ws, ROW_DISC, 4, "Inputs tab", s.NOTE)
    inp["disc"] = f"$C${ROW_DISC}"

    # Source notes beside the input block.
    s.put(ws, INPUT_ROW0 - 1, 6, "Source / status", s.BOLD)
    s.put(ws, INPUT_ROW0, 6, f"Source IDs: {row['source_ids']} (see Sources tab)", s.NOTE)
    for j, note in enumerate(notes.itertuples()):
        s.put(ws, INPUT_ROW0 + 1 + j, 6, f"[{note.status}] {note.field}: {note.note}", s.NOTE)

    s.header_row(ws, TS_HEADER, 2, ["US$M unless noted", "Unit", "LOM / NPV"])
    s.header_row(ws, TS_HEADER, FIRST_YEAR_COL, a.years)
    for key, label, unit, fmt, total in _ts_rows:
        r = TS[key]
        bold = key in ("rev", "ebitda", "fcf", "pv")
        s.put(ws, r, 2, label, s.BOLD if bold else s.FORMULA)
        s.put(ws, r, 3, unit, s.NOTE)
        for year in a.years:
            c = ctx.year_col(year)
            L, P = s.col(c), s.col(c - 1)
            first = year == a.first_year
            if key == "cu_price":
                formula, font = f"={INPUTS}!{L}${ROW_CU_ACTIVE}", s.LINK
            elif key == "au_price":
                formula, font = f"={INPUTS}!{L}${ROW_AU_ACTIVE}", s.LINK
            else:
                formula, font = _asset_formula(key, L, P, inp, first), s.FORMULA
            cell = s.put(ws, r, c, formula, font, fmt)
            if bold:
                cell.border = s.TOTAL_BORDER
                cell.font = s.BOLD
        if total:
            first_l, last_l = s.col(FIRST_YEAR_COL), s.col(ctx.last_col)
            s.put(ws, r, 4, f"=SUM({first_l}{r}:{last_l}{r})", s.BOLD if bold else s.FORMULA, fmt)
        ws.column_dimensions[s.col(FIRST_YEAR_COL)].width = 9
    for year in a.years:
        ws.column_dimensions[s.col(ctx.year_col(year))].width = 9

    F, LAST = s.col(FIRST_YEAR_COL), s.col(ctx.last_col)
    s.section(ws, SUMMARY0 - 1, "Asset summary", 4)
    primary = "oz Au" if gold else "lb Cu"
    if gold:
        aisc = lambda c: (f"=IFERROR((-({c}{TS['tcrc']}+{c}{TS['royalty']}+{c}{TS['opex']}+{c}{TS['sust']})"
                          f"-{c}{TS['cu_rev']})/{c}{TS['au_oz']}*1000,0)")
    else:
        aisc = lambda c: (f"=IFERROR((-({c}{TS['tcrc']}+{c}{TS['royalty']}+{c}{TS['opex']}+{c}{TS['sust']})"
                          f"-{c}{TS['au_rev']})/{c}{TS['cu_lb']},0)")
    rows = (
        ("npv", "NPV (100% basis)", f"=SUM({F}{TS['pv']}:{LAST}{TS['pv']})", s.USD_M, "US$M"),
        ("nav", "Attributable NAV", f"=C{SUMMARY0}*{inp['ownership']}", s.USD_M, "US$M"),
        ("life", "Remaining mine life", f'=COUNTIF({F}{TS["ore"]}:{LAST}{TS["ore"]},">0")', "0", "years"),
        ("aisc_y1", f"AISC {a.first_year}, by-product", aisc(F), s.USD_2 if not gold else s.USD_0,
         f"US$/{primary}"),
        ("aisc_lom", "AISC life of mine, by-product", aisc("D"), s.USD_2 if not gold else s.USD_0,
         f"US$/{primary}"),
    )
    for key, label, formula, fmt, unit in rows:
        r = SUMMARY[key]
        s.put(ws, r, 2, label, s.BOLD)
        cell = s.put(ws, r, 3, formula, s.BOLD, fmt)
        if key == "nav":
            cell.fill = s.KEY_FILL
        s.put(ws, r, 4, unit, s.NOTE)
        ctx.refs[f"{asset}.{key}"] = s.absref(asset, 3, r)
        s.define(ctx.wb, f"{asset}_{key.upper()}", asset, 3, r)
    ctx.refs[f"{asset}.ownership"] = s.absref(asset, 3, INPUT_ROW0)


# --- Corporate NAV --------------------------------------------------------------------

CORP = "Corporate_NAV"
NAV_ROWS = {}
FIN_YEARS = (2027, 2028)


def _asset_sum(ctx: ModelContext, company: str, ts_key: str, year: int) -> str:
    L = s.col(ctx.year_col(year))
    terms = [f"'{asset}'!{L}{TS[ts_key]}*'{asset}'!$C${INPUT_ROW0}" for asset in ctx.a.assets_of(company)]
    return "=" + "+".join(terms)


def build_corporate_nav(ctx: ModelContext) -> None:
    a = ctx.a
    ws = ctx.wb.create_sheet(CORP)
    s.title(ws, "Corporate NAV and financial summary",
            "Sum of attributable asset NAVs plus corporate items. P/NAV uses the spot share price and "
            "spot CAD/USD.")
    ws.column_dimensions["B"].width = 40
    for c in "CDEF":
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["G"].width = 60
    cols = {"CG": 3, "ERO": 4}
    s.header_row(ws, 4, 2, ["US$M unless noted", "Centerra (CG)", "Ero (ERO)", "", "", "Note"])

    r = 5
    for asset in a.assets.index:
        s.put(ws, r, 2, f"{a.assets.loc[asset, 'display_name']} NAV")
        for company, c in cols.items():
            value = f"={ctx.refs[f'{asset}.nav']}" if a.assets.loc[asset, "company"] == company else 0
            s.put(ws, r, c, value, s.LINK if isinstance(value, str) else s.FORMULA, s.USD_M)
        r += 1
    asset_rows = (5, r - 1)

    def line(key, label, values, fmt=s.USD_M, bold=False, note=""):
        nonlocal r
        s.put(ws, r, 2, label, s.BOLD if bold else s.FORMULA)
        for company, c in cols.items():
            value, font = values(company)
            cell = s.put(ws, r, c, value, s.BOLD if bold and font is s.FORMULA else font, fmt)
            if bold:
                cell.border = s.TOTAL_BORDER
        if note:
            s.put(ws, r, 7, note, s.NOTE)
        NAV_ROWS[key] = r
        r += 1

    ticker = lambda company: a.corporate.loc[company, "ticker"]
    corp = a.corporate
    line("asset_nav", "Total attributable asset NAV",
         lambda co: (f"=SUM({s.col(cols[co])}{asset_rows[0]}:{s.col(cols[co])}{asset_rows[1]})", s.FORMULA),
         bold=True)
    line("ga", "Corporate G&A (capitalised)",
         lambda co: (f"=-{corp.loc[co, 'ga_usd_m']}*(1-(1+{ctx.refs['ga_rate']})^-{ctx.refs['ga_years']})"
                     f"/{ctx.refs['ga_rate']}", s.FORMULA),
         note="G&A (US$M/yr, data/corporate.csv) x annuity factor")
    line("other", "Other assets and exploration",
         lambda co: (float(corp.loc[co, "other_assets_usd_m"]), s.INPUT),
         note="CG: " + corp.loc["CG", "other_assets_note"] + "; ERO: " + corp.loc["ERO", "other_assets_note"])
    line("hedges", "Hedge book mark-to-market",
         lambda co: (float(corp.loc[co, "hedge_mtm_usd_m"]), s.INPUT), note="Neither company has a material hedge book")
    line("cash", "Cash and equivalents",
         lambda co: (f"={ctx.refs[f'mkt.{ticker(co)}.cash_usd_m']}", s.LINK), note="Market tab")
    line("debt", "Debt",
         lambda co: (f"=-{ctx.refs[f'mkt.{ticker(co)}.debt_usd_m']}", s.LINK), note="Market tab")
    line("nav", "Net asset value",
         lambda co: (f"=SUM({s.col(cols[co])}{NAV_ROWS['asset_nav']}:{s.col(cols[co])}{NAV_ROWS['debt']})",
                     s.FORMULA), bold=True)
    line("shares", "Shares outstanding (M)",
         lambda co: (f"={ctx.refs[f'mkt.{ticker(co)}.shares_m']}", s.LINK), note="Basic shares; Market tab")
    line("navps_usd", "NAV per share (US$)",
         lambda co: (f"={s.col(cols[co])}{NAV_ROWS['nav']}/{s.col(cols[co])}{NAV_ROWS['shares']}", s.FORMULA),
         fmt=s.USD_2, bold=True)
    line("navps_cad", "NAV per share (C$)",
         lambda co: (f"={s.col(cols[co])}{NAV_ROWS['navps_usd']}/{ctx.refs['fx_spot']}", s.FORMULA), fmt=s.USD_2)
    line("price_cad", "Share price (C$)",
         lambda co: (f"={ctx.refs[f'mkt.{ticker(co)}.price_local']}", s.LINK), fmt=s.USD_2)
    line("pnav", "P/NAV",
         lambda co: (f"={s.col(cols[co])}{NAV_ROWS['price_cad']}/{s.col(cols[co])}{NAV_ROWS['navps_cad']}",
                     s.FORMULA), fmt=s.MULT, bold=True)
    for company, c in cols.items():
        for key in ("nav", "navps_usd", "navps_cad", "pnav", "shares", "cash", "debt", "asset_nav", "ga", "other",
                    "hedges", "price_cad"):
            ctx.refs[f"{company}.{key}"] = s.absref(CORP, c, NAV_ROWS[key])
            s.define(ctx.wb, f"{company}_{key.upper()}", CORP, c, NAV_ROWS[key])

    # Financial summary: two forecast years per company.
    r += 1
    s.section(ws, r, "Financial summary (feeds the accretion/dilution analysis)", 7)
    r += 1
    fin_cols = {}
    headers = []
    for i, (company, year) in enumerate((co, y) for co in COMPANIES for y in FIN_YEARS):
        fin_cols[(company, year)] = 3 + i
        headers.append(f"{company} {year}E")
    s.header_row(ws, r, 2, ["US$M unless noted"] + headers)
    r += 1

    FIN = {}

    def fin(key, label, build, fmt=s.USD_M, bold=False, note=""):
        nonlocal r
        s.put(ws, r, 2, label, s.BOLD if bold else s.FORMULA)
        for (company, year), c in fin_cols.items():
            value, font = build(company, year, s.col(c))
            cell = s.put(ws, r, c, value, s.BOLD if bold and font is s.FORMULA else font, fmt)
            if bold:
                cell.border = s.TOTAL_BORDER
        if note:
            s.put(ws, r, 7, note, s.NOTE)
        FIN[key] = r
        r += 1

    shares_ref = lambda co: f"{s.col(cols[co])}${NAV_ROWS['shares']}"
    debt_ref = lambda co: f"{s.col(cols[co])}${NAV_ROWS['debt']}"
    fin("cu_lb", "Payable copper (Mlb, attributable)", lambda co, y, L: (_asset_sum(ctx, co, "cu_lb", y), s.LINK))
    fin("au_oz", "Payable gold (koz, attributable)", lambda co, y, L: (_asset_sum(ctx, co, "au_oz", y), s.LINK))
    fin("rev", "Revenue", lambda co, y, L: (_asset_sum(ctx, co, "rev", y), s.LINK))
    fin("mine_ebitda", "Mine EBITDA", lambda co, y, L: (_asset_sum(ctx, co, "ebitda", y), s.LINK))
    fin("ga", "Corporate G&A", lambda co, y, L: (-float(corp.loc[co, "ga_usd_m"]), s.INPUT))
    fin("ebitda", "EBITDA", lambda co, y, L: (f"={L}{FIN['mine_ebitda']}+{L}{FIN['ga']}", s.FORMULA), bold=True)
    fin("dna", "Depreciation and amortisation",
        lambda co, y, L: (-float(corp.loc[co, f"dna_{y}_usd_m"]), s.INPUT), note="data/corporate.csv (estimate)")
    fin("interest", "Interest expense",
        lambda co, y, L: (f"={debt_ref(co)}*{corp.loc[co, 'interest_rate']}", s.FORMULA),
        note="Debt x coupon (data/corporate.csv)")
    fin("ebt", "Pre-tax income", lambda co, y, L: (f"={L}{FIN['ebitda']}+{L}{FIN['dna']}+{L}{FIN['interest']}",
                                                     s.FORMULA))
    fin("book_tax", "Income tax", lambda co, y, L: (f"=-MAX(0,{L}{FIN['ebt']})*{corp.loc[co, 'book_tax_rate']}",
                                                     s.FORMULA), note="Book tax at effective rate (data/corporate.csv)")
    fin("ni", "Net income", lambda co, y, L: (f"={L}{FIN['ebt']}+{L}{FIN['book_tax']}", s.FORMULA), bold=True)
    fin("eps", "EPS (US$)", lambda co, y, L: (f"={L}{FIN['ni']}/{shares_ref(co)}", s.FORMULA), fmt=s.USD_2,
        bold=True)
    fin("cash_tax", "Cash taxes (asset level)", lambda co, y, L: (_asset_sum(ctx, co, "tax", y), s.LINK))
    fin("cfo", "Operating cash flow", lambda co, y, L: (f"={L}{FIN['ebitda']}+{L}{FIN['interest']}"
                                                         f"+{L}{FIN['cash_tax']}", s.FORMULA), bold=True)
    fin("cfps", "CFPS (US$)", lambda co, y, L: (f"={L}{FIN['cfo']}/{shares_ref(co)}", s.FORMULA), fmt=s.USD_2,
        bold=True)
    for (company, year), c in fin_cols.items():
        for key in FIN:
            ctx.refs[f"{company}.{key}.{year}"] = s.absref(CORP, c, FIN[key])
        for key in ("ebitda", "ni", "eps", "cfo", "cfps"):
            s.define(ctx.wb, f"{company}_{key.upper()}_{year}", CORP, c, FIN[key])
    ws.freeze_panes = "C5"


def build(ctx: ModelContext) -> None:
    build_inputs(ctx)
    build_market(ctx)
    for asset in ctx.a.assets.index:
        build_asset(ctx, asset)
    build_corporate_nav(ctx)
