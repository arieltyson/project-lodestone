"""Synergies, pro forma merger, contribution analysis, scenarios and sensitivity tabs.

Formulas lean on workbook-level defined names (CG_NAV, DEAL_PREMIUM, ...) so a reviewer
can read `=ERO_PRICE_CAD*(1+DEAL_PREMIUM)` instead of chasing cell addresses.
"""

from __future__ import annotations

from lodestone.workbook import styles as s
from lodestone.workbook.context import ModelContext

FIN_YEARS = (2027, 2028)


def _name(ctx, ws, name: str, row: int, col: int = 3) -> None:
    s.define(ctx.wb, name, ws.title, col, row)


def _block(ctx, ws, row: int, rows, value_col: int = 3, note_col: int = 5) -> int:
    """Write (label, name, value, font, fmt, note) rows; returns the next free row."""
    for label, name, value, font, fmt, note in rows:
        s.put(ws, row, 2, label, s.BOLD if font is s.BOLD else s.FORMULA)
        s.put(ws, row, value_col, value, font, fmt)
        if note:
            s.put(ws, row, note_col, note, s.NOTE)
        if name:
            _name(ctx, ws, name, row, value_col)
        row += 1
    return row


# --- Synergies ------------------------------------------------------------------------

def build_synergies(ctx: ModelContext) -> None:
    syn = ctx.a._csv("synergies.csv")
    ws = ctx.wb.create_sheet("Synergies")
    s.title(ws, "Run-rate synergy case", "Pre-tax, US$M per year at run-rate. Basis for each line in column E.")
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 3
    ws.column_dimensions["E"].width = 110
    s.header_row(ws, 4, 2, ["Category", "Run-rate (US$M)", "", "Basis"])
    first = 5
    for i, line in syn.iterrows():
        r = first + i
        s.put(ws, r, 2, line["category"])
        cell = s.put(ws, r, 3, float(line["run_rate_usd_m"]), s.INPUT, s.USD_M)
        cell.fill = s.KEY_FILL
        s.put(ws, r, 5, line["basis"], s.NOTE)
    last = first + len(syn) - 1
    r = last + 1
    s.put(ws, r, 2, "Total run-rate synergies", s.BOLD)
    cell = s.put(ws, r, 3, f"=SUM(C{first}:C{last})", s.BOLD, s.USD_M)
    cell.border = s.TOTAL_BORDER
    _name(ctx, ws, "SYN_RUNRATE", r)

    r += 2
    s.header_row(ws, r, 2, ["Phase-in and value", "Value", "", "Note"])
    r += 1
    r = _block(ctx, ws, r, (
        ("Synergies 2027E (pre-tax)", "SYN_2027", "=SYN_RUNRATE*DEAL_PHASE_Y1", s.FORMULA, s.USD_M,
         "Run-rate x year-1 phase-in"),
        ("Synergies 2028E (pre-tax)", "SYN_2028", "=SYN_RUNRATE*DEAL_PHASE_Y2", s.FORMULA, s.USD_M,
         "Run-rate x year-2 phase-in"),
        ("One-off integration costs", "SYN_INTEGRATION", "=SYN_RUNRATE*DEAL_INTEGRATION_X", s.FORMULA, s.USD_M,
         "Cost to achieve; excluded from EPS and CFPS, deducted from NAV"),
        ("NPV of after-tax synergies, net of integration costs", "SYN_NPV",
         "=SYN_RUNRATE*(1-DEAL_TAX)*((1-(1+DEAL_SYN_RATE)^-DEAL_SYN_YEARS)/DEAL_SYN_RATE)*(1+DEAL_SYN_RATE)^0.5"
         "-SYN_RUNRATE*(1-DEAL_PHASE_Y1)*(1-DEAL_TAX)/(1+DEAL_SYN_RATE)^0.5-SYN_INTEGRATION",
         s.BOLD, s.USD_M, "Mid-year annuity over the capitalisation period, less the year-1 phase-in shortfall"),
        ("Capitalised value per pro forma share (US$)", "SYN_NPV_PS", "=SYN_NPV/PF_SHARES", s.FORMULA, s.USD_2, ""),
    ))


# --- Merger ---------------------------------------------------------------------------

def build_merger(ctx: ModelContext) -> None:
    deal = ctx.a._csv("deal_assumptions.csv").set_index("key")
    v = lambda key: float(deal.loc[key, "value"])
    ws = ctx.wb.create_sheet("Merger")
    s.title(ws, "Pro forma merger model — Centerra acquires Ero",
            "Offer structure, funding, pro forma ownership and accretion/dilution to EPS, CFPS and NAVPS.")
    ws.column_dimensions["B"].width = 46
    for c in "CDE":
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["F"].width = 70

    s.header_row(ws, 4, 2, ["Deal inputs (data/deal_assumptions.csv)", "Value", "", "", "Note"])
    r = _block(ctx, ws, 5, tuple(
        (label, name, v(key), s.INPUT, fmt, deal.loc[key, "note"] if deal.loc[key, "note"] == deal.loc[key, "note"] else "")
        for label, name, key, fmt in (
            ("Offer premium to unaffected price", "DEAL_PREMIUM", "offer_premium", s.PCT),
            ("Consideration (1 = all-stock, 2 = cash-and-stock)", "DEAL_CASE", "consideration_case", "0"),
            ("Cash portion when cash-and-stock", "DEAL_CASH_PCT_INPUT", "cash_pct", s.PCT),
            ("Minimum pro forma cash (US$M)", "DEAL_MIN_CASH", "min_cash_usd_m", s.USD_M),
            ("Cost of new acquisition debt", "DEAL_DEBT_RATE", "new_debt_rate", s.PCT),
            ("Yield forgone on cash used", "DEAL_CASH_YIELD", "cash_yield", s.PCT),
            ("Transaction costs (US$M)", "DEAL_TXN_COSTS", "transaction_costs_usd_m", s.USD_M),
            ("Tax rate on synergies and financing", "DEAL_TAX", "deal_tax_rate", s.PCT),
            ("Synergy phase-in, 2027E", "DEAL_PHASE_Y1", "synergy_phase_in_y1", s.PCT),
            ("Synergy phase-in, 2028E", "DEAL_PHASE_Y2", "synergy_phase_in_y2", s.PCT),
            ("Integration cost (x run-rate)", "DEAL_INTEGRATION_X", "integration_cost_multiple", s.MULT),
            ("Synergy capitalisation (years)", "DEAL_SYN_YEARS", "synergy_npv_years", "0"),
            ("Synergy discount rate", "DEAL_SYN_RATE", "synergy_discount_rate", s.PCT),
        )
    ))
    for row in range(5, r):
        if ws.cell(row, 2).value and ws.cell(row, 2).value.startswith(("Offer premium", "Consideration", "Cash portion")):
            ws.cell(row, 3).fill = s.KEY_FILL

    r += 1
    s.header_row(ws, r, 2, ["Offer and funding", "Value", "", "", "Note"])
    r = _block(ctx, ws, r + 1, (
        ("Ero unaffected share price (C$)", "DEAL_UNAFFECTED", "=ERO_PRICE_CAD", s.LINK, s.USD_2, "Market tab"),
        ("Offer price per Ero share (C$)", "DEAL_OFFER_CAD", "=DEAL_UNAFFECTED*(1+DEAL_PREMIUM)", s.BOLD, s.USD_2, ""),
        ("Cash portion applied", "DEAL_CASH_PCT", "=IF(DEAL_CASE=1,0,DEAL_CASH_PCT_INPUT)", s.FORMULA, s.PCT, ""),
        ("Cash per Ero share (C$)", "DEAL_CASH_PS", "=DEAL_OFFER_CAD*DEAL_CASH_PCT", s.FORMULA, s.USD_2, ""),
        ("Exchange ratio (CG shares per Ero share)", "DEAL_ER", "=DEAL_OFFER_CAD*(1-DEAL_CASH_PCT)/CG_PRICE_CAD",
         s.BOLD, "0.0000x", "Fixed at Centerra's unaffected price"),
        ("New Centerra shares issued (M)", "DEAL_NEW_SHARES", "=ERO_SHARES*DEAL_ER", s.FORMULA, s.USD_M, ""),
        ("Equity purchase price (US$M)", "DEAL_EQUITY_USD", "=DEAL_OFFER_CAD*ERO_SHARES*FX_SPOT", s.BOLD, s.USD_M,
         "At spot CAD/USD"),
        ("  of which cash (US$M)", "DEAL_CASH_USD", "=DEAL_EQUITY_USD*DEAL_CASH_PCT", s.FORMULA, s.USD_M, ""),
        ("  of which stock (US$M)", "DEAL_STOCK_USD", "=DEAL_EQUITY_USD-DEAL_CASH_USD", s.FORMULA, s.USD_M, ""),
        ("Ero net debt assumed (US$M)", "DEAL_NET_DEBT", "=-(ERO_DEBT+ERO_CASH)", s.FORMULA, s.USD_M, "Debt less cash"),
        ("Implied enterprise value (US$M)", "DEAL_EV_USD", "=DEAL_EQUITY_USD+DEAL_NET_DEBT", s.BOLD, s.USD_M, ""),
        ("Implied P/NAV (offer / Ero NAVPS)", "DEAL_PNAV", "=DEAL_OFFER_CAD/ERO_NAVPS_CAD", s.FORMULA, s.MULT,
         "Reserve-only NAV at the analyst deck"),
        ("Implied EV / 2027E EBITDA", "DEAL_EV_EBITDA", "=DEAL_EV_USD/ERO_EBITDA_2027", s.FORMULA, s.MULT, ""),
        ("Cash required: consideration + costs (US$M)", "DEAL_CASH_NEED", "=DEAL_CASH_USD+DEAL_TXN_COSTS", s.FORMULA,
         s.USD_M, ""),
        ("Funded from balance sheet cash (US$M)", "DEAL_CASH_USED",
         "=MIN(DEAL_CASH_NEED,MAX(0,CG_CASH+ERO_CASH-DEAL_MIN_CASH))", s.FORMULA, s.USD_M,
         "Combined cash above the minimum balance"),
        ("Funded from new debt (US$M)", "DEAL_NEW_DEBT", "=DEAL_CASH_NEED-DEAL_CASH_USED", s.FORMULA, s.USD_M, ""),
    ))

    r += 1
    s.header_row(ws, r, 2, ["Pro forma ownership", "Value", "", "", ""])
    r = _block(ctx, ws, r + 1, (
        ("Centerra shares outstanding (M)", None, "=CG_SHARES", s.LINK, s.USD_M, ""),
        ("Pro forma shares outstanding (M)", "PF_SHARES", "=CG_SHARES+DEAL_NEW_SHARES", s.BOLD, s.USD_M, ""),
        ("Centerra shareholders", "PF_OWN_CG", "=CG_SHARES/PF_SHARES", s.FORMULA, s.PCT, ""),
        ("Ero shareholders", "PF_OWN_ERO", "=DEAL_NEW_SHARES/PF_SHARES", s.FORMULA, s.PCT, ""),
    ))

    r += 1
    s.header_row(ws, r, 2, ["Accretion / dilution (US$M unless noted)", "2027E", "2028E", "", "Note"])
    r += 1
    cols = {2027: 3, 2028: 4}

    def acc(label, name, formula, fmt=s.USD_M, font=s.FORMULA, note=""):
        nonlocal r
        s.put(ws, r, 2, label, s.BOLD if font is s.BOLD else s.FORMULA)
        for year, c in cols.items():
            cell = s.put(ws, r, c, formula.replace("{Y}", str(year)), font, fmt)
            if font is s.BOLD:
                cell.border = s.TOTAL_BORDER
            if name:
                s.define(ctx.wb, f"{name}_{year}", ws.title, c, r)
        if note:
            s.put(ws, r, 6, note, s.NOTE)
        r += 1

    acc("Centerra standalone net income", "SA_NI", "=CG_NI_{Y}", font=s.LINK)
    acc("Ero standalone net income", None, "=ERO_NI_{Y}", font=s.LINK)
    acc("Pre-tax synergies", "PF_SYN", "=SYN_{Y}", font=s.LINK)
    acc("After-tax synergies", "PF_SYN_AT", "=PF_SYN_{Y}*(1-DEAL_TAX)")
    acc("After-tax interest on new debt", "PF_DEBT_COST", "=-DEAL_NEW_DEBT*DEAL_DEBT_RATE*(1-DEAL_TAX)")
    acc("After-tax interest forgone on cash", "PF_CASH_COST", "=-DEAL_CASH_USED*DEAL_CASH_YIELD*(1-DEAL_TAX)")
    acc("Pro forma net income", "PF_NI", "=SA_NI_{Y}+ERO_NI_{Y}+PF_SYN_AT_{Y}+PF_DEBT_COST_{Y}+PF_CASH_COST_{Y}",
        font=s.BOLD, note="Excludes one-off integration and transaction costs; no purchase price allocation step-up")
    acc("Centerra standalone EPS (US$)", None, "=CG_EPS_{Y}", s.USD_2, s.LINK)
    acc("Pro forma EPS (US$)", "PF_EPS", "=PF_NI_{Y}/PF_SHARES", s.USD_2, s.BOLD)
    acc("EPS accretion / (dilution)", "ACC_EPS", "=PF_EPS_{Y}/CG_EPS_{Y}-1", s.PCT, s.BOLD)
    acc("Pro forma operating cash flow", "PF_CFO",
        "=CG_CFO_{Y}+ERO_CFO_{Y}+PF_SYN_AT_{Y}+PF_DEBT_COST_{Y}+PF_CASH_COST_{Y}")
    acc("Centerra standalone CFPS (US$)", None, "=CG_CFPS_{Y}", s.USD_2, s.LINK)
    acc("Pro forma CFPS (US$)", "PF_CFPS", "=PF_CFO_{Y}/PF_SHARES", s.USD_2, s.BOLD)
    acc("CFPS accretion / (dilution)", "ACC_CFPS", "=PF_CFPS_{Y}/CG_CFPS_{Y}-1", s.PCT, s.BOLD)
    acc("Pro forma EBITDA (incl. synergies)", "PF_EBITDA", "=CG_EBITDA_{Y}+ERO_EBITDA_{Y}+PF_SYN_{Y}")

    r += 1
    s.header_row(ws, r, 2, ["NAV per share", "Value", "", "", "Note"])
    r = _block(ctx, ws, r + 1, (
        ("Centerra standalone NAV (US$M)", None, "=CG_NAV", s.LINK, s.USD_M, ""),
        ("Ero standalone NAV (US$M)", None, "=ERO_NAV", s.LINK, s.USD_M, ""),
        ("NPV of synergies, net of integration (US$M)", None, "=SYN_NPV", s.LINK, s.USD_M, ""),
        ("Less: cash consideration and transaction costs (US$M)", None, "=-DEAL_CASH_NEED", s.FORMULA, s.USD_M,
         "Paid from cash or new debt; either way it reduces NAV"),
        ("Pro forma NAV (US$M)", "PF_NAV", "=CG_NAV+ERO_NAV+SYN_NPV-DEAL_CASH_NEED", s.BOLD, s.USD_M, ""),
        ("Pro forma NAV per share (US$)", "PF_NAVPS", "=PF_NAV/PF_SHARES", s.BOLD, s.USD_2, ""),
        ("Centerra standalone NAVPS (US$)", None, "=CG_NAVPS_USD", s.LINK, s.USD_2, ""),
        ("NAVPS accretion / (dilution)", "ACC_NAVPS", "=PF_NAVPS/CG_NAVPS_USD-1", s.BOLD, s.PCT,
         "The metric that matters most for miners"),
        ("NAVPS accretion before synergies", "ACC_NAVPS_EX_SYN",
         "=(CG_NAV+ERO_NAV-DEAL_CASH_NEED)/PF_SHARES/CG_NAVPS_USD-1", s.FORMULA, s.PCT, ""),
        ("Pro forma net debt (US$M)", "PF_NET_DEBT", "=-(CG_CASH+ERO_CASH+CG_DEBT+ERO_DEBT)+DEAL_CASH_NEED",
         s.FORMULA, s.USD_M, "Combined net debt plus cash paid out"),
        ("Pro forma net debt / 2027E EBITDA", "PF_LEVERAGE", "=PF_NET_DEBT/PF_EBITDA_2027", s.BOLD, s.MULT, ""),
    ))


# --- Contribution ---------------------------------------------------------------------

def build_contribution(ctx: ModelContext) -> None:
    ws = ctx.wb.create_sheet("Contribution")
    s.title(ws, "Contribution analysis", "What each company brings versus what its shareholders receive.")
    ws.column_dimensions["B"].width = 38
    for c in "CDEF":
        ws.column_dimensions[c].width = 14
    s.header_row(ws, 4, 2, ["Metric", "Centerra", "Ero", "Centerra %", "Ero %"])
    ratio = "(Inputs!$AC$17/Inputs!$AC$16)"  # lb Cu per oz Au at the long-term deck
    au = lambda co: f"{ctx.refs[f'{co}.au_oz.2027']}+{ctx.refs[f'{co}.cu_lb.2027']}*1000/{ratio}"
    rows = (
        ("Net asset value (US$M)", "=CG_NAV", "=ERO_NAV", s.USD_0),
        ("Market capitalisation (US$M)", "=CG_SHARES*CG_PRICE_CAD*FX_SPOT", "=ERO_SHARES*ERO_PRICE_CAD*FX_SPOT", s.USD_0),
        ("2027E EBITDA (US$M)", "=CG_EBITDA_2027", "=ERO_EBITDA_2027", s.USD_0),
        ("2028E EBITDA (US$M)", "=CG_EBITDA_2028", "=ERO_EBITDA_2028", s.USD_0),
        ("2027E operating cash flow (US$M)", "=CG_CFO_2027", "=ERO_CFO_2027", s.USD_0),
        ("2027E net income (US$M)", "=CG_NI_2027", "=ERO_NI_2027", s.USD_0),
        ("2027E production (koz AuEq)", f"={au('CG')}", f"={au('ERO')}", s.USD_0),
        ("P&P reserves (Mt CuEq)", "=Comps!L5", "=Comps!L6", s.USD_2),
    )
    for i, (label, cg, ero, fmt) in enumerate(rows):
        r = 5 + i
        s.put(ws, r, 2, label)
        s.put(ws, r, 3, cg, s.LINK, fmt)
        s.put(ws, r, 4, ero, s.LINK, fmt)
        s.put(ws, r, 5, f"=C{r}/(C{r}+D{r})", s.FORMULA, s.PCT)
        s.put(ws, r, 6, f"=D{r}/(C{r}+D{r})", s.FORMULA, s.PCT)
    r = 5 + len(rows)
    s.put(ws, r, 2, "Pro forma ownership", s.BOLD)
    s.put(ws, r, 5, "=PF_OWN_CG", s.LINK, s.PCT)
    s.put(ws, r, 6, "=PF_OWN_ERO", s.LINK, s.PCT)
    s.put(ws, r + 2, 2, "AuEq converts copper at the long-term deck ratio. Centerra % above its ownership "
          "means Centerra contributes more than its holders keep.", s.NOTE)


# --- Scenarios and sensitivities ------------------------------------------------------

SCENARIOS = "Scenarios"
COPPER_GRID = (4.00, 4.50, 5.00, 5.50, 6.00, 6.50)
PREMIUM_GRID = (0.15, 0.20, 0.25, 0.30, 0.35, 0.40)
SYNERGY_GRID = (0, 50, 100, 155, 200)
CASH_GRID = (0.0, 0.15, 0.25, 0.35, 0.50)
# Scenario name, copper case, gold case, custom copper. Filled by scripts/run_scenarios.py.
SCENARIO_CASES = (("Deck", 1, 1, None), ("Spot", 2, 2, None)) + tuple(
    (f"Cu {p:.2f} flat", 3, 1, p) for p in COPPER_GRID
)
SCENARIO_OUTPUTS = ("CG_NAV", "ERO_NAV", "CG_NAVPS_CAD", "ERO_NAVPS_CAD", "CG_PNAV", "ERO_PNAV")
SCEN_FIRST_ROW = 6


def build_scenarios(ctx: ModelContext) -> None:
    ws = ctx.wb.create_sheet(SCENARIOS)
    s.title(ws, "Price scenarios",
            "Values produced by scripts/run_scenarios.py, which re-runs this workbook with each price case "
            "and records the outputs. Rerun `make model` after changing any input.")
    s.header_row(ws, 5, 2, ["Scenario", "Cu case", "Au case", "Custom Cu"] + list(SCENARIO_OUTPUTS))
    ws.column_dimensions["B"].width = 16
    for j in range(len(SCENARIO_OUTPUTS) + 3):
        ws.column_dimensions[s.col(3 + j)].width = 14
    for i, (label, cu_case, au_case, cu) in enumerate(SCENARIO_CASES):
        r = SCEN_FIRST_ROW + i
        s.put(ws, r, 2, label, s.BOLD)
        s.put(ws, r, 3, cu_case, s.INPUT, "0")
        s.put(ws, r, 4, au_case, s.INPUT, "0")
        s.put(ws, r, 5, cu, s.INPUT, s.USD_2)
        for j, output in enumerate(SCENARIO_OUTPUTS):
            fmt = s.MULT if output.endswith("PNAV") else (s.USD_2 if "NAVPS" in output else s.USD_M)
            s.put(ws, r, 6 + j, None, s.INPUT, fmt)
        s.define(ctx.wb, f"SCEN_{label.split()[0].upper()}_ROW" if cu is None else f"SCEN_CU_{int(cu * 100)}_ROW",
                 SCENARIOS, 2, r)


def _scen(output: str, row: int) -> str:
    return f"{SCENARIOS}!${s.col(6 + SCENARIO_OUTPUTS.index(output))}${row}"


def build_sensitivity(ctx: ModelContext) -> None:
    ws = ctx.wb.create_sheet("Sensitivity")
    s.title(ws, "Sensitivity tables", "Each cell recomputes the deal with the row and column values; "
            "everything else stays at the Merger tab inputs.")
    ws.column_dimensions["B"].width = 26
    for j in range(8):
        ws.column_dimensions[s.col(3 + j)].width = 12

    # Table 1: NAVPS accretion, premium x flat copper price (NAVs from the Scenarios tab).
    r = 4
    s.section(ws, r, "NAVPS accretion / (dilution): offer premium (rows) vs flat copper price, US$/lb (columns)", 9)
    s.header_row(ws, r + 1, 2, ["Premium \\ Copper"] + [p for p in COPPER_GRID])
    for j in range(len(COPPER_GRID)):
        ws.cell(r + 1, 3 + j).number_format = s.USD_2
    for i, prem in enumerate(PREMIUM_GRID):
        row = r + 2 + i
        s.put(ws, row, 2, prem, s.INPUT, s.PCT)
        for j, cu in enumerate(COPPER_GRID):
            srow = SCEN_FIRST_ROW + 2 + j
            cg, ero = _scen("CG_NAV", srow), _scen("ERO_NAV", srow)
            offer = f"(ERO_PRICE_CAD*(1+$B{row}))"
            new_sh = f"(ERO_SHARES*{offer}*(1-DEAL_CASH_PCT)/CG_PRICE_CAD)"
            cash = f"(ERO_SHARES*{offer}*DEAL_CASH_PCT*FX_SPOT+DEAL_TXN_COSTS)"
            formula = (f"=IFERROR(({cg}+{ero}+SYN_NPV-{cash})/(CG_SHARES+{new_sh})"
                       f"/({cg}/CG_SHARES)-1,\"n/a\")")
            s.put(ws, row, 3 + j, formula, s.FORMULA, s.PCT)
    r = r + 3 + len(PREMIUM_GRID)

    # Table 2: 2028E EPS accretion, run-rate synergies x cash portion.
    s.section(ws, r, "2028E EPS accretion / (dilution): run-rate synergies, US$M (rows) vs cash portion of offer (columns)", 9)
    s.header_row(ws, r + 1, 2, ["Synergies \\ Cash %"] + list(CASH_GRID))
    for j in range(len(CASH_GRID)):
        ws.cell(r + 1, 3 + j).number_format = s.PCT
    for i, syn in enumerate(SYNERGY_GRID):
        row = r + 2 + i
        s.put(ws, row, 2, syn, s.INPUT, s.USD_0)
        for j in range(len(CASH_GRID)):
            c = s.col(3 + j)
            pct = f"{c}${r + 1}"
            new_sh = f"(ERO_SHARES*DEAL_OFFER_CAD*(1-{pct})/CG_PRICE_CAD)"
            need = f"(DEAL_EQUITY_USD*{pct}+DEAL_TXN_COSTS)"
            used = f"MIN({need},MAX(0,CG_CASH+ERO_CASH-DEAL_MIN_CASH))"
            ni = (f"(CG_NI_2028+ERO_NI_2028+$B{row}*DEAL_PHASE_Y2*(1-DEAL_TAX)"
                  f"-({need}-{used})*DEAL_DEBT_RATE*(1-DEAL_TAX)-{used}*DEAL_CASH_YIELD*(1-DEAL_TAX))")
            s.put(ws, row, 3 + j, f"=IFERROR({ni}/(CG_SHARES+{new_sh})/CG_EPS_2028-1,\"n/a\")", s.FORMULA, s.PCT)
    r = r + 3 + len(SYNERGY_GRID)
    s.put(ws, r, 2, "Table 1 holds gold at the analyst deck and uses the Merger tab's cash portion. "
          "Table 2 holds the offer premium at the Merger tab input.", s.NOTE)


def build(ctx: ModelContext) -> None:
    build_synergies(ctx)
    build_merger(ctx)
    build_contribution(ctx)
    build_scenarios(ctx)
    build_sensitivity(ctx)
