"""Pull share prices, shares outstanding, balance sheet items, CAD/USD and metal prices.

Writes timestamped CSVs to data/market/ and refreshes the `latest_*.csv` copies that the
model build reads. Every run keeps its own snapshot, so price changes stay auditable.

Usage:
    python scripts/fetch_market_data.py
"""

from __future__ import annotations

import argparse
import shutil
import sys
import warnings
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import yfinance as yf

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", message=".*inplace.*")

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
MARKET = DATA / "market"

# yfinance tickers for the macro inputs. Copper is COMEX HG (US$/lb), gold is COMEX GC (US$/oz).
MACRO_TICKERS = {
    "copper_usd_lb": "HG=F",
    "gold_usd_oz": "GC=F",
    "cadusd": "CADUSD=X",
}

INFO_FIELDS = {
    "price_local": "currentPrice",
    "price_currency": "currency",
    "shares_out": "sharesOutstanding",
    "market_cap_local": "marketCap",
    "week52_low": "fiftyTwoWeekLow",
    "week52_high": "fiftyTwoWeekHigh",
    "cash_fin": "totalCash",
    "debt_fin": "totalDebt",
    "ebitda_ttm_fin": "ebitda",
    "cfo_ttm_fin": "operatingCashflow",
    "eps_ttm_fin": "trailingEps",
    "financial_currency": "financialCurrency",
    "target_low": "targetLowPrice",
    "target_mean": "targetMeanPrice",
    "target_high": "targetHighPrice",
}


def last_close(ticker: str) -> float:
    history = yf.Ticker(ticker).history(period="5d")
    if history.empty:
        raise RuntimeError(f"No price history returned for {ticker}")
    return float(history["Close"].iloc[-1])


def fetch_macro() -> dict[str, float]:
    return {name: last_close(ticker) for name, ticker in MACRO_TICKERS.items()}


def fetch_company(ticker: str) -> dict[str, object]:
    info = yf.Ticker(ticker).info
    row: dict[str, object] = {"ticker": ticker}
    for column, key in INFO_FIELDS.items():
        row[column] = info.get(key)
    if row["price_local"] is None:
        row["price_local"] = last_close(ticker)
    return row


def to_usd_millions(frame: pd.DataFrame, cadusd: float) -> pd.DataFrame:
    """Add US$M columns so the model never mixes currencies or units."""
    fx = frame["price_currency"].map({"CAD": cadusd, "USD": 1.0})
    fin_fx = frame["financial_currency"].map({"CAD": cadusd, "USD": 1.0}).fillna(1.0)
    out = frame.copy()
    out["price_usd"] = out["price_local"] * fx
    out["market_cap_usd_m"] = out["price_usd"] * out["shares_out"] / 1e6
    for column in ("cash", "debt", "ebitda_ttm", "cfo_ttm"):
        out[f"{column}_usd_m"] = out[f"{column}_fin"] * fin_fx / 1e6
    out["ev_usd_m"] = out["market_cap_usd_m"] + out["debt_usd_m"] - out["cash_usd_m"]
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--universe", type=Path, default=DATA / "universe.csv")
    args = parser.parse_args()

    universe = pd.read_csv(args.universe)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%MZ")
    MARKET.mkdir(parents=True, exist_ok=True)

    macro = fetch_macro()
    macro_frame = pd.DataFrame([{"timestamp_utc": stamp, **macro}])

    rows = []
    for ticker in universe["ticker"]:
        try:
            rows.append(fetch_company(ticker))
        except Exception as exc:  # one bad ticker should not sink the whole pull
            print(f"warning: {ticker}: {exc}", file=sys.stderr)
    companies = to_usd_millions(pd.DataFrame(rows), macro["cadusd"])
    companies.insert(0, "timestamp_utc", stamp)

    outputs = {
        "macro": macro_frame,
        "companies": companies,
    }
    for name, frame in outputs.items():
        snapshot = MARKET / f"{name}_{stamp}.csv"
        frame.to_csv(snapshot, index=False)
        shutil.copyfile(snapshot, MARKET / f"latest_{name}.csv")
        print(f"wrote {snapshot.relative_to(ROOT)} ({len(frame)} rows)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
