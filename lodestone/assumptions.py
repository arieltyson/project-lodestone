"""Load every model assumption from the CSV files under data/.

The CSVs are the single source of truth. Both the Excel build (scripts/build_model.py)
and the Python valuation engine read them through this module, so the two can never
drift apart on inputs.
"""

from __future__ import annotations

from dataclasses import dataclass
from functools import cached_property
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"

COMPANIES = ("CG", "ERO")


@dataclass(frozen=True)
class Assumptions:
    data_dir: Path = DATA

    def _csv(self, name: str, **kwargs) -> pd.DataFrame:
        return pd.read_csv(self.data_dir / name, **kwargs)

    @cached_property
    def settings(self) -> dict[str, str]:
        frame = self._csv("model_settings.csv")
        return dict(zip(frame["key"], frame["value"].astype(str)))

    @property
    def first_year(self) -> int:
        return int(self.settings["first_year"])

    @property
    def last_year(self) -> int:
        return int(self.settings["last_year"])

    @property
    def years(self) -> list[int]:
        return list(range(self.first_year, self.last_year + 1))

    @cached_property
    def price_deck(self) -> pd.DataFrame:
        """Deck expanded to one row per model year; the last deck year is held flat."""
        deck = self._csv("price_deck.csv").set_index("year").sort_index()
        expanded = deck.reindex(self.years).ffill()
        return expanded[["copper_usd_lb", "gold_usd_oz"]]

    @cached_property
    def discount_rates(self) -> pd.Series:
        return self._csv("discount_rates.csv").set_index("asset")["discount_rate"]

    @cached_property
    def discount_rationale(self) -> pd.Series:
        return self._csv("discount_rates.csv").set_index("asset")["rationale"]

    @cached_property
    def assets(self) -> pd.DataFrame:
        return self._csv("assets.csv").set_index("asset")

    def assets_of(self, company: str) -> list[str]:
        return list(self.assets.index[self.assets["company"] == company])

    @cached_property
    def corporate(self) -> pd.DataFrame:
        return self._csv("corporate.csv").set_index("company")

    @cached_property
    def macro(self) -> dict[str, float]:
        row = self._csv("market/latest_macro.csv").iloc[0]
        return {k: row[k] for k in ("copper_usd_lb", "gold_usd_oz", "cadusd")} | {
            "timestamp_utc": str(row["timestamp_utc"])
        }

    @cached_property
    def market(self) -> pd.DataFrame:
        return self._csv("market/latest_companies.csv").set_index("ticker")

    @cached_property
    def sources(self) -> pd.DataFrame:
        return self._csv("sources.csv")

    @cached_property
    def asset_notes(self) -> pd.DataFrame:
        return self._csv("asset_assumption_notes.csv")

    def optional(self, name: str) -> pd.DataFrame | None:
        path = self.data_dir / name
        return pd.read_csv(path) if path.exists() else None
