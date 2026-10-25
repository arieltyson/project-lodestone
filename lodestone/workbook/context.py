"""State shared by the tab builders while the workbook is assembled."""

from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl import Workbook

from lodestone.assumptions import Assumptions

# Time series start in column F on the Inputs tab and every asset tab, so a year always
# sits in the same column and cross-sheet links stay readable (Inputs!H19 is 2029).
FIRST_YEAR_COL = 6


@dataclass
class ModelContext:
    wb: Workbook
    a: Assumptions
    refs: dict[str, str] = field(default_factory=dict)

    def year_col(self, year: int) -> int:
        return FIRST_YEAR_COL + year - self.a.first_year

    @property
    def last_col(self) -> int:
        return self.year_col(self.a.last_year)
