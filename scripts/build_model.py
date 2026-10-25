"""Build model/lodestone_model.xlsx from the CSVs under data/.

The workbook is generated, never edited by hand: every number in it is either a formula
or an input copied from a CSV, so a reviewer can diff the CSVs to see what changed.
Run scripts/recalc.py afterwards to compute cached values (or use `make model`).
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from lodestone.assumptions import Assumptions  # noqa: E402
from lodestone.workbook import cover, football, merger, nav, precedents  # noqa: E402
from lodestone.workbook.context import ModelContext  # noqa: E402

OUTPUT = ROOT / "model" / "lodestone_model.xlsx"


def build(output: Path = OUTPUT) -> Path:
    ctx = ModelContext(wb=Workbook(), a=Assumptions())
    cover.build_cover(ctx)
    nav.build(ctx)
    precedents.build(ctx)
    merger.build(ctx)
    football.build(ctx)
    cover.build_sources(ctx)
    output.parent.mkdir(parents=True, exist_ok=True)
    ctx.wb.save(output)
    return output


if __name__ == "__main__":
    print(f"wrote {build().relative_to(ROOT)}")
