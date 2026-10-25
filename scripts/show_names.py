"""Print the cached value of workbook defined names matching a prefix (debug helper)."""

import sys
from pathlib import Path

from openpyxl import load_workbook

wb = load_workbook(Path(__file__).resolve().parents[1] / "model" / "lodestone_model.xlsx", data_only=True)
prefixes = tuple(sys.argv[1:]) or ("",)
for name, dn in sorted(wb.defined_names.items()):
    if name.startswith(prefixes):
        sheet, coord = next(iter(dn.destinations))
        value = wb[sheet][coord.replace("$", "")].value
        print(f"{name:24s} {value:,.4f}" if isinstance(value, (int, float)) else f"{name:24s} {value}")
