PY ?= .venv/bin/python

.PHONY: data model test

data:
	$(PY) scripts/fetch_market_data.py

model:
	$(PY) scripts/build_model.py
	$(PY) scripts/recalc.py model/lodestone_model.xlsx

test:
	$(PY) -m pytest -q
