PY ?= .venv/bin/python

.PHONY: data model test

data:
	$(PY) scripts/fetch_market_data.py

model:
	$(PY) scripts/build_model.py
	$(PY) scripts/build_comps.py
	$(PY) scripts/recalc.py model/lodestone_model.xlsx
	$(PY) scripts/run_scenarios.py
	$(PY) scripts/recalc.py model/lodestone_model.xlsx

test:
	$(PY) -m pytest -q

.PHONY: deck
deck:
	$(PY) scripts/export_outputs.py
	cd deck && npm install --silent && node build_deck.js
	cd deck && soffice --headless --convert-to pdf pitch_book.pptx

.PHONY: site all
site:
	$(PY) scripts/build_site.py

all: model deck site test
