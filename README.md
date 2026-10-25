# Project Lodestone

A hypothetical buy-side M&A pitch recommending a merger of two TSX-listed mid-tier
copper-gold producers, with the full model behind it.

Work in progress. See [DISCLAIMER.md](DISCLAIMER.md).

## Layout

| Folder | Contents |
|---|---|
| `data/` | Every assumption as CSV, so changes show up as readable diffs |
| `model/` | The Excel model (NAV, comps, precedents, merger) |
| `scripts/` | Data pipeline and model build scripts |
| `lodestone/` | Python valuation engine |
| `app/` | Streamlit dashboard |
| `deck/` | Pitch book source and PDF |
| `tests/` | Reconciliation and unit tests |

## Setup

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```
