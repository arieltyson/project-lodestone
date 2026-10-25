# Security Policy

Project Lodestone is a static analysis project: a data pipeline, an Excel model, a pitch
book and a Streamlit dashboard. It stores no user data and needs no credentials.

## Reporting a vulnerability

If you find a security issue (for example, a dependency with a known vulnerability or a
way to make the dashboard execute untrusted input), please open a
[private security advisory](../../security/advisories/new) on this repository rather than
a public issue. You can expect an acknowledgement within 7 days.

## Scope notes

- The data pipeline only reads public market data through `yfinance`; no API keys are used.
- The dashboard accepts numeric slider input only and does not execute uploaded files.
- Never commit secrets. `.streamlit/secrets.toml` is git-ignored.
