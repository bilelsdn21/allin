# Changelog

Bump the version in `passport_check.py` (`__version__`) and add a line here
every time the app changes. The version is shown in the window title, in
`report.txt` / `report.xlsx`, and in every archived run - so during testing
you always know which version produced which result.

## 1.1.0 - 2026-07-20
- Version number shown in the app title, reports, and logs.
- Every photo decision is now logged to `decisions.csv` (what matched, why).
- Every run is archived in `runs/<date_time>/` (reports + decision log +
  `summary.json`) so testing runs can be compared - nothing is overwritten.
- Added automated test suite (`tests/`, run with `python -m pytest tests`).
- Matching behavior unchanged: results are identical to 1.0.

## 1.0 - before 2026-07-20
- Initial app: Excel matching, ID stamping, bilingual reports, PDF splitting.
