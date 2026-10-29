# Testing

Two layers: **backend** (Python) and **frontend/UI** (Node + jsdom, under
`tests/`). Neither needs a real browser, and the UI tests don't need a
running server.

## Backend tests (Python)

```bash
pip install -r requirements.txt
python make_fixtures.py     # exercises upload + every table/text operation end-to-end
python tests/test_html_tables.py   # HTML table rowspan/colspan parsing (13 checks)
python tests/test_backend_fixes.py # find & replace, request locking, upload limit (18 checks)
```

If that writes `fixtures.json` without error, the backend is working.

## Frontend / UI tests (Node + jsdom)

```bash
cd tests
npm install
npm test              # runs ui_test.mjs (27 checks) + perf_test.mjs (3 checks)
```

These load the **real** `templates/index.html` in a headless DOM, wire
`fetch` to responses captured from the real backend, and simulate the exact
clicks a user makes - including opening the table editor popup, editing a
cell, sorting a column, and running the full merge-multiple-tables flow
(check 2 tables -> open merge popup -> combined rows appear, deduped header
-> delete a row -> save -> asserts the right `/table` and `/delete` calls
fire in the right order).

**Why `tests/` is a separate folder, and why it matters:** Flask's
auto-reloader (enabled by `debug=True`) watches the working directory for
file changes and restarts the whole process the moment it sees one - which
silently drops any request that's in flight. If `node_modules` (thousands of
files) sits in the same directory as `app.py`, running `npm install` there
while the server is running is a very effective way to make uploads "hang"
forever with no error, because the process gets killed mid-request and never
sends a response. `app.py` now runs with `use_reloader=False` for exactly
this reason - that's the actual fix - but keeping test tooling in its own
subfolder is good hygiene on top of it.

### Regenerating fixtures

`fixtures.json` (at the project root) holds real backend responses so the UI
test's mocked `fetch` matches production shapes. Regenerate whenever the
backend response format changes:

```bash
python make_fixtures.py
```
