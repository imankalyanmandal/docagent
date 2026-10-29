# Testing

Two layers: **backend** (Python) and **frontend/UI** (Node + jsdom, under
`tests/`). Neither needs a real browser, and the UI tests don't need a
running server.

## Backend tests (Python)

Run everything at once:

```bash
pip install -r requirements.txt pillow
python tests/run_all.py            # all Python suites, combined result (311 checks)
```

Or individually:

```bash
python make_fixtures.py     # exercises upload + every table/text operation end-to-end
python tests/test_html_tables.py   # HTML table rowspan/colspan parsing (13 checks)
python tests/test_backend_fixes.py # find & replace, request locking, upload limit (18 checks)
python tests/test_fidelity.py      # formatting fidelity + tracked-changes export (33 checks)
python tests/test_table_corner_cases.py  # merged cells (colspan/rowspan) through every table op (172 checks)
python tests/test_corner_cases.py  # move/delete/duplicate/undo/find-replace/edit/insert/export edges (68 checks)
```

If that writes `fixtures.json` without error, the backend is working.

### What the merged-table suite checks

`test_table_corner_cases.py` builds tables with a colspan header, a 3-row
rowspan and a 2x2 block, then runs every table operation on them through the
real HTTP routes - value edits (including in a *covered* position of a merged
cell), add/delete row and column at the start, end and *inside* spans, sort,
transpose, merging two tables, the table-editor popup's shape changes (row
delete/add/reorder, column delete), replacing a table with a merged one, and
HTML tables with rowspan/colspan. Odd structures are covered too: missing
tblGrid, nested tables, rows starting late (`gridBefore`), 1x1 tables,
full-width and full-height merges.

After every edit it asserts the table is still valid Word XML (every row
covers the grid width; every vMerge continuation sits under a matching merged
cell; cell widths match their grid span; the file saves and re-opens) and at
the end opens every resulting document in LibreOffice (skipped if not
installed). Pillow is needed for the image tests.

## Frontend / UI tests (Node + jsdom)

```bash
cd tests
npm install
npm test              # runs ui_test.mjs (81 checks) + perf_test.mjs (3 checks)
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
