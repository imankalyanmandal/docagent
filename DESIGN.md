# Doc Editor — Design Document

A working-notes design doc: what the system is, how the pieces fit, what was
built in each round of changes, and where the planned HTML/Froala integration
would attach. For the deeper "why" behind the core decisions, see
`ARCHITECTURE.md`; this doc is the higher-level map.

---

## 1. What the tool is

Upload a Word document, see it broken into editable **blocks** (headings,
paragraphs, tables), edit / reorder / delete them, pull content in from a
second document, and export a Word file back out — with the original
formatting preserved wherever possible.

The whole thing is a small Flask app (`app.py`) plus a single-page frontend
(`templates/index.html`). State lives in memory for the session. No database
yet.

---

## 2. The one idea everything hangs off

**The uploaded `.docx` is opened once and kept alive in memory. Every edit
mutates that real document object. We never rebuild it from scratch.**

Two consequences flow from this and explain most of the design:

1. **Formatting survives.** Because we edit the actual Word XML in place
   (change a run's text, clone a table column) rather than regenerating the
   file, anything we don't touch keeps its exact original styling.

2. **Block IDs are always honest.** After every edit we *re-derive* the block
   list from the live document (`resync()`) instead of maintaining a separate
   bookkeeping structure. So IDs, ordering, and hierarchy always match reality
   — insert two sections and the ones after them renumber automatically;
   delete a heading and its orphaned children re-parent cleanly.

The document is the single source of truth. The block list the browser sees
is just a read-only *view* of it, recomputed on demand.

---

## 3. The pieces

```
                    ┌──────────────────────────────────────────┐
                    │            Browser (index.html)           │
                    │  - renders blocks as a checkbox list      │
                    │  - inline text edit (✎ on a block)        │
                    │  - table editor popup (values/sort/rows)  │
                    │  - merge-tables popup                     │
                    │  - focus mode, export, undo               │
                    └───────────────┬──────────────────────────┘
                                    │ JSON over fetch()
                                    │ every mutating call returns
                                    │   { doc_id, blocks, can_undo }
                    ┌───────────────▼──────────────────────────┐
                    │              Flask app (app.py)           │
                    │                                           │
                    │  DOC_STORE: doc_id -> {                   │
                    │     document  (live python-docx object)   │
                    │     blocks    (JSON view for the browser) │
                    │     refs      (block_id -> live XML obj)  │
                    │     undo      (byte snapshots)            │
                    │  }                                        │
                    │                                           │
                    │  parse:  docx -> blocks + refs            │
                    │  ops:    op_replace / op_move / op_table* │
                    │  export: live document -> .docx bytes     │
                    └───────────────────────────────────────────┘
```

### The block model (the contract between front and back)

Every block is one of three types, each with a hierarchical id derived from
heading structure (`1`, `1.1`, `1.2.1`, …):

| type        | shape                                             |
|-------------|---------------------------------------------------|
| `header`    | `{id, type, level, text, order}`                  |
| `paragraph` | `{id, type, text, order}`                         |
| `table`     | `{id, type, order, rows: [[cell, …], …]}`         |

This model is the important abstraction: **everything downstream operates on
blocks, not on Word internals.** That's exactly why HTML input (section 7)
slots in so cleanly — it just has to produce blocks.

---

## 4. How an operation flows (end to end)

Example: user reorders a section.

1. Browser sends `POST /move {doc_id, block_id, direction}`.
2. Route wraps the work in the `mutation(entry)` context manager, which:
   - snapshots the document's bytes (for undo),
   - runs `op_move(...)` which relocates the real XML element(s),
   - on success calls `resync()` to rebuild the block view,
   - on failure discards the snapshot so a rejected op costs nothing.
3. Route returns `{doc_id, blocks, can_undo}`.
4. Browser replaces its block view with the response and re-renders.

The browser **never computes state locally** — it always re-renders from the
server's authoritative response, because IDs can renumber after any op.

---

## 5. What each round of changes added

The tool grew in layers. Roughly in order:

**Round 1 — parse & view.** docx → hierarchical block JSON, rendered as an
editable list in the browser. Confirmed the round-trip idea.

**Round 2 — copy between docs + export.** A second "other" document you could
pull blocks from; `replace` and `export`. First proof that formatting
survived a real edit.

**Round 3 — formatting preservation, done properly.** Switched from
"rebuild the document" to "mutate the live document in place." This is when
the single-source-of-truth architecture landed.

**Round 4 — multi-select.** Insert several blocks from the other doc after an
anchor; delete several at once, with cascade (deleting a heading takes its
whole section).

**Round 5 — optimisation & robustness.** The big cleanup:
- `resync()` — re-derive blocks after each op instead of patching them,
  killing a class of ID-drift bugs.
- **Undo** via whole-document byte snapshots.
- The `mutation()` context manager, so a failed op never corrupts undo
  history or leaves half-applied state.

**Round 6 — reordering + deterministic table ops.** Move up/down (headings
move as whole sections), move-to-anchor, and table operations: add/delete
row & column, sort (numeric-aware), transpose, merge. Merge deliberately
*refuses* to guess when column counts differ — that refusal is the designed
handoff point for a future LLM layer.

**Round 7 — the `tblGrid` bug.** Real-world / tool-generated docx files can
omit the `<w:tblGrid>` element, which made python-docx throw on any table
resize. Added `ensure_tblgrid()` to rebuild a valid grid before table ops —
which also *repairs* malformed tables on the way through.

**Round 8 — UI performance & scroll.** Selecting a block used to rebuild the
entire pane's DOM (~1400 nodes, ~130 ms on a big doc). Rewrote rendering to
paint once and only toggle a CSS class on selection (0 nodes, ~15 ms), and
to preserve scroll position across edits.

**Round 9 — editing, popups, and the current UI (most recent).** This is the
round that made it feel like a real editor:
- **Inline text editing** for paragraphs/headers (`/edit_text` endpoint) via
  the ✎ button, which sits in one compact control cluster with the ↑/↓ move
  buttons.
- **Table editor popup** — the *only* place table values and column sorting
  are edited. The first row is treated as the header row and shown with its
  real values (not "Col 1"); each column header has sort + delete-column
  controls; each data row has up/down reorder + delete controls; nothing
  saves until you click Save. Backed by a new `set_content` table op that
  overwrites the whole matrix in place.
- **Merge popup** — check 2+ tables, get them combined into one editable grid
  (duplicate header auto-dropped, each row tagged with its source), shuffle /
  edit / delete rows, then Save writes the result into the first table and
  removes the rest.
- **Focus mode** — hide the "other document" pane for full-width editing of
  the base.
- **Export without navigating away** — download via a temporary anchor
  instead of `window.location.href`, so the page stays put.
- **Reliability fix:** the dev server previously ran with Flask's auto-reloader
  on, which restarts the process on any file change and silently drops
  in-flight requests — the actual cause of "uploads hang forever." Now runs
  with `use_reloader=False`, plus a global error handler so any server-side
  exception returns clean JSON instead of an HTML debug page.

---

## 6. The API surface

All mutating routes return `{doc_id, blocks, can_undo}`.

| Method | Route | Purpose |
|---|---|---|
| POST | `/upload` | parse a `.docx` into blocks |
| POST | `/replace` | overwrite one base block's content with a source block's |
| POST | `/insert` | splice source blocks in after an anchor |
| POST | `/move` | move a block/section up or down |
| POST | `/move_to` | relocate blocks to before/after an anchor |
| POST | `/delete` | delete blocks (optional cascade to sub-content) |
| POST | `/edit_text` | overwrite a paragraph/header's text |
| POST | `/table` | table ops: `add_row` `delete_row` `add_column` `delete_column` `sort` `transpose` `merge` `set_content` |
| POST | `/undo` | step back one operation |
| GET | `/export/<doc_id>` | download the edited `.docx` |

The operation functions behind these routes (`op_replace`, `op_move`,
`op_table_*`, …) are deliberately small, deterministic, and independently
tested — which is what makes them a clean **tool surface for a future LLM
layer** to call.

---

## 7. Where HTML / Froala input would attach (planned, not built)

The key realisation: **the docx parser's only job is `docx -> blocks`. HTML
is just a second parser feeding the same block model.** Everything downstream
— editing, reordering, merging, export — is unchanged because it only knows
about blocks.

### The intended flow (as discussed)

Several Froala editors each emit an HTML chunk. Concatenate the chunks into
one HTML document, then produce a Word file from it:

```
Froala #1 HTML ┐
Froala #2 HTML ┼─► merge into one HTML ─► html_to_blocks() ─► (edit?) ─► build .docx ─► export
Froala #3 HTML ┘
```

### Two shapes this could take (decision still open)

- **A — pure converter.** Merge the HTML, convert straight to a `.docx`,
  download. No editing, no block UI. Smallest possible addition: essentially
  one `html_to_docx()` function and an endpoint.

- **B — HTML becomes a document in the existing editor.** The merged HTML is
  parsed into blocks and loaded into `DOC_STORE` exactly like an uploaded
  docx — so it's editable, reorderable, mergeable, and *then* exportable.
  Reuses everything already built. Worth it only if the assembled result
  needs editing inside this tool rather than in Froala.

### Mapping HTML to the block model

| HTML | block |
|---|---|
| `<h1>`–`<h6>` | `header` (level from tag) |
| `<p>` | `paragraph` |
| `<table>` | `table` (rows/cells) |
| `<ul>` / `<ol>` | open question: paragraphs, or a new `list` block type |

### The one real caveat

**HTML → docx breaks the formatting-preservation guarantee**, by necessity.
There is no original Word XML to preserve, so HTML-sourced content must be
*built fresh* in python-docx (`<b>`→bold run, `<table>`→table, etc.). That's
the "rebuild from scratch" path we deliberately avoid for docx — it's fine
and expected here, just a different code path with different fidelity. How
much of Froala's inline styling (fonts, colours, spacing) to carry over is a
scope decision: basic structure + bold/italic/lists is straightforward;
pixel-exact styling is materially more work.

### Practical notes for whenever this gets built

- **Froala HTML is messy** (inline styles, `<span>` soup, `<div>` wrappers).
  A sanitising pass (BeautifulSoup/lxml) to reduce it to semantic structure
  should come before `html_to_blocks()`.
- **Source of the HTML**: pasted/uploaded by the user, vs. fetched
  server-side from the tool's own endpoint — different integrations; the
  fetch version couples to your tool's auth.
- **Reverse direction** (blocks → HTML) is easy to add and would make the
  loop bidirectional if edited content ever needs to flow back to Froala.

### Smallest first step, when ready

A `/upload_html` endpoint (or a "paste HTML" box) that runs `html_to_blocks()`
and drops the result into `DOC_STORE` as an "other document." That alone lets
Froala content be pulled into a base docx using every operation that already
exists — no other changes required.

---

## 8. Known limitations (current)

- Images, headers/footers, footnotes, content controls aren't parsed into
  blocks (they survive export untouched but can't be selected/moved).
- Empty spacing paragraphs are skipped by the parser.
- Same-type text replace keeps the target's first-run formatting only.
- State is in-memory: a restart loses everything; no auth; docs evict past
  `MAX_DOCS`.
- **Persistence is the biggest robustness gap** — docs-on-disk + a small DB
  is the highest-value non-feature work before real team use.

---

## 9. Testing (how to trust changes)

- **Backend**: `python make_fixtures.py` exercises upload + every op via
  Flask's test client.
- **Frontend**: `cd tests && npm install && npm test` — loads the real
  `index.html` in jsdom, drives the actual buttons/popups with `fetch` backed
  by responses captured from the real backend, and includes a performance
  test asserting selection doesn't rebuild the DOM. ~47 checks total.

The frontend tests mock `fetch` against *captured real backend responses*
(regenerated by `make_fixtures.py`), so the mock can't silently drift from
the real API as long as fixtures are regenerated when the API changes.
