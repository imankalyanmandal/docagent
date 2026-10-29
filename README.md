# Doc Editor Prototype

## Run
```
pip install -r requirements.txt
python app.py
```
Open http://localhost:5000

**If you've run this before and don't see new buttons/features:** stop the
old `python app.py` process completely (Ctrl+C, then confirm it's not still
running), replace the old folder's files with these ones, then start it
fresh and hard-refresh the browser (Ctrl+Shift+R / Cmd+Shift+R, or open in a
private/incognito window). Browsers aggressively cache plain HTML/JS served
from Flask, and a stale tab or a server still running old code will look
identical to "nothing changed" even when the files on disk are correct.

## What you can do
Upload a **Base document** (edited + exported) and an **Other document**
(content source). Both render as checkbox-selectable blocks with
hierarchical ids like `1.2.1`.

**Cross-document**
- **Replace (1->1)** - one checked each side; overwrites the base block's
  content, keeping its position.
- **Insert after anchor** - one checked in base (the anchor) + one or more
  in other; splices copies in after the anchor, in source order.

**Within the base document**
- **Move up / down** - hover any block for arrow buttons. A heading moves
  as a whole section, stepping over the adjacent sibling section; a
  paragraph or table steps over the adjacent block.
- **Move to target** - check the blocks to move, then Ctrl/Cmd-click a
  destination block (highlighted amber), then click *Move to target*.
- **Delete** - one or more checked. The *incl. sub-content* toggle decides
  whether deleting a heading also removes everything nested under it.

**Editing text**
- Each block shows a compact **✎ ↑ ↓** control cluster. Click **✎** on a
  paragraph or header to edit its text in place; Enter or click-away saves,
  Escape cancels. **↑ / ↓** move the block (a heading moves its whole
  section).

**Table editor popup** - click **✎** on a table (or check one table and use
**Edit table…**):
- The top (shaded) row is the **header row**, shown with its real values in
  editable inputs - not generic "Col 1" labels.
- Each column header has a **sort** button (numeric-aware; click again to
  reverse - only the data rows sort, the header stays pinned on top) and a
  **✕ col** button to delete that column.
- Each data row has **↑ / ↓** to reorder it and **✕** to delete it.
- **+ Row** / **+ Column** to grow the table.
- Nothing is written until you click **Save**. All cell editing, row
  reordering, and adding/deleting rows & columns happen here - not in the
  panel outside.

**Merge tables** - check **two or more tables**, then **Merge selected
tables…** in the toolbar. Opens the same popup pre-filled with every checked
table's rows combined (duplicate header auto-dropped), each row tagged with
its source. Edit / sort / reorder / delete freely, then **Save** writes the
result into the first table and removes the others.

**Focus mode** - **Focus base** in the toolbar hides the Other-document pane
so the base document gets the full width and a taller editing area. Click
again to bring the other pane back.

**Export .docx** downloads the edited base document **without navigating the
page away** - you stay exactly where you are.

**Paste HTML (Froala)** - each upload box has a **Paste HTML…** button. Paste
HTML (as produced by Froala or anything else) and it's converted into editable
blocks - headings, paragraphs, captions, lists, and tables. You then choose
where it goes: **Append to base** (add to the current base doc), **Load as
Other document** (put it in the right pane to pull pieces from), or **Load as
Base document** (start a fresh base from it). Structure plus bold/italic/
underline and tables carry over; Froala's inline styles (fonts, colours, pixel
spacing) do not.

**Duplicate** - check one or more blocks and click **Duplicate** to clone them
in place, right after the selection. Duplicating a heading clones its whole
section (heading + everything nested under it).

**Find & replace** - **Find & replace…** in the toolbar replaces text across
the whole base document, including inside table cells. Options: case
sensitivity, and scope (everywhere / text only / tables only). Per-run
replacement keeps formatting, so a bolded word stays bold after replacing.

**Outline** - **Outline** opens a sidebar listing every heading (indented by
level) with a live word count. Click any heading to smooth-scroll to it in the
base pane, with a brief highlight.

**Undo** steps back one operation (20 deep per document).
**Export .docx** downloads the edited base document.

## Architecture

### The live document is the single source of truth
The .docx is opened once and its python-docx `Document` object stays alive
for the session. Every operation mutates that real XML tree, and afterwards
we **re-derive** block ids and ordering from the document (`resync()`)
rather than maintaining a parallel structure.

That's why ids stay honest: insert two sections before "Scope" and Scope
renumbers `1.2 -> 1.4` by itself; delete a heading without cascade and the
orphaned children re-parent cleanly. Re-deriving is O(n) per operation -
negligible here, and it removes a whole class of drift bugs.

### Documents are never rebuilt from scratch
Mutating existing XML rather than writing a fresh document is what keeps
formatting:
- Untouched content keeps 100% of its original formatting.
- **Same-type replace** (paragraph<->header, table<->table) mutates the
  target in place, so it keeps the **target's** style - font, colour, table
  borders and shading - even when the table's shape changes. New rows and
  cells are cloned from existing ones so they inherit their formatting.
- **Cross-type replace** (a table replacing a paragraph) can't mutate in
  place, so the source element is cloned into the target's position -
  keeping the **source's** formatting instead.
- **Insert** deep-copies the source element, so inserted content arrives
  with the other document's own styling.
- **Move** relocates existing elements without touching them at all.

### Undo, and failure atomicity
Every mutating route runs inside a `mutation(entry)` context manager that
snapshots the document first, resyncs after, and **discards the snapshot if
the operation raises**. A rejected operation therefore never consumes an
undo slot or leaves a half-applied state.

## Endpoints
| Method | Route | Body |
|---|---|---|
| POST | `/upload` | multipart file |
| POST | `/upload_html` | `{html, name?}` (new doc) or `{html, append_to}` (append into a doc) |
| POST | `/replace` | `{base_doc_id, base_block_id, source_doc_id, source_block_id}` |
| POST | `/insert` | `{base_doc_id, anchor_block_id, source_doc_id, source_block_ids[]}` |
| POST | `/move` | `{doc_id, block_id, direction:'up'\|'down'}` |
| POST | `/move_to` | `{doc_id, block_ids[], anchor_block_id, position:'before'\|'after'}` |
| POST | `/delete` | `{doc_id, block_ids[], cascade}` |
| POST | `/table` | `{doc_id, block_id, operation, ...params}` |
| POST | `/edit_text` | `{doc_id, block_id, text}` |
| POST | `/duplicate` | `{doc_id, block_ids[]}` (a heading clones its whole section) |
| POST | `/find_replace` | `{doc_id, find, replace, case_sensitive?, scope?}` → `{..., replacements}` |
| POST | `/undo` | `{doc_id}` |
| GET | `/export/<doc_id>` | - |

`/table` operations: `add_row` `delete_row` `add_column` `delete_column`
`sort` `transpose` `merge` `set_content`. `set_content` takes `{rows: [[...]]}`
- a full replacement matrix - and is what the table editor popup's Save
button calls; ragged rows are padded rather than rejected.

All mutating routes return `{doc_id, blocks, can_undo}` so the client just
re-renders from the server's authoritative state.

## Deliberate design choice: where ambiguity stops
The table operations only do what's unambiguous. **Merge refuses when the
two tables have different column counts**, with an error saying so, rather
than guessing whether to drop, pad, or map the extra column. That refusal
is the seam for the LLM layer: the model resolves the ambiguity (choosing a
column mapping) and then calls these same deterministic tools to execute
it. Same for "combine these paragraphs sensibly" - judgement in the model,
execution in tested code.

## Known limitations
- **Empty paragraphs are skipped** by the parser, so pure-spacing paragraphs
  aren't selectable. They stay in the document and export fine.
- **Replace loses per-word formatting from the source** - a same-type
  paragraph replace reuses the target's first run, so a single bold word
  mid-sentence in the source is dropped; the target's style wins.
- **Images, headers/footers, footnotes and content controls aren't parsed**
  into blocks. They survive export untouched but can't be selected or moved
  - so moving a section that contains an image will leave the image behind.
- **Moving a paragraph down past a heading** puts it at the end of that
  heading's section rather than immediately after the heading.
- **State is in-memory** - restarting clears everything, docs evict past
  `MAX_DOCS` (50), and there's no auth: any session can address any
  `doc_id` it knows.

## UI notes

- **Text editing** (paragraphs/headers) is inline via the ✎ button.
  **Table values and column sorting are popup-only, by design** - hover a
  table's ✎ Edit table button or use the table panel's *Edit table…*. This
  keeps table structure edits (which need a real grid, not a text cursor)
  separate from quick paragraph tweaks.
- **Merging tables is a popup workflow, not a one-click dropdown.** Check
  2+ tables, click *Merge selected tables…*, and you get a combined,
  editable, sortable grid to shuffle rows, fix mismatched columns, or
  delete anything before committing - rather than the tool silently
  guessing how to align columns that don't match.
- **Table operations panel** appears when exactly one table (and nothing
  else) is checked in the base pane. If you check a table *plus* other
  blocks, the panel stays visible but its buttons are disabled and it shows
  a hint telling you why - it never silently does nothing.
- **Selecting blocks is cheap.** The pane is painted once and selection only
  toggles a CSS class; the DOM is rebuilt only when the block *structure*
  actually changes (after an edit). Scroll position is preserved across
  edits, and wide tables get their own horizontal scrollbar.
- **Uploads (and every other request) no longer hang.** The dev server
  previously ran with Flask's auto-reloader on, which restarts the whole
  process - silently dropping any in-flight request - the instant it sees a
  file change anywhere in the working directory. It's now started with
  `use_reloader=False`. A global error handler also ensures any unexpected
  server-side exception returns clean JSON instead of an HTML debug page,
  so a failure always surfaces as a visible error instead of a stuck
  "loading" state.
- **Action buttons disable while a request is in flight**, so a double-click
  can't stack two operations.

## Testing

See `TESTING.md`. In short:

```bash
python make_fixtures.py    # regenerate backend response fixtures if the API changes
cd tests
npm install                 # jsdom for the headless UI tests
npm test                    # runs ui_test.mjs (27 checks) + perf_test.mjs (3 checks)
```

Test tooling lives in `tests/`, separate from `app.py`, so `npm install`
never puts `node_modules` in the server's own working directory.

The UI tests load the real `templates/index.html` in jsdom and drive it with
`fetch` backed by responses captured from the real backend, so they verify
the actual wiring (buttons -> requests -> re-render) without a browser or a
running server.

## Suggested next steps
1. **Persistence** (docs on disk + metadata in SQLite) - the biggest
   robustness gap; a restart currently loses everyone's work.
2. **Snippet library** - save blocks for reuse across sessions, so common
   clauses and standard tables don't need a source doc re-upload.
3. **LLM agent layer** on top of these tools, handling the ambiguous cases
   the deterministic ops deliberately refuse.
4. **Virtualized rendering** - the render path is now efficient for edits,
   but a document with thousands of blocks would still paint them all at
   once. Windowed rendering (only paint visible blocks) would be the next
   scale step if needed.
