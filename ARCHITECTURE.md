# Doc Editor Prototype — Architecture & Design Logic

This document explains *why* the code is built the way it is, not just what
each function does. If you're picking this project back up later, or handing
it to someone else, start here.

---

## 1. The problem this architecture solves

The obvious way to build "upload a Word doc, edit it, export it" is:

```
docx → parse to JSON → let the user/agent edit the JSON → rebuild a docx
```

That's what a first version of this project did. It's simple, but it throws
away Word formatting: fonts, heading colors, table borders and shading,
theme, margins — none of that lives in a plain-text JSON block, so rebuilding
from JSON always produces a plainer, blander document than what came in.

The fix changes one thing, but it's the thing everything else depends on:

> **Never rebuild the document. Keep the real one open and mutate it.**

The uploaded `.docx` is opened once with `python-docx` and the resulting
`Document` object is kept alive in server memory for the whole session.
Every operation — replace, insert, delete, move, table edits — reaches in
and mutates the *actual* XML elements inside that object. Export just saves
the same object back out. Nothing is ever thrown away and rebuilt, so
anything you don't touch keeps its exact original formatting, and even
edited content usually keeps its own style because we're mutating the
existing element rather than manufacturing a fresh, unstyled one.

Everything else in this document follows from that one decision.

---

## 2. The data model: blocks, ids, and the document as source of truth

### What a "block" is

The document is walked top-to-bottom and split into three kinds of blocks:

| Type        | What it maps to                          |
|-------------|-------------------------------------------|
| `header`    | A paragraph styled `Heading N` or `Title` |
| `paragraph` | Any other non-empty paragraph              |
| `table`     | A table, represented as `rows: [[cell, ...], ...]` |

Each block gets a **hierarchical id** derived from heading structure, e.g.:

```
1          Project Overview      (Heading 1)
1.1        Intro paragraph.
1.2        Scope                 (Heading 2, nested under 1)
1.2.1      Scope text.
1.2.2      [table]
2          Budget                (Heading 1)
2.1        Budget text.
```

A block's id tells you exactly where it sits in the outline — same scheme
Word's own "navigation pane" numbering implies, just made explicit and
addressable.

### The critical decision: ids are *derived*, never stored

Early on, the block list and the live document were two separate things kept
in sync by hand — insert a block, remember to also update its neighbors'
order fields; delete a block, remember to renumber its children. This drifts
in exactly the way you'd expect: synthetic patched-on ids (`ins_a1b2c3d4`),
fractional order values, and orphaned children left pointing at a deleted
parent's id.

The fix was to stop maintaining a parallel structure entirely:

```python
def resync(doc_entry):
    """Re-derive blocks/refs from the live document. Call after any mutation."""
    doc_entry["blocks"], doc_entry["refs"] = extract_blocks_and_refs(doc_entry["document"])
```

After *every* mutating operation, the block list is thrown away and rebuilt
by re-walking the actual document XML. The live document is the only source
of truth; `blocks` is just a read-only view of it, recomputed on demand.

This is what makes the following "just work" instead of needing special-case
logic:

- Insert two new sections before "Scope" → Scope automatically renumbers
  from `1.2` to `1.4`, because that's genuinely where it is now.
- Delete a heading *without* cascading to its children → the orphaned
  children get folded into whatever section now precedes them, instead of
  keeping stale ids that point at nothing.

Re-deriving the whole block list is O(n) in the number of paragraphs/tables
— for realistic documents (tens to low hundreds of blocks) this is
imperceptible, and it eliminates an entire category of drift bugs in
exchange for a small, constant cost per edit.

### `refs`: the bridge between JSON and the live document

Alongside `blocks` (the JSON view sent to the browser), `extract_blocks_and_refs`
also builds `refs`: a `block_id → python-docx Paragraph/Table object` map.
`blocks` is what the frontend reads; `refs` is what the backend writes
through. Every mutation looks up the real object via `refs`, edits it in
place, and only then calls `resync()` to get a fresh, correct `blocks` view.

---

## 3. How formatting survives an edit

Different operations touch the document differently, and each is
deliberately shaped around *whose* formatting should win.

### Same-type replace → mutate in place, target's look + source's emphasis

Replacing a paragraph with another paragraph, or a table with another
table, reuses the *target's* existing XML element: its paragraph properties
(style, spacing, numbering) are kept, and its first text run supplies the
base look (font, size, colour). The source's runs are then copied in and
restyled onto that base, keeping only the source's *emphasis* - see
"Same-type replace keeps the source's emphasis" below. A bold word
mid-sentence in the source is still bold in the result; a different font in
the source is not imported.

Tables work the same way — `set_table_content` resizes the existing table
(see below) and fills cell text, then `fidelity.copy_table_content` upgrades
each cell to the source cell's rich content, so the table's style, borders,
and shading are never touched, only its content.

Plain-text editing (`set_paragraph_text`, used by inline edit and the table
editor) keeps the first text run's formatting and leaves non-text runs -
inline images, fields - where they are.

### Resizing a table without losing its style

Word tables carry their formatting in two places: a `tblGrid` (column
widths) and per-cell `tcPr` (borders, shading). Naively deleting and
re-adding columns loses both. Instead, `resize_table_columns` **clones the
last existing column's XML** when growing, and only removes the trailing
`gridCol`/`tc` elements when shrinking:

```python
new_tc = deepcopy(tcs[-1])                      # clone, don't invent
for t in new_tc.findall('.//' + qn('w:t')):
    t.text = ''                                  # then blank the cloned text
row._tr.append(new_tc)
```

A new column inherits the shading/border of its neighbor instead of
appearing as a bare, unstyled cell. `resize_table_rows` similarly uses
`table.add_row()` (which itself clones the last row) or trims trailing rows.

### Cross-type replace → clone the source, source's formatting wins

Replacing a paragraph's *position* with a table's content (or vice versa)
can't be done by editing text in place — the element kinds don't match.
Here the source element is deep-copied wholesale and swapped into the
target's position:

```python
clone = deepcopy(source_ref._element)
parent.insert(parent.index(target_element), clone)
parent.remove(target_element)
```

Because the clone *is* the source's own XML, it arrives with the source's
own formatting — this is the one case where the *source's* style wins
instead of the target's, and that's intentional: there's no sensible
"target style" to preserve when the target was a completely different kind
of content.

### Insert → clone the source, source's formatting wins (same reasoning)

Inserting blocks from the other document is structurally the same move as
cross-type replace: deep-copy each selected source element and splice it
into the base document. Inserted content keeps its *own* original
formatting rather than being reflowed into the base document's style —
which matches the mental model of "pasting a section from another doc."

### Cross-document copies → re-point what the XML references

A paragraph's XML is not self-contained: images and hyperlinks are referenced
by relationship id, list bullets by numbering id, styles by style id - and
those ids only mean something inside their own document. Every element that
crosses from the Other document into the base (insert, cross-type replace,
HTML append) goes through `fidelity.Importer`, which copies the image into
the base package, re-creates the hyperlink relationship, gives an imported
list its own numbering definition, and copies styles the base doesn't have.
One Importer is shared per operation, so ten items from one source list map
onto one new list and keep numbering 1, 2, 3...

### Same-type replace keeps the source's emphasis

`fidelity.copy_paragraph_content` copies the source's runs into the target
paragraph and restyles each one as *target base run formatting + source
emphasis* (bold, italic, underline, strike, vertical alignment, caps,
highlight, character style). `rPr` children are re-sorted into the schema's
required order, because Word rejects out-of-order run properties.

### Move → no formatting logic needed at all

Moving a block relocates the *existing* XML element to a different position
in the tree. Nothing is cloned, nothing is rewritten — so there's no
formatting question to answer here at all.

---

## 4. Operations as a tool surface

Every mutation is a small, named, deterministic Python function:
`op_replace`, `op_insert_after`, `op_delete`, `op_move`,
`op_move_to_anchor`, and the `op_table_*` family (`add_row`, `delete_row`,
`add_column`, `delete_column`, `sort`, `transpose`, `merge`). Flask routes
are thin wrappers that validate input and call these.

This split is deliberate and forward-looking: these functions are the
**tool surface an LLM agent will eventually call**. Keeping them
side-effect-contained, deterministic, and independently testable (see the
test suite built up over development) means the agent layer's job is
narrow — pick a tool, fill in its parameters — rather than needing to
understand or touch document XML itself.

### Where determinism stops on purpose

Not every table operation *can* be made deterministic without guessing, and
the code is deliberately written to refuse rather than guess:

```python
if cols_a != cols_b:
    raise ValueError(
        f"column counts differ ({cols_a} vs {cols_b}) - ambiguous how to "
        f"align them, so this needs an explicit mapping"
    )
```

Merging two tables with the same columns is unambiguous — just append rows,
and auto-detect whether the second header row is a duplicate to drop
(`headers_match`). Merging two tables with *different* columns requires a
judgment call (drop a column? pad it? map "Team" onto "Department"?) that
this code has no business making silently. The `ValueError` here isn't just
error handling — it's the designed handoff point where an LLM would step in,
resolve the ambiguity, and then call the same `op_table_merge` with an
explicit mapping (a future extension). Deterministic code does the
mechanical work; judgment calls get refused up to whoever has the context
to make them.

### Cascading operations: headings carry their subtree

Deleting, moving, or otherwise operating on a heading block generally means
"operate on this whole section," not just the one line of text. This is
implemented once and reused everywhere:

```python
def expand_with_descendants(blocks, block_ids):
    """Hierarchical ids make this a prefix match: "1.2" also pulls in
    "1.2.1", "1.2.2.1", and so on."""
```

Because ids *are* the hierarchy path, "everything under this heading" is
just a string-prefix filter — no tree-walking required. `op_delete` uses
this for cascade-delete; `op_move` and `op_move_to_anchor` use the same
`_group_span` helper so a heading always drags its subsections along when
reordered.

### Move: swapping with a sibling *group*, not a single element

The trickiest logic in the codebase is `op_move`. "Move this block up one
position" is simple for a lone paragraph, but for a heading it has to mean
"swap places with the *entire* preceding section," not just the line
directly above it (which might be the last paragraph of that section, not
its heading). `_preceding_sibling_group` / `_following_sibling_group` handle
this by walking backward/forward past whole `expand_with_descendants` groups,
using heading level to find the correct sibling boundary rather than just
looking at the adjacent line.

---

## 5. Reliability: undo and failure atomicity

### Undo: whole-document byte snapshots

Undo doesn't try to model "the opposite of an insert" or "the opposite of a
delete" — that would mean writing and maintaining an inverse for every
operation. Instead, every mutating call snapshots the *entire document's
bytes* before making any change:

```python
def snapshot(doc_entry):
    buf = io.BytesIO()
    doc_entry["document"].save(buf)
    doc_entry["undo"].append(buf.getvalue())
```

Undo just reloads the most recent snapshot as a fresh `Document` and
resyncs. It's less memory-efficient than a proper diff/patch undo stack
would be, but it's trivially correct — there's no way for an inverse
operation to be subtly wrong, because there is no inverse operation, only
"go back to a known-good copy." Capped at `MAX_UNDO_STEPS` (20) per
document.

### The `mutation()` context manager: no half-applied state

An earlier version snapshotted before an operation and resynced after,
written out by hand at each route. That had a real bug: if the operation
raised partway through (e.g. `move` rejecting a request to move the very
first block up), the snapshot had already been taken but the operation
never completed — so undo would silently do nothing (there was nothing
*wrong* to undo), while a genuinely failed request had still consumed an
undo slot.

The fix centralizes the pattern:

```python
@contextmanager
def mutation(entry):
    snapshot(entry)
    try:
        yield
    except Exception:
        entry["undo"].pop()      # discard the snapshot - nothing actually changed
        raise
    resync(entry)
```

Every mutating route now reads as `with mutation(entry): op_whatever(...)`.
If `op_whatever` raises, the snapshot is discarded and the exception
propagates up to `json_route`, which turns it into a clean `400` — the undo
stack is left exactly as it was, and the (unresynced) `blocks` view is never
touched by a failed attempt.

---

## 6. Request/response shape

Every mutating route returns the same envelope:

```json
{ "doc_id": "...", "blocks": { ... }, "can_undo": true }
```

The frontend never computes state locally — it just replaces its view of
`blocks` with whatever the server returns and re-renders. This matters
because ids can renumber after *any* operation (see §2), so a client that
tried to patch its own local copy would drift from the server almost
immediately. Treating the server's response as authoritative after every
call is what keeps the two in sync.

Route-level boilerplate (parse JSON, look up doc entries, turn a raised
`ValueError` into a `400` with a clean message) is factored into a
`@json_route` decorator plus two small helpers, `require_doc` /
`require_blocks`, so each route body is just: validate inputs → run the
operation inside `mutation()` → return `doc_state()`.

---

## 7. Known limitations (by design, for a prototype)

These are conscious scope cuts, not oversights — worth knowing before
extending the tool:

- **Blocks are headings, paragraphs, tables and images** (a paragraph that
  holds only a picture). Headers/footers, footnotes and body-level content
  controls pass through export untouched but aren't addressable.
- **Empty paragraphs aren't blocks** — they're *extras* of the block before
  them (`extract_layout`), so they move, duplicate and delete with it.
- **Same-type replace takes its base look from the target's first text
  run.** A target paragraph whose runs used several different fonts ends up
  in the first run's font (the source's emphasis is kept - see §3).
- **State is entirely in-memory** (`DOC_STORE`, a plain dict). A server
  restart loses every uploaded document. There's also no auth boundary —
  any client that knows a `doc_id` can read or mutate it.
- **No persistence layer.** For real use this needs documents on disk (or
  object storage) plus metadata in a real database, not a Python dict capped
  by a crude `MAX_DOCS` eviction.

---

## 8. Tracked-changes export (`tracked.py`)

`/export/<id>?track=1` builds a new document showing *original → current* as
Word revisions. The original bytes are kept at upload time; the export diffs
content rather than replaying history, so any sequence of edits and undos
gives a correct result:

1. Top-level elements are reduced to signatures and aligned with difflib.
2. Runs of changed elements are paired by kind and text similarity (a small
   order-preserving alignment), so an edited paragraph is matched with its
   own earlier version.
3. Paired paragraphs are diffed word-by-word over a char-level view of their
   runs - kept/inserted text uses the current runs' formatting, deleted text
   the original's - and a style change becomes `w:pPrChange`.
4. Paired tables with the same column count are diffed row-by-row, then
   cell-by-cell with the same paragraph logic.

Accept All yields the current document and Reject All the original; the test
suite checks both by simulating them on the exported XML.

## 9. Where an LLM layer would attach

Nothing in this codebase calls an LLM — every operation here is
deterministic, tested Python. The natural extension point is a thin layer
on top that:

1. Takes a natural-language instruction ("merge the two team tables,
   dropping the extra 'Notes' column from the second one").
2. Maps it to one or more of the existing `op_*` functions with concrete
   parameters — the same functions the UI's buttons already call.
3. For genuinely ambiguous cases (column mismatches, "combine these
   paragraphs sensibly"), the LLM is the thing making the judgment call;
   the deterministic layer still does the actual XML mutation and still
   gets the same formatting guarantees described in §3.

Because the operations are already isolated, parameterized, and return
structured errors on invalid input (see §4's merge example), they're
close to being directly usable as tool-call targets without modification.
