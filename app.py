"""
Doc editor prototype backend.

TWO CORE DESIGN DECISIONS
-------------------------

1. The live document is the SINGLE SOURCE OF TRUTH.
   The uploaded .docx is opened once and its python-docx Document object is
   kept alive for the session. Every operation mutates that real XML tree.
   After each mutation we RE-DERIVE the blocks/refs view from the document
   rather than patching it incrementally. That keeps ids, ordering and
   hierarchy always consistent with reality, and is what lets deletes and
   inserts renumber cleanly instead of leaving orphans or synthetic ids.

2. We NEVER rebuild a document from scratch.
   Because we mutate the existing XML elements instead of writing a fresh
   document, untouched content keeps 100% of its original formatting
   (theme fonts, colors, table borders/shading, spacing, margins), and
   edited content keeps its own paragraph/table style.

No LLM involved anywhere in this file - pure python-docx. The LLM agent
layer sits ON TOP of these operations, choosing which to call.
"""
import io
import re
import uuid
import logging
from copy import deepcopy
from contextlib import contextmanager
from functools import wraps

from flask import Flask, request, jsonify, render_template, send_file
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from bs4 import BeautifulSoup

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Logging - prints to the same console running `python app.py`.
# The HTML/table conversion path is instrumented so we can see, live, exactly
# what the parser produced and whether cell merges were applied. Look for
# lines tagged [HTML] and [TABLE].  Set DOCAGENT_DEBUG=0 to quieten it.
# ---------------------------------------------------------------------------
import os as _os
logging.basicConfig(
    level=logging.DEBUG if _os.environ.get("DOCAGENT_DEBUG", "1") != "0" else logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("docagent")

# A build stamp so you can confirm WHICH app.py is actually running. If the
# log line below doesn't appear on startup, you're running an old file/process.
BUILD_STAMP = "2026-08-15-tablemerge-v2"
log.info("docagent backend starting - BUILD %s", BUILD_STAMP)

MAX_UNDO_STEPS = 20   # per document; snapshots are whole-file bytes
MAX_DOCS = 50         # oldest docs are evicted past this (prototype memory bound)

# in-memory store: doc_id -> {"filename", "document", "refs", "blocks", "undo"}
#   document : the live python-docx Document object
#   refs     : block_id -> live Paragraph/Table object inside "document"
#   blocks   : block_id -> plain-JSON view sent to the browser
#   undo     : list of .docx byte snapshots (most recent last)
# (swap for Redis/DB + on-disk files when this outgrows a prototype)
DOC_STORE = {}


# ---------------------------------------------------------------------------
# Parsing (docx -> blocks JSON + refs to the live objects)
# ---------------------------------------------------------------------------

def iter_block_items(document):
    """
    Yield each paragraph and table in the document, IN THE ORDER they appear.
    python-docx splits document.paragraphs and document.tables into two lists
    with no shared ordering, so we walk the XML body directly instead.
    """
    for child in document.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, document)
        elif child.tag == qn('w:tbl'):
            yield Table(child, document)


def heading_level(paragraph):
    """Return 1, 2, 3... for Heading/Title styles, or None if not a heading."""
    style_name = (paragraph.style.name or "") if paragraph.style else ""
    style_name = style_name.strip()
    if style_name == "Title":
        return 1
    if style_name.startswith("Heading"):
        parts = style_name.split(" ")
        if len(parts) == 2 and parts[1].isdigit():
            return int(parts[1])
        return 1
    return None


def path_key(path):
    return ".".join(str(p) for p in path) if path else "0"


def extract_blocks_and_refs(document):
    """
    Walk the document in order and assign each block a hierarchical key based
    on heading structure ("1", "1.1", "1.2.1", ...).

    Returns (blocks, refs):
      blocks : plain JSON dict for the browser
      refs   : block_id -> the live Paragraph/Table object, so edits can be
               applied in place without losing formatting

    Called fresh after EVERY mutation, so ids always reflect the document's
    real current structure.
    """
    blocks, refs = {}, {}
    order = 0
    current_path = []
    counts = {}

    def next_path(parent_path):
        parent_key = tuple(parent_path)
        counts[parent_key] = counts.get(parent_key, 0) + 1
        return parent_path + [counts[parent_key]]

    for item in iter_block_items(document):
        order += 1

        if isinstance(item, Paragraph):
            level = heading_level(item)

            if level is not None:
                text = item.text.strip()
                if not text:
                    continue
                current_path = next_path(current_path[: level - 1])
                key = path_key(current_path)
                blocks[key] = {"id": key, "type": "header", "level": level,
                               "text": text, "order": order}
                refs[key] = item
                continue

            if not item.text.strip():
                continue
            key = path_key(next_path(current_path))
            blocks[key] = {"id": key, "type": "paragraph",
                           "text": item.text, "order": order}
            refs[key] = item

        elif isinstance(item, Table):
            key = path_key(next_path(current_path))
            blocks[key] = {
                "id": key, "type": "table", "order": order,
                "rows": [[_cell_text(cell) for cell in row.cells] for row in item.rows],
            }
            refs[key] = item

    return blocks, refs


def _cell_text(cell):
    """
    Read a table cell's text for the block model. Merged cells (from rowspan/
    colspan) make python-docx join covered paragraphs with newlines and can
    leave a trailing empty paragraph; collapse those so the editor shows clean
    values instead of stray blank lines.
    """
    text = cell.text
    # collapse internal blank lines and trim trailing/leading whitespace-only lines
    lines = [ln for ln in text.split("\n")]
    while lines and lines[-1].strip() == "":
        lines.pop()
    while lines and lines[0].strip() == "":
        lines.pop(0)
    return "\n".join(lines)


def resync(doc_entry):
    """Re-derive blocks/refs from the live document. Call after any mutation."""
    doc_entry["blocks"], doc_entry["refs"] = extract_blocks_and_refs(doc_entry["document"])


# ---------------------------------------------------------------------------
# Undo support
# ---------------------------------------------------------------------------

def snapshot(doc_entry):
    """Save the document's current bytes so the last operation can be undone."""
    buf = io.BytesIO()
    doc_entry["document"].save(buf)
    doc_entry["undo"].append(buf.getvalue())
    if len(doc_entry["undo"]) > MAX_UNDO_STEPS:
        doc_entry["undo"].pop(0)


def restore_last_snapshot(doc_entry):
    """Roll the document back one step. Returns False if nothing to undo."""
    if not doc_entry["undo"]:
        return False
    doc_entry["document"] = Document(io.BytesIO(doc_entry["undo"].pop()))
    resync(doc_entry)
    return True


@contextmanager
def mutation(entry):
    """
    Wrap a mutating operation: snapshot first, resync after, and if the
    operation raises, discard the snapshot so a failed attempt never
    consumes an undo slot or leaves the view stale.
    """
    snapshot(entry)
    try:
        yield
    except Exception:
        entry["undo"].pop()
        raise
    resync(entry)


# ---------------------------------------------------------------------------
# In-place mutation helpers (this is what preserves formatting)
# ---------------------------------------------------------------------------

def set_paragraph_text(paragraph, new_text):
    """
    Set a paragraph's text WITHOUT resetting it to a bare unformatted run.
    Reuses the first existing run (keeping its font/size/bold/color) and
    drops the rest. Per-word formatting from the SOURCE isn't carried over,
    but the TARGET's own style survives untouched.
    """
    runs = paragraph.runs
    if runs:
        runs[0].text = new_text
        for extra in runs[1:]:
            extra._element.getparent().remove(extra._element)
    else:
        paragraph.add_run(new_text)


def set_heading_level(paragraph, level):
    """Re-point a paragraph at a different Heading N style, if it exists."""
    try:
        paragraph.style = paragraph.part.document.styles[f"Heading {level}"]
    except KeyError:
        pass  # template doesn't define that level - leave the style alone


def ensure_tblgrid(table):
    """
    Guarantee the table has a <w:tblGrid> with one <w:gridCol> per column.

    Some documents - especially those generated by non-Word tools - omit
    tblGrid, or have one that's out of sync with the actual cells. python-docx
    then raises InvalidXmlError the moment you touch table.columns/table.rows.
    We rebuild it from the real cell count of the first row so every later
    operation has a consistent grid to work against.
    """
    tbl = table._tbl
    # count columns from the first row's actual cells (spans counted simply)
    first_tr = tbl.find(qn('w:tr'))
    if first_tr is None:
        return  # no rows at all - nothing to grid
    n_cols = len(first_tr.findall(qn('w:tc')))
    if n_cols == 0:
        return

    grid = tbl.find(qn('w:tblGrid'))
    if grid is None:
        # create a fresh tblGrid and insert it right after tblPr (or first)
        grid = OxmlElement('w:tblGrid')
        tblpr = tbl.find(qn('w:tblPr'))
        if tblpr is not None:
            tblpr.addnext(grid)
        else:
            tbl.insert(0, grid)

    existing = grid.findall(qn('w:gridCol'))
    if len(existing) == n_cols:
        return  # already consistent

    # rebuild gridCol children to match the real column count, reusing any
    # existing width so we don't lose column sizing where it was present
    default_w = existing[0].get(qn('w:w')) if existing else None
    for gc in existing:
        grid.remove(gc)
    for _ in range(n_cols):
        gc = OxmlElement('w:gridCol')
        if default_w:
            gc.set(qn('w:w'), default_w)
        grid.append(gc)


def resize_table_columns(table, target_cols):
    """Add/remove columns in place by cloning the last column's XML, so
    borders/shading/width carry over instead of appearing unstyled."""
    ensure_tblgrid(table)
    current_cols = len(table.columns)
    if target_cols == current_cols:
        return

    grid = table._tbl.find(qn('w:tblGrid'))

    if target_cols > current_cols:
        grid_cols = grid.findall(qn('w:gridCol')) if grid is not None else []
        last_gridcol = grid_cols[-1] if grid_cols else None
        for _ in range(target_cols - current_cols):
            if grid is not None and last_gridcol is not None:
                grid.append(deepcopy(last_gridcol))
            for row in table.rows:
                tcs = row._tr.findall(qn('w:tc'))
                if not tcs:
                    continue
                new_tc = deepcopy(tcs[-1])
                for t in new_tc.findall('.//' + qn('w:t')):
                    t.text = ''
                row._tr.append(new_tc)
    else:
        diff = current_cols - target_cols
        if grid is not None:
            for gc in grid.findall(qn('w:gridCol'))[-diff:]:
                grid.remove(gc)
        for row in table.rows:
            for tc in row._tr.findall(qn('w:tc'))[-diff:]:
                row._tr.remove(tc)


def resize_table_rows(table, target_rows):
    current_rows = len(table.rows)
    if target_rows > current_rows:
        for _ in range(target_rows - current_rows):
            table.add_row()
    elif target_rows < current_rows:
        for row in list(table.rows)[target_rows:]:
            table._tbl.remove(row._tr)


def set_table_content(table, new_rows):
    """Resize the live table in place then fill cell text - the table's
    style, borders and shading stay intact throughout."""
    if not new_rows:
        return
    target_cols = max(len(r) for r in new_rows)
    resize_table_columns(table, target_cols)
    resize_table_rows(table, len(new_rows))

    for r_idx, row_data in enumerate(new_rows):
        cells = table.rows[r_idx].cells
        for c_idx in range(target_cols):
            text = row_data[c_idx] if c_idx < len(row_data) else ""
            set_paragraph_text(cells[c_idx].paragraphs[0], text)


# ---------------------------------------------------------------------------
# Operations
# ---------------------------------------------------------------------------

def op_replace(base_entry, base_block_id, source_entry, source_block_id):
    """
    Overwrite a base block's CONTENT with a source block's content, keeping
    the base block's own position and formatting.

    Same-type replaces (paragraph<->header, table<->table) mutate the target
    in place. A cross-type replace (e.g. table content into a paragraph slot)
    can't be done in place, so we clone the source element and swap it into
    the target's position instead - which keeps the SOURCE's formatting.
    """
    target_ref = base_entry["refs"][base_block_id]
    source_ref = source_entry["refs"][source_block_id]
    source_block = source_entry["blocks"][source_block_id]
    target_block = base_entry["blocks"][base_block_id]

    same_kind = (
        isinstance(target_ref, Table) and isinstance(source_ref, Table)
    ) or (
        isinstance(target_ref, Paragraph) and isinstance(source_ref, Paragraph)
    )

    if same_kind and isinstance(target_ref, Table):
        set_table_content(target_ref, source_block["rows"])

    elif same_kind:
        set_paragraph_text(target_ref, source_block["text"])
        # a header target keeps being a header; adopt the source's level if
        # the source is also a header, otherwise keep the target's own level
        if target_block["type"] == "header":
            level = source_block.get("level", target_block.get("level", 1)) \
                if source_block["type"] == "header" else target_block.get("level", 1)
            set_heading_level(target_ref, level)
        elif source_block["type"] == "header":
            set_heading_level(target_ref, source_block.get("level", 1))

    else:
        # cross-type: swap the element wholesale, preserving source formatting
        target_element = target_ref._element
        parent = target_element.getparent()
        clone = deepcopy(source_ref._element)
        parent.insert(parent.index(target_element), clone)
        parent.remove(target_element)


def expand_with_descendants(blocks, block_ids):
    """
    Expand a selection so that operating on a heading also covers everything
    nested under it. Hierarchical ids make this a prefix match: "1.2" also
    pulls in "1.2.1", "1.2.2.1", and so on.
    """
    expanded = set()
    for block_id in block_ids:
        expanded.add(block_id)
        if blocks.get(block_id, {}).get("type") != "header":
            continue
        prefix = block_id + "."
        expanded.update(bid for bid in blocks if bid.startswith(prefix))
    return expanded


def ordered_elements(entry, block_ids):
    """XML elements for the given blocks, in document order."""
    ids = sorted(block_ids, key=lambda b: entry["blocks"][b]["order"])
    return [entry["refs"][b]._element for b in ids], ids


# --- Move / reorder ---------------------------------------------------------

def _sorted_block_ids(blocks):
    return sorted(blocks, key=lambda b: blocks[b]["order"])


def _group_span(blocks, block_id):
    """
    The set of block ids that move together with block_id: a heading takes
    its whole section, anything else moves alone.
    """
    return expand_with_descendants(blocks, [block_id])


def _preceding_sibling_group(blocks, block_id):
    """
    The group immediately BEFORE this one that we'd swap with when moving up.

    For a heading we skip over the entire preceding sibling section (so
    sections reorder as units); for a paragraph/table we step over just the
    one preceding block.
    """
    ordered = _sorted_block_ids(blocks)
    group = _group_span(blocks, block_id)
    first_idx = min(ordered.index(b) for b in group)
    if first_idx == 0:
        return None

    prev_id = ordered[first_idx - 1]
    block = blocks[block_id]

    if block["type"] != "header":
        return _group_span(blocks, prev_id)

    # walk back to the nearest preceding heading at the same level or higher
    level = block["level"]
    for idx in range(first_idx - 1, -1, -1):
        candidate = blocks[ordered[idx]]
        if candidate["type"] == "header" and candidate["level"] <= level:
            return _group_span(blocks, ordered[idx])
    # no sibling heading before it - move above everything that precedes it
    return set(ordered[:first_idx])


def _following_sibling_group(blocks, block_id):
    """The group immediately AFTER this one, for moving down."""
    ordered = _sorted_block_ids(blocks)
    group = _group_span(blocks, block_id)
    last_idx = max(ordered.index(b) for b in group)
    if last_idx >= len(ordered) - 1:
        return None

    next_id = ordered[last_idx + 1]
    block = blocks[block_id]

    if block["type"] != "header":
        return _group_span(blocks, next_id)

    # the next block after a section is by definition the next sibling
    # heading (or a stray block at a shallower level) - move past its group
    return _group_span(blocks, next_id)


def op_move(entry, block_id, direction):
    """
    Move a block (and its section, if it's a heading) one position up or
    down, swapping with the adjacent sibling group. Pure XML relocation -
    the elements themselves are untouched, so all formatting is preserved.
    """
    blocks = entry["blocks"]
    if block_id not in blocks:
        raise ValueError(f"unknown block_id: {block_id}")

    group = _group_span(blocks, block_id)
    elements, _ = ordered_elements(entry, group)
    if not elements:
        return False

    parent = elements[0].getparent()

    if direction == "up":
        neighbour = _preceding_sibling_group(blocks, block_id)
        if not neighbour:
            return False
        n_elements, _ = ordered_elements(entry, neighbour)
        target = n_elements[0]           # insert our group before it
        index = parent.index(target)
        for offset, element in enumerate(elements):
            parent.remove(element)
            parent.insert(index + offset, element)

    elif direction == "down":
        neighbour = _following_sibling_group(blocks, block_id)
        if not neighbour:
            return False
        n_elements, _ = ordered_elements(entry, neighbour)
        cursor = n_elements[-1]          # insert our group after it
        for element in elements:
            parent.remove(element)
            parent.insert(parent.index(cursor) + 1, element)
            cursor = element
    else:
        raise ValueError("direction must be 'up' or 'down'")

    return True


def op_move_to_anchor(entry, block_ids, anchor_block_id, position="after"):
    """
    Relocate blocks (headings bring their sections) to sit before/after an
    anchor block, within the same document.
    """
    blocks = entry["blocks"]
    group = expand_with_descendants(blocks, block_ids)
    if anchor_block_id in group:
        raise ValueError("anchor cannot be inside the blocks being moved")

    elements, _ = ordered_elements(entry, group)
    if not elements:
        return False

    # for an "after" move onto a heading, land past its whole section
    anchor_group = _group_span(blocks, anchor_block_id) if position == "after" \
        else {anchor_block_id}
    anchor_elements, _ = ordered_elements(entry, anchor_group)
    parent = elements[0].getparent()

    if position == "after":
        cursor = anchor_elements[-1]
        for element in elements:
            parent.remove(element)
            parent.insert(parent.index(cursor) + 1, element)
            cursor = element
    else:
        target = anchor_elements[0]
        index = parent.index(target)
        for offset, element in enumerate(elements):
            parent.remove(element)
            parent.insert(index + offset, element)

    return True


# --- Deterministic table operations -----------------------------------------

def _require_table(entry, block_id):
    block = entry["blocks"].get(block_id)
    if block is None:
        raise ValueError(f"unknown block_id: {block_id}")
    if block["type"] != "table":
        raise ValueError(f"block {block_id} is not a table")
    return entry["refs"][block_id], block


def op_table_add_row(entry, block_id, index=None, values=None):
    """Insert a row, cloning an existing row's XML so borders/shading carry."""
    table, block = _require_table(entry, block_id)
    rows = [r[:] for r in block["rows"]]
    n_cols = len(rows[0]) if rows else 1
    index = len(rows) if index is None else max(0, min(index, len(rows)))

    new_row = (values or [])[:n_cols]
    new_row += [""] * (n_cols - len(new_row))
    rows.insert(index, new_row)
    set_table_content(table, rows)


def op_table_delete_row(entry, block_id, index):
    table, block = _require_table(entry, block_id)
    rows = [r[:] for r in block["rows"]]
    if not 0 <= index < len(rows):
        raise ValueError(f"row {index} out of range (0-{len(rows)-1})")
    if len(rows) == 1:
        raise ValueError("cannot delete the only row - delete the table instead")
    rows.pop(index)
    set_table_content(table, rows)


def op_table_add_column(entry, block_id, index=None, header="", values=None):
    """Add a column; `header` fills row 0, `values` fills the rest."""
    table, block = _require_table(entry, block_id)
    rows = [r[:] for r in block["rows"]]
    n_cols = len(rows[0]) if rows else 0
    index = n_cols if index is None else max(0, min(index, n_cols))

    values = values or []
    for r_idx, row in enumerate(rows):
        if r_idx == 0:
            cell = header
        else:
            cell = values[r_idx - 1] if r_idx - 1 < len(values) else ""
        row.insert(index, cell)
    set_table_content(table, rows)


def op_table_delete_column(entry, block_id, index):
    table, block = _require_table(entry, block_id)
    rows = [r[:] for r in block["rows"]]
    n_cols = len(rows[0]) if rows else 0
    if not 0 <= index < n_cols:
        raise ValueError(f"column {index} out of range (0-{n_cols-1})")
    if n_cols == 1:
        raise ValueError("cannot delete the only column - delete the table instead")
    for row in rows:
        if index < len(row):
            row.pop(index)
    set_table_content(table, rows)


def op_table_sort(entry, block_id, column, descending=False, has_header=True):
    """Sort data rows by a column - numerically when the whole column parses
    as numbers, otherwise case-insensitive text."""
    table, block = _require_table(entry, block_id)
    rows = [r[:] for r in block["rows"]]
    head = rows[:1] if has_header else []
    body = rows[1:] if has_header else rows

    n_cols = len(rows[0]) if rows else 0
    if not 0 <= column < n_cols:
        raise ValueError(f"column {column} out of range (0-{n_cols-1})")

    def cell(row):
        return row[column] if column < len(row) else ""

    def as_number(value):
        try:
            return float(value.strip().replace(",", "").lstrip("$"))
        except (ValueError, AttributeError):
            return None

    numeric = [as_number(cell(r)) for r in body]
    if body and all(n is not None for n in numeric):
        body.sort(key=lambda r: as_number(cell(r)), reverse=descending)
    else:
        body.sort(key=lambda r: cell(r).strip().lower(), reverse=descending)

    set_table_content(table, head + body)


def op_table_transpose(entry, block_id):
    table, block = _require_table(entry, block_id)
    rows = block["rows"]
    if not rows:
        return
    n_cols = max(len(r) for r in rows)
    transposed = [
        [row[c] if c < len(row) else "" for row in rows]
        for c in range(n_cols)
    ]
    set_table_content(table, transposed)


def headers_match(rows_a, rows_b):
    """Do two tables share the same first row (ignoring case/whitespace)?"""
    if not rows_a or not rows_b:
        return False
    norm = lambda row: [c.strip().lower() for c in row]
    return norm(rows_a[0]) == norm(rows_b[0])


def op_table_merge(entry, target_id, other_id, skip_header=None, remove_other=True):
    """
    Append one table's rows onto another, in place, keeping the TARGET's
    formatting, then remove the now-redundant source table (set
    `remove_other=False` to keep it).

    `skip_header` defaults to auto: drop the second table's first row only
    when both tables share the same header row.

    Column counts must match - a mismatch is genuinely ambiguous (drop a
    column? pad it? map it to a differently-named one?) and is exactly the
    kind of call to hand to the LLM layer rather than guess at here.
    """
    target, target_block = _require_table(entry, target_id)
    other_ref, other_block = _require_table(entry, other_id)

    rows_a = [r[:] for r in target_block["rows"]]
    rows_b = [r[:] for r in other_block["rows"]]
    if not rows_a or not rows_b:
        raise ValueError("cannot merge an empty table")

    cols_a = max(len(r) for r in rows_a)
    cols_b = max(len(r) for r in rows_b)
    if cols_a != cols_b:
        raise ValueError(
            f"column counts differ ({cols_a} vs {cols_b}) - ambiguous how to "
            f"align them, so this needs an explicit mapping"
        )

    if skip_header is None:
        skip_header = headers_match(rows_a, rows_b)

    set_table_content(target, rows_a + (rows_b[1:] if skip_header else rows_b))

    if remove_other:
        element = other_ref._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)

    return {
        "skipped_header": skip_header,
        "rows_added": len(rows_b) - (1 if skip_header else 0),
        "removed_source_table": remove_other,
    }


def op_table_set_content(entry, block_id, rows):
    """
    Overwrite a table's ENTIRE content with an explicit row matrix - the
    backing operation for the table editor popup. Rows are padded to a
    rectangular shape (ragged input from the UI is tolerated, not rejected).
    """
    table, _ = _require_table(entry, block_id)
    if not isinstance(rows, list) or not rows:
        raise ValueError("rows must be a non-empty list of lists")
    clean = [[str(c) for c in row] for row in rows]
    width = max(len(r) for r in clean)
    clean = [r + [""] * (width - len(r)) for r in clean]
    set_table_content(table, clean)


def op_edit_text(entry, block_id, text):
    """Overwrite a paragraph/header block's text in place - the backing
    operation for inline text editing."""
    block = entry["blocks"].get(block_id)
    if block is None:
        raise ValueError(f"unknown block_id: {block_id}")
    if block["type"] not in ("header", "paragraph"):
        raise ValueError(f"block {block_id} is not editable text (it's a {block['type']})")
    set_paragraph_text(entry["refs"][block_id], text)


def op_delete(doc_entry, block_ids, cascade=True):
    """Remove blocks' XML elements from the live document."""
    targets = expand_with_descendants(doc_entry["blocks"], block_ids) if cascade \
        else set(block_ids)

    removed = []
    for block_id in targets:
        ref = doc_entry["refs"].get(block_id)
        if ref is None:
            continue
        element = ref._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)
            removed.append(block_id)
    return removed


def op_insert_after(base_entry, anchor_block_id, source_entry, source_block_ids):
    """
    Clone selected source blocks (keeping the SOURCE doc's formatting) and
    splice them into the base document after the anchor, in the order they
    appear in the source document.
    """
    anchor_element = base_entry["refs"][anchor_block_id]._element
    parent = anchor_element.getparent()

    ordered_ids = sorted(
        source_block_ids, key=lambda bid: source_entry["blocks"][bid]["order"]
    )

    cursor = anchor_element
    for source_block_id in ordered_ids:
        clone = deepcopy(source_entry["refs"][source_block_id]._element)
        parent.insert(parent.index(cursor) + 1, clone)
        cursor = clone

    return ordered_ids


def op_duplicate(entry, block_ids):
    """
    Clone selected blocks (and, for a heading, its whole section) and insert
    the copies immediately after the original selection - within the same
    document. Cloning the XML keeps all formatting. Returns nothing; the caller
    resyncs to get fresh ids.
    """
    # expand headings to their whole section, then take document order
    targets = expand_with_descendants(entry["blocks"], block_ids)
    ordered_ids = sorted(targets, key=lambda b: entry["blocks"][b]["order"])
    if not ordered_ids:
        return

    # insert the clones after the LAST element of the selection, preserving
    # the selection's internal order
    last_element = entry["refs"][ordered_ids[-1]]._element
    parent = last_element.getparent()

    cursor = last_element
    for block_id in ordered_ids:
        clone = deepcopy(entry["refs"][block_id]._element)
        parent.insert(parent.index(cursor) + 1, clone)
        cursor = clone


def _replace_in_paragraph(paragraph, find, replace, case_sensitive):
    """
    Replace text within a paragraph while keeping formatting where possible.
    Works run-by-run first (so a match inside one run keeps that run's style);
    falls back to a whole-paragraph rewrite when a match spans runs. Returns
    the number of replacements made.
    """
    flags = 0 if case_sensitive else re.IGNORECASE
    pattern = re.compile(re.escape(find), flags)

    # fast path: try per-run replacement (preserves each run's own formatting)
    made = 0
    spans_runs = False
    for run in paragraph.runs:
        if pattern.search(run.text):
            run.text, n = pattern.subn(replace, run.text)
            made += n
    # detect a match that spans runs: reconstruct full text and compare counts
    full = paragraph.text
    total_in_full = len(pattern.findall(full))
    if total_in_full > made:
        # a match crosses run boundaries - rewrite the whole paragraph text on
        # the first run (loses per-word styling for this paragraph only)
        spans_runs = True
        new_full = pattern.sub(replace, full)
        if paragraph.runs:
            paragraph.runs[0].text = new_full
            for extra in paragraph.runs[1:]:
                extra._element.getparent().remove(extra._element)
        else:
            paragraph.add_run(new_full)
        made = total_in_full
    return made


def op_find_replace(entry, find, replace, case_sensitive=False, scope="all"):
    """
    Find & replace across the whole document, in paragraphs/headers and table
    cells. `scope` can be 'all', 'text' (paragraphs+headers only), or 'tables'
    (table cells only). Returns the total number of replacements.
    """
    if not find:
        raise ValueError("nothing to find")

    total = 0
    for item in iter_block_items(entry["document"]):
        if isinstance(item, Paragraph):
            if scope in ("all", "text"):
                total += _replace_in_paragraph(item, find, replace, case_sensitive)
        elif isinstance(item, Table):
            if scope in ("all", "tables"):
                for row in item.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            total += _replace_in_paragraph(p, find, replace, case_sensitive)
    return total


# ---------------------------------------------------------------------------
# Route plumbing
# ---------------------------------------------------------------------------

def doc_state(doc_id):
    """Standard success payload - the browser re-renders from this."""
    entry = DOC_STORE[doc_id]
    return {
        "doc_id": doc_id,
        "blocks": entry["blocks"],
        "can_undo": bool(entry["undo"]),
    }


def json_route(fn):
    """Parse JSON body, resolve doc entries, and turn ValueError into a 400."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(request.get_json(force=True) or {}, *args, **kwargs)
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
    return wrapper


def require_doc(doc_id):
    if doc_id not in DOC_STORE:
        raise ValueError("unknown doc_id")
    return DOC_STORE[doc_id]


def require_blocks(entry, block_ids, label):
    missing = [b for b in block_ids if b not in entry["blocks"]]
    if missing:
        raise ValueError(f"unknown {label}: {missing}")


# ---------------------------------------------------------------------------
# HTML -> python-docx Document (for Froala / pasted HTML input)
# ---------------------------------------------------------------------------
#
# Design note: instead of a separate "HTML block model", we convert incoming
# HTML into a REAL python-docx Document and feed it through the exact same
# DOC_STORE path as an uploaded .docx. HTML input therefore reuses everything
# already built - parse-to-blocks, editing, reordering, table ops, merge,
# export - with no changes to any of it. The only new code is this converter.
#
# Fidelity is intentionally scoped to structure + basic inline emphasis:
# headings, paragraphs, captions, lists, tables, with bold/italic/underline.
# Froala's inline styles (fonts, colors, pixel spacing) are not carried over -
# see DESIGN.md for why HTML->docx can't preserve formatting like docx->docx.

_HTML_BLOCK_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6", "p", "table",
                    "ul", "ol", "caption", "figcaption", "blockquote", "div",
                    "section", "article", "figure"}


def _iter_inline(node, bold=False, italic=False, underline=False):
    """Yield (text, bold, italic, underline) runs from inline children.
    Nested emphasis compounds, e.g. <b><i>x</i></b> -> bold+italic."""
    from bs4 import NavigableString, Tag
    for child in node.children:
        if isinstance(child, NavigableString):
            text = str(child)
            if text:
                yield (text, bold, italic, underline)
        elif isinstance(child, Tag):
            name = child.name.lower()
            b = bold or name in ("b", "strong")
            i = italic or name in ("i", "em")
            u = underline or name == "u"
            if name == "br":
                yield ("\n", bold, italic, underline)
            else:
                yield from _iter_inline(child, b, i, u)


def _add_inline_runs(paragraph, node):
    """Add runs to a paragraph from an element's inline content, collapsing
    whitespace but preserving explicit <br> newlines. Returns True if any
    non-empty text was added."""
    any_text = False
    for text, b, i, u in _iter_inline(node):
        text = re.sub(r"[ \t\r\f\v]+", " ", text)
        if not text.strip() and text != "\n":
            continue
        run = paragraph.add_run(text)
        run.bold = b or None
        run.italic = i or None
        run.underline = u or None
        any_text = True
    return any_text


def _heading_level_from_tag(name):
    m = re.fullmatch(r"h([1-6])", name.lower())
    return int(m.group(1)) if m else None


def _looks_like_caption(tag):
    if tag.name.lower() in ("caption", "figcaption"):
        return True
    classes = " ".join(tag.get("class", [])).lower()
    return "caption" in classes


def _html_table_to_rows(table_tag):
    """
    Parse an HTML <table> honoring BOTH colspan and rowspan.

    Returns (matrix, merges):
      matrix - a rectangular list-of-lists of cell text. A spanned cell's text
               is placed in its top-left "anchor" position; the other positions
               it covers are left empty (so the value isn't duplicated).
      merges - a list of (top, left, bottom, right) inclusive grid regions that
               should be merged into a single cell. Building code turns these
               into real merged cells in the Word table, so a rowspan/colspan in
               the source HTML becomes an actual merged cell in the output -
               matching what the user drew in Froala.

    Naive row-by-row extraction breaks on rowspan: a cell that spans down leaves
    the rows beneath it short, shifting every later column and producing a ragged
    grid. We place each cell into an explicit grid, reserving the positions its
    spans occupy so following cells flow into the right slots.
    """
    grid = {}          # (r, c) -> text (anchor cell only)
    occupied = set()    # (r, c) covered by some span (incl. anchors)
    merges = []
    max_col = 0

    trs = table_tag.find_all("tr")
    for r, tr in enumerate(trs):
        c = 0
        for cell in tr.find_all(["th", "td"]):
            # skip columns already taken by a span from an earlier cell/row
            while (r, c) in occupied:
                c += 1

            text = cell.get_text(" ", strip=True)
            try:
                colspan = max(1, int(cell.get("colspan", 1)))
            except (TypeError, ValueError):
                colspan = 1
            try:
                rowspan = max(1, int(cell.get("rowspan", 1)))
            except (TypeError, ValueError):
                rowspan = 1

            # anchor gets the text; reserve every covered position
            grid[(r, c)] = text
            for dr in range(rowspan):
                for dc in range(colspan):
                    occupied.add((r + dr, c + dc))

            if rowspan > 1 or colspan > 1:
                merges.append((r, c, r + rowspan - 1, c + colspan - 1))
                log.debug("[TABLE] cell (r=%s,c=%s) text=%r rowspan=%s colspan=%s -> merge region %s",
                          r, c, text[:30], rowspan, colspan, (r, c, r + rowspan - 1, c + colspan - 1))

            c += colspan
            max_col = max(max_col, c)

    n_rows = len(trs)
    if n_rows == 0 or max_col == 0:
        return [], []

    matrix = [[grid.get((r, c), "") for c in range(max_col)] for r in range(n_rows)]

    # drop fully-empty trailing rows a spanning layout can leave behind, and
    # any merges that referenced only those rows
    while matrix and all(cell == "" for cell in matrix[-1]):
        matrix.pop()
    n_rows = len(matrix)
    merges = [m for m in merges if m[0] < n_rows]
    # clamp any merge that ran past the trimmed bottom
    merges = [(t, l, min(b, n_rows - 1), r) for (t, l, b, r) in merges]

    log.debug("[TABLE] parsed %s rows x %s cols, %s merge region(s): %s",
              len(matrix), max_col, len(merges), merges)
    return matrix, merges


def html_to_document(html):
    """
    Convert an HTML fragment (Froala output) into a python-docx Document:
      h1..h6   -> headings
      p / div  -> paragraphs (bold/italic/underline runs preserved)
      caption  -> italic caption paragraph
      ul / ol  -> List Bullet / List Number paragraphs
      table    -> a real Word table (Table Grid)
    Raises ValueError if nothing usable is found.
    """
    soup = BeautifulSoup(html or "", "lxml")
    body = soup.body or soup
    document = Document()
    produced = 0

    def emit(tag):
        nonlocal produced
        name = tag.name.lower()

        level = _heading_level_from_tag(name)
        if level is not None:
            text = tag.get_text(" ", strip=True)
            if text:
                document.add_heading(text, level=min(level, 4))
                produced += 1
            return

        if name == "table":
            rows, merges = _html_table_to_rows(tag)
            if rows:
                n_cols = max(len(r) for r in rows)
                log.debug("[HTML] building Word table %s x %s with %s merge(s)",
                          len(rows), n_cols, len(merges))
                t = document.add_table(rows=len(rows), cols=n_cols)
                t.style = "Table Grid"
                for r_idx, row in enumerate(rows):
                    for c_idx, val in enumerate(row):
                        t.cell(r_idx, c_idx).text = val
                # apply real cell merges for rowspan/colspan regions
                applied = 0
                for (top, left, bottom, right) in merges:
                    try:
                        a = t.cell(top, left)
                        b = t.cell(bottom, right)
                        if a is not b:
                            a.merge(b)
                            applied += 1
                    except Exception as merge_err:
                        # never let an odd span break the whole conversion;
                        # worst case the cells stay unmerged but present
                        log.warning("[HTML] merge %s failed: %s",
                                    (top, left, bottom, right), merge_err)
                log.debug("[HTML] applied %s/%s merges on this table", applied, len(merges))
                produced += 1
            return

        if name in ("ul", "ol"):
            style = "List Number" if name == "ol" else "List Bullet"
            for li in tag.find_all("li", recursive=False) or tag.find_all("li"):
                p = document.add_paragraph(style=style)
                if _add_inline_runs(p, li):
                    produced += 1
                else:
                    p._element.getparent().remove(p._element)
            return

        if _looks_like_caption(tag):
            text = tag.get_text(" ", strip=True)
            if text:
                p = document.add_paragraph()
                p.add_run(text).italic = True
                produced += 1
            return

        if name == "blockquote":
            text = tag.get_text(" ", strip=True)
            if text:
                document.add_paragraph(text)
                produced += 1
            return

        if name in ("p", "div", "section", "article", "figure"):
            child_blocks = [c for c in tag.find_all(recursive=False)
                            if getattr(c, "name", None) and c.name.lower() in _HTML_BLOCK_TAGS]
            if child_blocks:
                for c in child_blocks:
                    emit(c)
                return
            p = document.add_paragraph()
            if _add_inline_runs(p, tag):
                produced += 1
            else:
                p._element.getparent().remove(p._element)
            return

    top = [c for c in body.find_all(recursive=False) if getattr(c, "name", None)]
    if not top:
        text = body.get_text(" ", strip=True)
        if text:
            document.add_paragraph(text)
            produced += 1
    else:
        for tag in top:
            if tag.name.lower() in _HTML_BLOCK_TAGS:
                emit(tag)
            else:
                for c in tag.find_all(recursive=False):
                    if getattr(c, "name", None) and c.name.lower() in _HTML_BLOCK_TAGS:
                        emit(c)

    if produced == 0:
        raise ValueError("no headings, paragraphs, or tables found in the HTML")
    return document


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.errorhandler(Exception)
def handle_uncaught(e):
    """
    Any exception NOT already turned into a clean 400 by @json_route (e.g. a
    python-docx InvalidXmlError from a malformed uploaded file) still gets a
    JSON response instead of Flask's HTML debug page. An HTML response where
    the frontend expects JSON either throws confusingly or, worse, leaves a
    stuck "loading" state - this keeps failures visible and recoverable.
    """
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e   # keep normal 404s etc. as-is
    import traceback
    traceback.print_exc()   # full trace still lands in the server console
    return jsonify({"error": f"{type(e).__name__}: {e}"}), 500


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "no file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith(".docx"):
        return jsonify({"error": "only .docx files are supported"}), 400

    try:
        document = Document(io.BytesIO(f.read()))
    except Exception:
        return jsonify({"error": "could not read that file - is it a valid .docx?"}), 400

    doc_id = _store_document(document, f.filename)
    payload = doc_state(doc_id)
    payload["filename"] = f.filename
    return jsonify(payload)


def _store_document(document, filename):
    """Register a python-docx Document in DOC_STORE and return its id."""
    doc_id = str(uuid.uuid4())[:8]
    entry = {"filename": filename, "document": document, "undo": []}
    resync(entry)
    DOC_STORE[doc_id] = entry
    # crude memory bound: evict oldest docs (dicts keep insertion order)
    while len(DOC_STORE) > MAX_DOCS:
        DOC_STORE.pop(next(iter(DOC_STORE)))
    return doc_id


def append_document_body(target_entry, source_document):
    """
    Append every top-level block (paragraph/table) of source_document to the
    end of target_entry's live document, cloning the XML so the source's
    formatting comes along. Used to merge HTML content straight into a base doc.
    """
    target_body = target_entry["document"].element.body
    # insert before sectPr (the trailing section properties), if present
    sectpr = target_body.find(qn('w:sectPr'))
    for child in list(source_document.element.body.iterchildren()):
        if child.tag == qn('w:sectPr'):
            continue
        clone = deepcopy(child)
        if sectpr is not None:
            sectpr.addprevious(clone)
        else:
            target_body.append(clone)


@app.route("/upload_html", methods=["POST"])
@json_route
def upload_html(data):
    """
    Accept an HTML fragment (e.g. from Froala) and either:
      - load it as a NEW document (default), returned like /upload, so it can
        be pulled into the base doc via the normal copy/insert flow; or
      - append it directly to an existing base document, when
        {append_to: <doc_id>} is provided.

    The 'destination' is chosen by the caller, matching "let me choose where
    it goes".
    """
    html = data.get("html", "")
    if not html or not html.strip():
        raise ValueError("no HTML provided")

    log.info("[HTML] /upload_html received %s chars (append_to=%s)",
             len(html), data.get("append_to"))

    try:
        document = html_to_document(html)
    except ValueError as e:
        raise ValueError(str(e))

    append_to = data.get("append_to")
    if append_to:
        target = require_doc(append_to)
        with mutation(target):
            append_document_body(target, document)
        payload = doc_state(append_to)
        payload["appended"] = True
        _log_table_blocks("append_to base", payload)
        return jsonify(payload)

    # otherwise store as a new document (lands in whichever pane the UI asked)
    name = data.get("name") or "pasted.html"
    doc_id = _store_document(document, name)
    payload = doc_state(doc_id)
    payload["filename"] = name
    _log_table_blocks("new doc", payload)
    return jsonify(payload)


def _log_table_blocks(where, payload):
    """Log the final table rows the browser will receive, so we can see whether
    the block model the editor renders still shows repeated span values."""
    for bid, blk in payload.get("blocks", {}).items():
        if blk.get("type") == "table":
            log.info("[HTML] table block %s (%s) rows -> %s", bid, where, blk["rows"])


@app.route("/replace", methods=["POST"])
@json_route
def replace(data):
    base_entry = require_doc(data.get("base_doc_id"))
    source_entry = require_doc(data.get("source_doc_id"))
    base_block_id = data.get("base_block_id")
    source_block_id = data.get("source_block_id")

    require_blocks(base_entry, [base_block_id], "base_block_id")
    require_blocks(source_entry, [source_block_id], "source_block_id")

    with mutation(base_entry):
        op_replace(base_entry, base_block_id, source_entry, source_block_id)

    return jsonify(doc_state(data["base_doc_id"]))


@app.route("/insert", methods=["POST"])
@json_route
def insert(data):
    base_entry = require_doc(data.get("base_doc_id"))
    source_entry = require_doc(data.get("source_doc_id"))
    anchor_block_id = data.get("anchor_block_id")
    source_block_ids = data.get("source_block_ids", [])

    if not source_block_ids:
        raise ValueError("no source_block_ids provided")
    require_blocks(base_entry, [anchor_block_id], "anchor_block_id")
    require_blocks(source_entry, source_block_ids, "source_block_ids")

    with mutation(base_entry):
        op_insert_after(base_entry, anchor_block_id, source_entry, source_block_ids)

    return jsonify(doc_state(data["base_doc_id"]))


@app.route("/delete", methods=["POST"])
@json_route
def delete(data):
    doc_id = data.get("doc_id")
    entry = require_doc(doc_id)
    block_ids = data.get("block_ids", [])
    cascade = data.get("cascade", True)

    if not block_ids:
        raise ValueError("no block_ids provided")
    require_blocks(entry, block_ids, "block_ids")

    with mutation(entry):
        removed = op_delete(entry, block_ids, cascade=cascade)

    payload = doc_state(doc_id)
    payload["removed"] = removed
    return jsonify(payload)


@app.route("/move", methods=["POST"])
@json_route
def move(data):
    doc_id = data.get("doc_id")
    entry = require_doc(doc_id)
    block_id = data.get("block_id")
    direction = data.get("direction")

    require_blocks(entry, [block_id], "block_id")

    with mutation(entry):
        if not op_move(entry, block_id, direction):
            raise ValueError("already at the edge of the document")

    return jsonify(doc_state(doc_id))


@app.route("/duplicate", methods=["POST"])
@json_route
def duplicate(data):
    doc_id = data.get("doc_id")
    entry = require_doc(doc_id)
    block_ids = data.get("block_ids", [])

    if not block_ids:
        raise ValueError("no block_ids provided")
    require_blocks(entry, block_ids, "block_ids")

    with mutation(entry):
        op_duplicate(entry, block_ids)

    return jsonify(doc_state(doc_id))


@app.route("/find_replace", methods=["POST"])
@json_route
def find_replace(data):
    doc_id = data.get("doc_id")
    entry = require_doc(doc_id)
    find = data.get("find", "")
    replace = data.get("replace", "")
    case_sensitive = bool(data.get("case_sensitive", False))
    scope = data.get("scope", "all")

    if not find:
        raise ValueError("nothing to find")
    if scope not in ("all", "text", "tables"):
        raise ValueError("scope must be 'all', 'text', or 'tables'")

    with mutation(entry):
        count = op_find_replace(entry, find, replace, case_sensitive, scope)

    payload = doc_state(doc_id)
    payload["replacements"] = count
    return jsonify(payload)


@app.route("/move_to", methods=["POST"])
@json_route
def move_to(data):
    doc_id = data.get("doc_id")
    entry = require_doc(doc_id)
    block_ids = data.get("block_ids", [])
    anchor_block_id = data.get("anchor_block_id")
    position = data.get("position", "after")

    if not block_ids:
        raise ValueError("no block_ids provided")
    if position not in ("before", "after"):
        raise ValueError("position must be 'before' or 'after'")
    require_blocks(entry, block_ids, "block_ids")
    require_blocks(entry, [anchor_block_id], "anchor_block_id")

    with mutation(entry):
        op_move_to_anchor(entry, block_ids, anchor_block_id, position)

    return jsonify(doc_state(doc_id))


# --- Table operations -------------------------------------------------------

TABLE_OPS = {
    "add_row":       lambda e, b, p: op_table_add_row(e, b, p.get("index"), p.get("values")),
    "delete_row":    lambda e, b, p: op_table_delete_row(e, b, p["index"]),
    "add_column":    lambda e, b, p: op_table_add_column(e, b, p.get("index"), p.get("header", ""), p.get("values")),
    "delete_column": lambda e, b, p: op_table_delete_column(e, b, p["index"]),
    "sort":          lambda e, b, p: op_table_sort(e, b, p["column"], p.get("descending", False), p.get("has_header", True)),
    "transpose":     lambda e, b, p: op_table_transpose(e, b),
    "merge":         lambda e, b, p: op_table_merge(e, b, p["other_block_id"], p.get("skip_header"), p.get("remove_other", True)),
    "set_content":   lambda e, b, p: op_table_set_content(e, b, p["rows"]),
}


@app.route("/table", methods=["POST"])
@json_route
def table_op(data):
    doc_id = data.get("doc_id")
    entry = require_doc(doc_id)
    block_id = data.get("block_id")
    operation = data.get("operation")

    if operation not in TABLE_OPS:
        raise ValueError(f"unknown operation '{operation}' - "
                         f"expected one of {sorted(TABLE_OPS)}")
    require_blocks(entry, [block_id], "block_id")
    if operation == "merge":
        require_blocks(entry, [data.get("other_block_id")], "other_block_id")
        if data.get("other_block_id") == block_id:
            raise ValueError("cannot merge a table with itself")

    with mutation(entry):
        try:
            result = TABLE_OPS[operation](entry, block_id, data)
        except (KeyError, IndexError) as e:
            raise ValueError(f"missing or invalid parameter for '{operation}': {e}")

    payload = doc_state(doc_id)
    if isinstance(result, dict):
        payload["result"] = result
    return jsonify(payload)


@app.route("/edit_text", methods=["POST"])
@json_route
def edit_text(data):
    doc_id = data.get("doc_id")
    entry = require_doc(doc_id)
    block_id = data.get("block_id")
    text = data.get("text", "")

    require_blocks(entry, [block_id], "block_id")

    with mutation(entry):
        op_edit_text(entry, block_id, text)

    return jsonify(doc_state(doc_id))


@app.route("/undo", methods=["POST"])
@json_route
def undo(data):
    doc_id = data.get("doc_id")
    entry = require_doc(doc_id)
    if not restore_last_snapshot(entry):
        raise ValueError("nothing to undo")
    return jsonify(doc_state(doc_id))


@app.route("/export/<doc_id>", methods=["GET"])
def export(doc_id):
    if doc_id not in DOC_STORE:
        return jsonify({"error": "unknown doc_id"}), 400

    entry = DOC_STORE[doc_id]

    # Diagnostic: report the real merged-cell structure of each table in the
    # document being exported. If a table has vertically-merged cells, the same
    # underlying <w:tc> element is shared across rows - we count distinct cells
    # per column so you can confirm the EXPORTED file has true merges even
    # though the on-screen editor grid shows repeated values.
    try:
        for ti, tbl in enumerate(entry["document"].tables):
            n_rows, n_cols = len(tbl.rows), len(tbl.columns)
            merged_cols = []
            for c in range(n_cols):
                distinct = len({tbl.cell(r, c)._tc for r in range(n_rows)})
                if distinct < n_rows:
                    merged_cols.append(f"col{c}:{distinct}cells/{n_rows}rows")
            log.info("[EXPORT] doc %s table %s is %sx%s; vertically-merged columns: %s",
                     doc_id, ti, n_rows, n_cols, merged_cols or "none")
    except Exception as e:
        log.warning("[EXPORT] merge-inspection failed: %s", e)

    buf = io.BytesIO()
    entry["document"].save(buf)   # the SAME document object - never rebuilt
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=entry["filename"].rsplit(".", 1)[0] + "_edited.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    # use_reloader=False is deliberate: the auto-reloader restarts the whole
    # process the moment it sees ANY file change in this directory (including
    # non-.py files on some setups), which silently drops in-flight requests
    # with no response at all - this is what "upload just hangs forever"
    # usually turns out to be. threaded=True lets the dev server handle more
    # than one request at a time.
    app.run(debug=True, use_reloader=False, threaded=True, host="0.0.0.0", port=5000)
