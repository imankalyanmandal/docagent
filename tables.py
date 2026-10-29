"""
Merge-aware table editing.

Word tables are not grids of cells. A row is a list of <w:tc> cells; a cell
can span several grid columns (<w:gridSpan>, i.e. colspan) and can continue a
cell from the row above (<w:vMerge>, i.e. rowspan). Anything that edits a
table by rewriting a plain text matrix - "write row 2, column 3" - breaks as
soon as a table has merged cells: the row no longer has a cell per column, so
the edit lands in the wrong place, rows end up covering different numbers of
columns, and Word reports the file as damaged.

This module edits the real structure instead:

  layout(table)          -> per row, which cell covers which grid columns
  apply_layout(table, row_map, col_map)
                         -> the ONE structural primitive. Each new row / column
                            says which old row / column it comes from (or None
                            for a brand-new one). Kept rows and cells are moved,
                            not rebuilt, so their formatting and merges survive.
                            A new column inside a colspan widens the span; a new
                            row inside a rowspan extends it; deleting the first
                            row of a rowspan hands the merge to the row below.
  write_values(table, rows)
                         -> writes a value matrix onto the (possibly merged)
                            cells, touching only cells whose value changed
  infer_maps(old, new)   -> works out row/column maps from two value matrices,
                            for the table editor popup, which only sends the
                            final grid (including moved rows, so a reordered row
                            keeps its own formatting)
  regions / flatten / apply_regions
                         -> read, remove and re-create merges, for operations
                            that legitimately re-shape a table (transpose,
                            replacing one table with another)

Every operation leaves the table structurally valid: each row covers exactly
the tblGrid width, and every vMerge "continue" cell sits under a merged cell
with the same column span.
"""
from copy import deepcopy
from difflib import SequenceMatcher

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell

import fidelity

W = lambda tag: qn("w:" + tag)


# ---------------------------------------------------------------------------
# reading the structure
# ---------------------------------------------------------------------------

def _tcpr(tc, create=False):
    pr = tc.find(W("tcPr"))
    if pr is None and create:
        pr = OxmlElement("w:tcPr")
        tc.insert(0, pr)
    return pr


def span_of(tc):
    pr = _tcpr(tc)
    gs = pr.find(W("gridSpan")) if pr is not None else None
    try:
        return max(1, int(gs.get(W("val")))) if gs is not None else 1
    except (TypeError, ValueError):
        return 1


def vmerge_of(tc):
    """None, 'restart' or 'continue'."""
    pr = _tcpr(tc)
    vm = pr.find(W("vMerge")) if pr is not None else None
    if vm is None:
        return None
    return "restart" if vm.get(W("val")) == "restart" else "continue"


def set_span(tc, span):
    pr = _tcpr(tc, create=span > 1)
    if pr is None:
        return
    gs = pr.find(W("gridSpan"))
    if span <= 1:
        if gs is not None:
            pr.remove(gs)
        return
    if gs is None:
        gs = OxmlElement("w:gridSpan")
        # CT_TcPr order: cnfStyle, tcW, gridSpan, hMerge, vMerge, ...
        anchor = pr.find(W("tcW"))
        if anchor is None:
            anchor = pr.find(W("cnfStyle"))
        if anchor is not None:
            anchor.addnext(gs)
        else:
            pr.insert(0, gs)
    gs.set(W("val"), str(span))


def set_vmerge(tc, state):
    pr = _tcpr(tc, create=state is not None)
    if pr is None:
        return
    vm = pr.find(W("vMerge"))
    if state is None:
        if vm is not None:
            pr.remove(vm)
        return
    if vm is None:
        vm = OxmlElement("w:vMerge")
        anchor = None
        for tag in ("hMerge", "gridSpan", "tcW", "cnfStyle"):
            anchor = pr.find(W(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addnext(vm)
        else:
            pr.insert(0, vm)
    if state == "restart":
        vm.set(W("val"), "restart")
    elif W("val") in vm.attrib:
        del vm.attrib[W("val")]


def rows_of(tbl):
    return tbl.findall(W("tr"))


def row_layout(tr):
    """[(start_col, span, tc)] for the cells of one row."""
    out, col = [], _grid_before(tr)
    for tc in tr.findall(W("tc")):
        s = span_of(tc)
        out.append((col, s, tc))
        col += s
    return out


def _grid_before(tr):
    el = tr.find(f"{W('trPr')}/{W('gridBefore')}")
    try:
        return int(el.get(W("val"))) if el is not None else 0
    except (TypeError, ValueError):
        return 0


def grid_width(tbl):
    widths = [0]
    for tr in rows_of(tbl):
        lay = row_layout(tr)
        ga = tr.find(f"{W('trPr')}/{W('gridAfter')}")
        extra = int(ga.get(W("val"))) if ga is not None else 0
        widths.append((lay[-1][0] + lay[-1][1] if lay else _grid_before(tr)) + extra)
    return max(widths)


def owners(tbl):
    """
    {(row, col): tc} - the cell that DISPLAYS at each grid position. A vMerge
    continue cell displays its merge head (the restart cell above it).
    """
    own = {}
    for r, tr in enumerate(rows_of(tbl)):
        for start, span, tc in row_layout(tr):
            head = tc
            if vmerge_of(tc) == "continue" and (r - 1, start) in own:
                head = own[(r - 1, start)]
            for c in range(start, start + span):
                own[(r, c)] = head
    return own


def cell_text(tc, table=None):
    text = _Cell(tc, table).text
    lines = text.split("\n")
    while lines and not lines[-1].strip():
        lines.pop()
    while lines and not lines[0].strip():
        lines.pop(0)
    return "\n".join(lines)


def matrix(table):
    """Rectangular value matrix; a merged value repeats in every position."""
    tbl = table._tbl
    n_rows, width = len(rows_of(tbl)), grid_width(tbl)
    own = owners(tbl)
    cache = {}
    out = []
    for r in range(n_rows):
        row = []
        for c in range(width):
            tc = own.get((r, c))
            if tc is None:
                row.append("")
                continue
            if id(tc) not in cache:
                cache[id(tc)] = cell_text(tc, table)
            row.append(cache[id(tc)])
        out.append(row)
    return out


def regions(table):
    """{(top, left, bottom, right)} of every cell covering >1 grid position."""
    own = owners(table._tbl)
    box = {}
    for (r, c), tc in own.items():
        t, l, b, rr = box.get(id(tc), (r, c, r, c))
        box[id(tc)] = (min(t, r), min(l, c), max(b, r), max(rr, c))
    return {v for v in box.values() if v[0] != v[2] or v[1] != v[3]}


def has_vertical_merge(table, rows=None):
    for r, tr in enumerate(rows_of(table._tbl)):
        if rows is not None and r not in rows:
            continue
        if any(vmerge_of(tc) for _, _, tc in row_layout(tr)):
            return True
    return False


# ---------------------------------------------------------------------------
# normalisation
# ---------------------------------------------------------------------------

def _blank_like(tc):
    """A clone of tc with the same cell/paragraph/run formatting but no text
    and no merge flags."""
    new = deepcopy(tc)
    for child in list(new):
        if child.tag != W("tcPr"):
            new.remove(child)
    pr = new.find(W("tcPr"))
    if pr is not None:
        for tag in ("gridSpan", "vMerge", "hMerge"):
            for el in pr.findall(W(tag)):
                pr.remove(el)
    p = OxmlElement("w:p")
    first_p = tc.find(W("p"))
    if first_p is not None:
        ppr = first_p.find(W("pPr"))
        if ppr is not None:
            p.append(deepcopy(ppr))
        r = first_p.find(W("r"))
        rpr = r.find(W("rPr")) if r is not None else None
        if rpr is not None:
            stub = OxmlElement("w:r")
            stub.append(deepcopy(rpr))
            p.append(stub)
    new.append(p)
    return new


def _plain_tc():
    tc = OxmlElement("w:tc")
    tc.append(OxmlElement("w:tcPr"))
    tc.append(OxmlElement("w:p"))
    return tc


def normalize(table):
    """
    Make the table regular enough to edit:
      - rows that start late / end early (gridBefore / gridAfter) or are short
        get real empty cells, so every row covers the full width
      - tblGrid exists and has exactly one gridCol per grid column
      - every cell has at least one paragraph
    Returns the grid width.
    """
    tbl = table._tbl
    width = grid_width(tbl)
    for tr in rows_of(tbl):
        trpr = tr.find(W("trPr"))
        lay = row_layout(tr)
        template = lay[0][2] if lay else None
        before = _grid_before(tr)
        if trpr is not None:
            for tag in ("gridBefore", "gridAfter", "wBefore", "wAfter"):
                for el in trpr.findall(W(tag)):
                    trpr.remove(el)
        first_tc = tr.find(W("tc"))
        for _ in range(before):
            cell = _blank_like(template) if template is not None else _plain_tc()
            if first_tc is not None:
                first_tc.addprevious(cell)
            else:
                tr.append(cell)
        lay = row_layout(tr)
        covered = lay[-1][0] + lay[-1][1] if lay else 0
        for _ in range(width - covered):
            last = tr.findall(W("tc"))
            tr.append(_blank_like(last[-1]) if last else _plain_tc())
        for tc in tr.findall(W("tc")):
            if tc.find(W("p")) is None:
                tc.append(OxmlElement("w:p"))
    _ensure_grid(tbl, width)
    fix_vmerges(table)
    return width


def _ensure_grid(tbl, width):
    grid = tbl.find(W("tblGrid"))
    if grid is None:
        grid = OxmlElement("w:tblGrid")
        tblpr = tbl.find(W("tblPr"))
        if tblpr is not None:
            tblpr.addnext(grid)
        else:
            tbl.insert(0, grid)
    cols = grid.findall(W("gridCol"))
    if len(cols) == width:
        return
    template = cols[-1] if cols else None
    for gc in cols[width:]:
        grid.remove(gc)
    for _ in range(width - len(cols)):
        grid.append(deepcopy(template) if template is not None else OxmlElement("w:gridCol"))


def fix_vmerges(table):
    """
    Repair vertical merges after rows moved, appeared or disappeared:
      - a 'continue' with no matching merged cell above becomes a 'restart'
      - a 'restart' with no 'continue' below becomes a plain cell
    """
    trs = rows_of(table._tbl)
    lays = [row_layout(tr) for tr in trs]
    for r, lay in enumerate(lays):
        above = {(s, k): tc for s, k, tc in lays[r - 1]} if r else {}
        for s, k, tc in lay:
            if vmerge_of(tc) == "continue":
                up = above.get((s, k))
                if up is None or vmerge_of(up) is None:
                    set_vmerge(tc, "restart")
    for r, lay in enumerate(lays):
        below = {(s, k): tc for s, k, tc in lays[r + 1]} if r + 1 < len(lays) else {}
        for s, k, tc in lay:
            if vmerge_of(tc) is not None:
                down = below.get((s, k))
                if down is None or vmerge_of(down) != "continue":
                    if vmerge_of(tc) == "restart":
                        set_vmerge(tc, None)
    # a lone 'continue' (restart removed above it) can't happen: every
    # continue now has a merged cell above, which keeps its own flag


# ---------------------------------------------------------------------------
# the structural primitive
# ---------------------------------------------------------------------------

def apply_layout(table, row_map, col_map):
    """
    Rebuild the table so that new row i comes from old row row_map[i] (None =
    a new, empty row) and new column j from old column col_map[j] (None = new
    empty column). Old rows/columns not mentioned are deleted.
    """
    tbl = table._tbl
    width = normalize(table)
    old_trs = rows_of(tbl)
    # remember each 'continue' cell's head: if the head's row is deleted (or
    # rows move apart), the continue cell becomes the new head and must show
    # the merged content instead of going blank
    old_own = owners(tbl)
    heads = []
    for r, tr in enumerate(old_trs):
        for s, k, tc in row_layout(tr):
            if vmerge_of(tc) == "continue" and old_own.get((r, s)) is not tc:
                heads.append((tc, old_own[(r, s)]))
    grid = tbl.find(W("tblGrid"))
    old_gc = grid.findall(W("gridCol"))
    identity_cols = list(col_map) == list(range(width))

    # 1. columns, row by row (only when columns actually change)
    if not identity_cols:
        for tr in old_trs:
            _relayout_row(tr, col_map)
        for gc in old_gc:
            grid.remove(gc)
        for j, oc in enumerate(col_map):
            src = oc if oc is not None else _nearest_mapped(col_map, j)
            grid.append(deepcopy(old_gc[src]) if src is not None and src < len(old_gc)
                        else OxmlElement("w:gridCol"))
        _keep_total_width(old_gc, grid.findall(W("gridCol")))

    # 2. rows
    new_trs, used, inserted = [], set(), []
    for i, oi in enumerate(row_map):
        if oi is None:
            inserted.append(i)
            new_trs.append(None)
            continue
        tr = old_trs[oi]
        if oi in used:
            tr = deepcopy(tr)
        used.add(oi)
        new_trs.append(tr)
    for i in inserted:
        tmpl = _template_row(new_trs, row_map, old_trs, i)
        new_trs[i] = _new_row(tmpl, len(col_map), keep_header=(i == 0))

    for tr in old_trs:
        tbl.remove(tr)
    for tr in new_trs:
        tbl.append(tr)

    # 3. a row inserted INSIDE a rowspan extends it (what Word does)
    trs = rows_of(tbl)
    ins_set = set(inserted)
    for i in inserted:
        if i == 0:
            continue
        below = next((k for k in range(i + 1, len(trs)) if k not in ins_set), None)
        if below is None:
            continue
        up = {(s, k): tc for s, k, tc in row_layout(trs[i - 1])}
        for s, k, tc in row_layout(trs[below]):
            if vmerge_of(tc) != "continue" or k < 1:
                continue
            above = up.get((s, k))
            if above is None or vmerge_of(above) is None:
                continue
            _merge_into_continue(trs[i], s, k, tc)

    fix_vmerges(table)
    for tc, head in heads:
        if tc.getparent() is not None and vmerge_of(tc) != "continue" \
                and not cell_text(tc).strip():
            for child in list(tc):
                if child.tag != W("tcPr"):
                    tc.remove(child)
            for child in head:
                if child.tag != W("tcPr"):
                    tc.append(deepcopy(child))
    fix_widths(table)


def write_exclusive(table, cells):
    """
    Write {(row, col): value} into cells that belong ONLY to those positions'
    row (for a new row) or column (for a new column). A position that is part
    of a merged cell extending beyond it keeps the merged value - e.g. a new
    row inserted inside a rowspan shows the rowspan's value there.
    """
    own = owners(table._tbl)
    spread = {}
    for (r, c), tc in own.items():
        spread.setdefault(id(tc), (tc, set()))[1].add((r, c))
    done = set()
    for (r, c), value in cells.items():
        tc = own.get((r, c))
        if tc is None or id(tc) in done:
            continue
        positions = spread[id(tc)][1]
        if positions <= set(cells) and value:
            set_cell_text(tc, value, table)
            done.add(id(tc))


def _keep_total_width(old_cols, new_cols):
    """New columns copy a neighbour's width, which would push the table past
    the page margin; scale every column so the table keeps its old width
    (what Word does). A table that lost columns is left narrower."""
    try:
        old_total = sum(int(gc.get(W("w"))) for gc in old_cols)
        new_w = [int(gc.get(W("w"))) for gc in new_cols]
    except (TypeError, ValueError):
        return
    new_total = sum(new_w)
    if not old_total or new_total <= old_total:
        return
    scaled = [max(1, w * old_total // new_total) for w in new_w]
    scaled[-1] += old_total - sum(scaled)          # absorb rounding
    for gc, w in zip(new_cols, scaled):
        gc.set(W("w"), str(w))


def _nearest_mapped(col_map, j):
    for d in range(1, len(col_map)):
        for k in (j - d, j + d):
            if 0 <= k < len(col_map) and col_map[k] is not None:
                return col_map[k]
    return None


def _relayout_row(tr, col_map):
    lay = row_layout(tr)
    mapped = {c for c in col_map if c is not None}

    def follows(last_oc, oc):
        # oc continues the span from last_oc if every old column between
        # them was deleted (so the span simply lost those columns)
        return last_oc is not None and last_oc < oc and \
            all(x not in mapped for x in range(last_oc + 1, oc))

    cover = {}
    for s, k, tc in lay:
        for c in range(s, s + k):
            cover[c] = (tc, s, k)
    groups = []                 # [tc, span, last_old_col]
    used = set()
    for j, oc in enumerate(col_map):
        if oc is None:
            # inside a span? (neighbours on both sides map to the same cell,
            # contiguously) -> widen that span instead of adding a cell
            left = groups[-1] if groups else None
            nxt = next((col_map[x] for x in range(j + 1, len(col_map)) if col_map[x] is not None), None)
            if left is not None and nxt is not None and nxt in cover and \
                    cover[nxt][0] is left[0] and follows(left[2], nxt):
                left[1] += 1
                continue
            ref = _nearest_mapped(col_map, j)
            tmpl = cover[ref][0] if ref is not None and ref in cover else (lay[0][2] if lay else None)
            groups.append([_blank_like(tmpl) if tmpl is not None else _plain_tc(), 1, None])
            continue
        tc = cover[oc][0]
        last = groups[-1] if groups else None
        if last is not None and last[0] is tc and follows(last[2], oc):
            last[1] += 1
            last[2] = oc
            continue
        if id(tc) in used:
            tc = deepcopy(tc)
        used.add(id(tc))
        groups.append([tc, 1, oc])
    for _, _, tc in lay:
        tr.remove(tc)
    for tc, span, _ in groups:
        set_span(tc, span)
        tr.append(tc)


def _template_row(new_trs, row_map, old_trs, i):
    for k in range(i - 1, -1, -1):
        if new_trs[k] is not None:
            return new_trs[k]
    for k in range(i + 1, len(new_trs)):
        if new_trs[k] is not None:
            return new_trs[k]
    return old_trs[0] if old_trs else None


def _new_row(template, width, keep_header):
    """An empty row: one plain cell per grid column, formatted like the
    template row's cells."""
    tr = OxmlElement("w:tr")
    if template is None:
        for _ in range(width):
            tr.append(_plain_tc())
        return tr
    for tag in ("tblPrEx", "trPr"):
        el = template.find(W(tag))
        if el is not None:
            clone = deepcopy(el)
            if tag == "trPr":
                for t in ("ins", "del", "trPrChange", "gridBefore", "gridAfter",
                          "wBefore", "wAfter") + (() if keep_header else ("tblHeader",)):
                    for x in clone.findall(W(t)):
                        clone.remove(x)
            tr.append(clone)
    cover = {}
    for s, k, tc in row_layout(template):
        for c in range(s, s + k):
            cover[c] = tc
    for c in range(width):
        src = cover.get(c)
        tr.append(_blank_like(src) if src is not None else _plain_tc())
    return tr


def _merge_into_continue(tr, start, span, model_tc):
    lay = row_layout(tr)
    inside = [tc for s, k, tc in lay if s >= start and s + k <= start + span]
    if sum(span_of(tc) for tc in inside) != span or not inside:
        return
    cont = _blank_like(model_tc)
    set_span(cont, span)
    set_vmerge(cont, "continue")
    inside[0].addprevious(cont)
    for tc in inside:
        tr.remove(tc)


def fix_widths(table):
    """Set each cell's tcW to the sum of the grid columns it spans (dxa)."""
    tbl = table._tbl
    cols = tbl.findall(f"{W('tblGrid')}/{W('gridCol')}")
    try:
        widths = [int(gc.get(W("w"))) for gc in cols]
    except (TypeError, ValueError):
        return
    for tr in rows_of(tbl):
        for s, k, tc in row_layout(tr):
            if s + k > len(widths):
                continue
            pr = _tcpr(tc, create=True)
            tcw = pr.find(W("tcW"))
            if tcw is None:
                tcw = OxmlElement("w:tcW")
                cnf = pr.find(W("cnfStyle"))
                if cnf is not None:
                    cnf.addnext(tcw)
                else:
                    pr.insert(0, tcw)
            tcw.set(W("w"), str(sum(widths[s:s + k])))
            tcw.set(W("type"), "dxa")


# ---------------------------------------------------------------------------
# values
# ---------------------------------------------------------------------------

def write_values(table, rows):
    """
    Write a value matrix onto the table. For a merged cell, the value at ANY
    of its positions that differs from the cell's current text wins (so an
    edit made in a covered position of the editor grid isn't lost). Cells
    whose value is unchanged are not touched at all - their mixed formatting
    survives.
    """
    tbl = table._tbl
    own = owners(tbl)
    order, positions = [], {}
    for (r, c), tc in sorted(own.items()):
        if id(tc) not in positions:
            positions[id(tc)] = (tc, [])
            order.append(id(tc))
        positions[id(tc)][1].append((r, c))
    for key in order:
        tc, pos = positions[key]
        current = cell_text(tc, table)
        values = [rows[r][c] for r, c in pos if r < len(rows) and c < len(rows[r])]
        chosen = next((v for v in values if v != current), current)
        if chosen != current:
            set_cell_text(tc, chosen, table)


def set_cell_text(tc, text, table=None):
    paras = tc.findall(W("p"))
    if not paras:
        tc.append(OxmlElement("w:p"))
        paras = tc.findall(W("p"))
    for p in paras[1:]:
        tc.remove(p)              # nested tables (w:tbl) are kept as content
    from docx.text.paragraph import Paragraph
    fidelity.set_plain_text(Paragraph(paras[0], _Cell(tc, table)), text)


# ---------------------------------------------------------------------------
# inferring maps (table editor popup sends only the final grid)
# ---------------------------------------------------------------------------

def infer_maps(old, new):
    """(row_map, col_map) explaining how matrix `old` became `new`."""
    old_w = len(old[0]) if old else 0
    new_w = len(new[0]) if new else 0
    if old_w == new_w:
        col_map = list(range(new_w))
    else:
        if len(old) == len(new):
            key_old = [tuple(r[c] for r in old) for c in range(old_w)]
            key_new = [tuple(r[c] for r in new) for c in range(new_w)]
        else:                               # rows changed too - match on the header row
            key_old = [old[0][c] for c in range(old_w)] if old else []
            key_new = [new[0][c] for c in range(new_w)] if new else []
        col_map = _sequence_map(key_old, key_new)
    project = lambda row, cmap: tuple(row[c] if c is not None and c < len(row) else None for c in cmap)
    old_keys = [project(r, col_map) for r in old]
    new_keys = [tuple(v if col_map[j] is not None else None for j, v in enumerate(r)) for r in new]
    row_map = _sequence_map(old_keys, new_keys)
    return row_map, col_map


def _sequence_map(old, new):
    """
    For each item of `new`, the index of the `old` item it came from, or
    None if it's new. Unchanged runs map directly; an item moved elsewhere
    (identical content) maps to its old self; remaining changed items pair up
    positionally (an edit in place).
    """
    sm = SequenceMatcher(None, old, new, autojunk=False)
    result = [None] * len(new)
    used = set()
    blocks = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                result[j1 + k] = i1 + k
                used.add(i1 + k)
        elif tag == "replace":
            blocks.append((i1, i2, j1, j2))
        elif tag == "insert":
            blocks.append((i1, i1, j1, j2))
        else:
            blocks.append((i1, i2, j1, j1))
    # moves: identical content somewhere else
    for j in range(len(new)):
        if result[j] is None:
            for i in sorted(range(len(old)), key=lambda i: abs(i - j)):
                if i not in used and old[i] == new[j]:
                    result[j] = i
                    used.add(i)
                    break
    # edits in place
    for i1, i2, j1, j2 in blocks:
        free_old = [i for i in range(i1, i2) if i not in used]
        free_new = [j for j in range(j1, j2) if result[j] is None]
        for i, j in zip(free_old, free_new):
            result[j] = i
            used.add(i)
    return result


# ---------------------------------------------------------------------------
# re-shaping: flatten, re-merge, transpose
# ---------------------------------------------------------------------------

def flatten(table):
    """Remove every merge: one cell per grid position. The merged value stays
    in the top-left cell; the uncovered positions are empty cells."""
    normalize(table)
    for tr in rows_of(table._tbl):
        for s, k, tc in row_layout(tr):
            set_vmerge(tc, None)
            if k > 1:
                set_span(tc, 1)
                cursor = tc
                for _ in range(k - 1):
                    blank = _blank_like(tc)
                    cursor.addnext(blank)
                    cursor = blank
    for tr in rows_of(table._tbl):
        for s, k, tc in row_layout(tr):
            if vmerge_of(tc) == "continue":
                set_vmerge(tc, None)
    fix_widths(table)


def apply_regions(table, regs):
    """Merge the given (top, left, bottom, right) regions of a table whose
    affected positions are currently plain single cells."""
    for top, left, bottom, right in sorted(regs):
        trs = rows_of(table._tbl)
        if bottom >= len(trs):
            continue
        span = right - left + 1
        for r in range(top, bottom + 1):
            lay = row_layout(trs[r])
            inside = [tc for s, k, tc in lay if s >= left and s + k <= right + 1]
            if not inside or sum(span_of(tc) for tc in inside) != span:
                break
            keep = inside[0]
            for extra in inside[1:]:
                trs[r].remove(extra)
            set_span(keep, span)
            if bottom > top:
                set_vmerge(keep, "restart" if r == top else "continue")
                if r > top:
                    for p in keep.findall(W("p"))[1:]:
                        keep.remove(p)
                    for t in keep.iter(W("t")):
                        t.text = ""
    fix_vmerges(table)
    fix_widths(table)


def transpose(table):
    """Transpose cells AND merges; each cell keeps its own formatting."""
    regs = regions(table)
    flatten(table)
    tbl = table._tbl
    trs = rows_of(tbl)
    grid = [[tc for _, _, tc in row_layout(tr)] for tr in trs]
    n_rows, n_cols = len(grid), len(grid[0]) if grid else 0
    template_trpr = trs[0].find(W("trPr")) if trs else None
    total = None
    gcs = tbl.findall(f"{W('tblGrid')}/{W('gridCol')}")
    try:
        total = sum(int(gc.get(W("w"))) for gc in gcs)
    except (TypeError, ValueError):
        pass
    for tr in trs:
        tbl.remove(tr)
    for c in range(n_cols):
        tr = OxmlElement("w:tr")
        if template_trpr is not None:
            pr = deepcopy(template_trpr)
            for x in pr.findall(W("tblHeader")):
                pr.remove(x)
            tr.append(pr)
        for r in range(n_rows):
            tr.append(grid[r][c])
        tbl.append(tr)
    grid_el = tbl.find(W("tblGrid"))
    for gc in grid_el.findall(W("gridCol")):
        grid_el.remove(gc)
    for _ in range(n_rows):
        gc = OxmlElement("w:gridCol")
        if total:
            gc.set(W("w"), str(total // max(1, n_rows)))
        grid_el.append(gc)
    fix_widths(table)
    apply_regions(table, {(l, t, r, b) for t, l, b, r in regs})
