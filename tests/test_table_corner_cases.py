"""
Corner cases for tables with merged cells (colspan / rowspan) and every
operation that edits a table.

Run from the project root:  python tests/test_table_corner_cases.py

Every scenario goes through the real HTTP routes (so undo snapshots, locking
and resync are exercised too) and, after each edit, checks the table is still
STRUCTURALLY VALID Word XML:

  - every row covers exactly the tblGrid width (sum of gridSpans)
  - every vertically-merged "continue" cell sits under a cell with the same
    start column and span that is itself part of a vertical merge
  - the document saves and re-opens

At the end every resulting document is opened in LibreOffice (when it's
installed) to catch anything python-docx tolerates but an office suite
rejects.
"""
import io
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("DOCAGENT_DEBUG", "0")

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import RGBColor

import app as appmod

passed = failed = 0
SAVED = []            # (name, bytes) of documents to open in LibreOffice at the end


def check(name, cond, detail=""):
    global passed, failed
    if cond:
        passed += 1
        print(f"  [PASS] {name}")
    else:
        failed += 1
        print(f"  [FAIL] {name}" + (f"  -> {detail}" if detail else ""))


def section(title):
    print(f"\n--- {title}")


W = lambda t: qn("w:" + t)
client = appmod.app.test_client()


# ---------------------------------------------------------------------------
# builders
# ---------------------------------------------------------------------------

def make_table(doc, rows, merges=()):
    """rows: list of lists of text. merges: (top, left, bottom, right) inclusive.
    The merged cell gets the top-left text; covered cells are blanked first so
    python-docx doesn't concatenate their text."""
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            t.cell(r, c).text = val
    for top, left, bottom, right in merges:
        text = rows[top][left]
        for r in range(top, bottom + 1):
            for c in range(left, right + 1):
                t.cell(r, c).text = ""
        m = t.cell(top, left).merge(t.cell(bottom, right))
        for p in m.paragraphs[1:]:
            p._p.getparent().remove(p._p)
        m.paragraphs[0].text = text
    return t


def upload(doc, name="t.docx"):
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    r = client.post("/upload", data={"file": (buf, name)}, content_type="multipart/form-data")
    assert r.status_code == 200, r.get_json()
    return r.get_json()["doc_id"]


def call(route, **payload):
    r = client.post(route, json=payload)
    return r.status_code, r.get_json()


def table_op(doc_id, block_id, operation, **params):
    return call("/table", doc_id=doc_id, block_id=block_id, operation=operation, **params)


def entry(doc_id):
    return appmod.DOC_STORE[doc_id]


def tables(doc_id):
    """(block_id, block) for every table, in order."""
    e = entry(doc_id)
    return [(k, b) for k, b in sorted(e["blocks"].items(), key=lambda kv: kv[1]["order"])
            if b["type"] == "table"]


def table_id(doc_id, n=0):
    return tables(doc_id)[n][0]


def rows_of(doc_id, n=0):
    return tables(doc_id)[n][1]["rows"]


def live_table(doc_id, n=0):
    return entry(doc_id)["refs"][table_id(doc_id, n)]


# ---------------------------------------------------------------------------
# structural inspection
# ---------------------------------------------------------------------------

def _span(tc):
    gs = tc.find(f"{W('tcPr')}/{W('gridSpan')}")
    return int(gs.get(W("val"))) if gs is not None else 1


def _vmerge(tc):
    vm = tc.find(f"{W('tcPr')}/{W('vMerge')}")
    if vm is None:
        return None
    return vm.get(W("val")) or "continue"


def row_layout(tr):
    """[(start_col, span, vmerge_state, tc)] for a row."""
    out, col = [], 0
    gb = tr.find(f"{W('trPr')}/{W('gridBefore')}")
    if gb is not None:
        col += int(gb.get(W("val")))
    for tc in tr.findall(W("tc")):
        out.append((col, _span(tc), _vmerge(tc), tc))
        col += _span(tc)
    return out


def validity_problems(table):
    """List of human-readable structural problems (empty = valid)."""
    tbl = table._tbl
    probs = []
    grid_cols = tbl.findall(f"{W('tblGrid')}/{W('gridCol')}")
    width = len(grid_cols)
    prev = None
    for r, tr in enumerate(tbl.findall(W("tr"))):
        lay = row_layout(tr)
        total = (lay[-1][0] + lay[-1][1]) if lay else 0
        ga = tr.find(f"{W('trPr')}/{W('gridAfter')}")
        if ga is not None:
            total += int(ga.get(W("val")))
        if total != width:
            probs.append(f"row {r} covers {total} grid cols, tblGrid has {width}")
        for start, span, vm, tc in lay:
            if vm == "continue":
                above = None if prev is None else \
                    next((x for x in prev if x[0] == start and x[1] == span), None)
                if above is None or above[2] is None:
                    probs.append(f"row {r} col {start}: vMerge continue with no merged cell above")
            if not tc.findall(W("p")):
                probs.append(f"row {r} col {start}: cell has no paragraph")
        prev = lay
    return probs


def merge_regions(table):
    """Set of (top, left, bottom, right) for every merged (multi-position) cell."""
    tbl = table._tbl
    trs = tbl.findall(W("tr"))
    owner = {}
    for r, tr in enumerate(trs):
        for start, span, vm, tc in row_layout(tr):
            key = None
            if vm == "continue" and r > 0:
                key = owner.get((r - 1, start))
            if key is None:
                key = (r, start)
            for c in range(start, start + span):
                owner[(r, c)] = key
    regions = {}
    for (r, c), key in owner.items():
        t, l, b, rr = regions.get(key, (r, c, r, c))
        regions[key] = (min(t, r), min(l, c), max(b, r), max(rr, c))
    return {v for v in regions.values() if v[0] != v[2] or v[1] != v[3]}


def assert_valid(name, doc_id, n=0):
    t = live_table(doc_id, n)
    probs = validity_problems(t)
    check(f"{name}: table structurally valid", not probs, "; ".join(probs[:3]))
    # save -> reopen round trip
    buf = io.BytesIO()
    entry(doc_id)["document"].save(buf)
    try:
        Document(io.BytesIO(buf.getvalue()))
        ok = True
    except Exception as e:  # pragma: no cover - reported as a failure
        ok, name = False, f"{name} ({e})"
    check(f"{name}: saves and re-opens", ok)
    SAVED.append((name, buf.getvalue()))


def rect(rows):
    return len({len(r) for r in rows}) == 1


# ---------------------------------------------------------------------------
# fixtures
# ---------------------------------------------------------------------------

def merged_doc():
    """
    5x4 table:
        H1   | H2 (colspan 2)  | H4
        A    | b1    | c1      | d1
        V    | b2    | c2      | d2      V = rowspan 3 (rows 1-3, col 0)
        V    | B (2x2 block: rows 3-4, cols 1-2)   | d3
        e    | B               | d4
    """
    d = Document()
    d.add_paragraph("before")
    rows = [["H1", "H2", "", "H4"],
            ["V", "b1", "c1", "d1"],
            ["", "b2", "c2", "d2"],
            ["", "B", "", "d3"],
            ["e", "", "", "d4"]]
    make_table(d, rows, merges=[(0, 1, 0, 2), (1, 0, 3, 0), (3, 1, 4, 2)])
    d.add_paragraph("after")
    return d


EXPECTED_REGIONS = {(0, 1, 0, 2), (1, 0, 3, 0), (3, 1, 4, 2)}

# ===========================================================================
section("parsing merged tables")

doc_id = upload(merged_doc())
rows = rows_of(doc_id)
check("block rows are rectangular", rect(rows), rows)
check("grid is 5 x 4", len(rows) == 5 and len(rows[0]) == 4)
check("colspan value repeated across its columns", rows[0][1] == rows[0][2] == "H2")
check("rowspan value repeated down its rows", rows[1][0] == rows[2][0] == rows[3][0] == "V")
check("2x2 block value in all four positions",
      rows[3][1] == rows[3][2] == rows[4][1] == rows[4][2] == "B")
check("merge regions detected", merge_regions(live_table(doc_id)) == EXPECTED_REGIONS,
      merge_regions(live_table(doc_id)))
assert_valid("fresh merged table", doc_id)

# tblGrid missing / wrong while the FIRST row has a colspan
d = Document()
t = make_table(d, [["wide", "", "x"], ["a", "b", "c"]], merges=[(0, 0, 0, 1)])
grid = t._tbl.find(W("tblGrid"))
t._tbl.remove(grid)
gid = upload(d)
status, body = table_op(gid, table_id(gid), "add_row", values=["1", "2", "3"])
check("missing tblGrid + colspan first row: op succeeds", status == 200, body)
if status == 200:
    check("missing tblGrid + colspan first row: grid rebuilt to 3 cols",
          len(live_table(gid)._tbl.findall(f"{W('tblGrid')}/{W('gridCol')}")) == 3)
    check("missing tblGrid + colspan first row: new row values intact",
          rows_of(gid)[-1] == ["1", "2", "3"], rows_of(gid))
    assert_valid("rebuilt tblGrid", gid)

# ===========================================================================
section("editing cell values in merged regions (set_content, same shape)")

doc_id = upload(merged_doc())
rows = [r[:] for r in rows_of(doc_id)]
rows[0][1] = "Header Two"                      # edit the anchor of a colspan
rows[2][0] = "Vertical"                        # edit a NON-anchor position of a rowspan
rows[1][3] = "d1 edited"                       # plain cell
status, body = table_op(doc_id, table_id(doc_id), "set_content", rows=rows)
check("set_content same shape: 200", status == 200, body)
new = rows_of(doc_id)
check("anchor edit lands in the colspan cell", new[0][1] == new[0][2] == "Header Two", new[0])
check("edit at a covered position updates the whole rowspan cell",
      new[1][0] == new[2][0] == new[3][0] == "Vertical", [r[0] for r in new])
check("plain cell edit", new[1][3] == "d1 edited")
check("merges untouched by value edits", merge_regions(live_table(doc_id)) == EXPECTED_REGIONS,
      merge_regions(live_table(doc_id)))
assert_valid("value edits", doc_id)

# formatting of UNCHANGED cells survives a table edit
d = Document()
t = make_table(d, [["Name", "Qty"], ["apple", "3"], ["pear", "1"]])
cell = t.cell(1, 0)
cell.paragraphs[0].text = ""
cell.paragraphs[0].add_run("ap").bold = True
cell.paragraphs[0].add_run("ple")
fid = upload(d)
rows = [r[:] for r in rows_of(fid)]
rows[2][1] = "9"
table_op(fid, table_id(fid), "set_content", rows=rows)
runs = live_table(fid).cell(1, 0).paragraphs[0].runs
check("untouched cell keeps mixed bold/plain runs",
      [(r.text, bool(r.bold)) for r in runs if r.text] == [("ap", True), ("ple", False)],
      [(r.text, r.bold) for r in runs])

# multi-line value round trip
rows = [r[:] for r in rows_of(fid)]
rows[2][0] = "line one\nline two"
table_op(fid, table_id(fid), "set_content", rows=rows)
check("multi-line cell value round-trips", rows_of(fid)[2][0] == "line one\nline two",
      repr(rows_of(fid)[2][0]))
rows = [r[:] for r in rows_of(fid)]
rows[1][1] = "ünïcødé ✓ 数字 🙂"
table_op(fid, table_id(fid), "set_content", rows=rows)
check("unicode / emoji cell value round-trips", rows_of(fid)[1][1] == "ünïcødé ✓ 数字 🙂")
rows = [r[:] for r in rows_of(fid)]
rows[1][1] = "x" * 5000
table_op(fid, table_id(fid), "set_content", rows=rows)
check("very long cell value round-trips", rows_of(fid)[1][1] == "x" * 5000)

# ragged rows from the UI are padded
status, _ = table_op(fid, table_id(fid), "set_content", rows=[["a", "b", "c"], ["d"]])
check("ragged set_content padded to rectangle",
      status == 200 and rows_of(fid) == [["a", "b", "c"], ["d", "", ""]], rows_of(fid))
status, _ = table_op(fid, table_id(fid), "set_content", rows=[])
check("empty set_content refused", status == 400)

# ===========================================================================
section("add_row")

doc_id = upload(merged_doc())
status, body = table_op(doc_id, table_id(doc_id), "add_row", values=["n0", "n1", "n2", "n3"])
check("append row: 200", status == 200, body)
check("append row: values", rows_of(doc_id)[-1] == ["n0", "n1", "n2", "n3"], rows_of(doc_id)[-1])
check("append row: existing merges kept",
      merge_regions(live_table(doc_id)) == EXPECTED_REGIONS, merge_regions(live_table(doc_id)))
assert_valid("append row", doc_id)

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "add_row", index=0, values=["top"])
check("insert row at top: values padded", rows_of(doc_id)[0] == ["top", "", "", ""], rows_of(doc_id)[0])
check("insert row at top: merges shift down by one",
      merge_regions(live_table(doc_id)) == {(1, 1, 1, 2), (2, 0, 4, 0), (4, 1, 5, 2)},
      merge_regions(live_table(doc_id)))
assert_valid("insert row at top", doc_id)

doc_id = upload(merged_doc())
# insert INSIDE the rowspan (between rows 1 and 2): the vertical merge extends
table_op(doc_id, table_id(doc_id), "add_row", index=2, values=["x", "y", "z", "w"])
regs = merge_regions(live_table(doc_id))
check("insert row inside a rowspan: the rowspan grows to 4 rows", (1, 0, 4, 0) in regs, regs)
check("insert row inside a rowspan: other values of the new row kept",
      rows_of(doc_id)[2][1:] == ["y", "z", "w"], rows_of(doc_id)[2])
check("insert row inside a rowspan: rowspan value shown in new row",
      rows_of(doc_id)[2][0] == "V", rows_of(doc_id)[2])
assert_valid("insert row inside rowspan", doc_id)

status, _ = table_op(doc_id, table_id(doc_id), "add_row", index=999)
check("add_row past the end clamps to append", status == 200 and rows_of(doc_id)[-1] == [""] * 4)
status, _ = table_op(doc_id, table_id(doc_id), "add_row", index=-5)
check("add_row negative index clamps to top", status == 200 and rows_of(doc_id)[0] == [""] * 4)
status, _ = table_op(doc_id, table_id(doc_id), "add_row", values=["1", "2", "3", "4", "5", "6"])
check("add_row with too many values truncates", status == 200 and rows_of(doc_id)[-1] == ["1", "2", "3", "4"])
assert_valid("add_row clamping", doc_id)

# ===========================================================================
section("delete_row")

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "delete_row", index=1)        # the rowspan's FIRST row
rows = rows_of(doc_id)
check("delete first row of a rowspan: rowspan value survives below",
      rows[1][0] == rows[2][0] == "V", [r[0] for r in rows])
check("delete first row of a rowspan: rowspan now 2 rows",
      (1, 0, 2, 0) in merge_regions(live_table(doc_id)), merge_regions(live_table(doc_id)))
check("delete first row of a rowspan: the right row went",
      [r[1] for r in rows] == ["H2", "b2", "B", "B"] and "b1" not in [r[1] for r in rows],
      [r[1] for r in rows])
assert_valid("delete rowspan head", doc_id)

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "delete_row", index=2)        # middle of the rowspan
check("delete middle row of a rowspan: rowspan shrinks to 2",
      (1, 0, 2, 0) in merge_regions(live_table(doc_id)), merge_regions(live_table(doc_id)))
check("delete middle row: row 'b2' gone", "b2" not in [r[1] for r in rows_of(doc_id)])
assert_valid("delete rowspan middle", doc_id)

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "delete_row", index=4)        # bottom of the 2x2 block
regs = merge_regions(live_table(doc_id))
check("delete bottom row of a 2x2 block: block becomes 1x2 colspan", (3, 1, 3, 2) in regs, regs)
check("delete bottom row of 2x2 block: value kept", rows_of(doc_id)[3][1:3] == ["B", "B"])
assert_valid("delete 2x2 bottom", doc_id)

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "delete_row", index=0)        # header with colspan
check("delete header row with a colspan", rows_of(doc_id)[0][0] == "V" and len(rows_of(doc_id)) == 4)
assert_valid("delete colspan header", doc_id)

status, body = table_op(doc_id, table_id(doc_id), "delete_row", index=99)
check("delete_row out of range refused", status == 400, body)
d = Document()
make_table(d, [["only", "row"]])
one = upload(d)
status, body = table_op(one, table_id(one), "delete_row", index=0)
check("deleting the only row refused", status == 400, body)

# ===========================================================================
section("add_column")

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "add_column", index=4, header="H5", values=["1", "2", "3", "4"])
rows = rows_of(doc_id)
check("append column: values", [r[4] for r in rows] == ["H5", "1", "2", "3", "4"], [r[4] for r in rows])
check("append column: merges kept", merge_regions(live_table(doc_id)) == EXPECTED_REGIONS,
      merge_regions(live_table(doc_id)))
assert_valid("append column", doc_id)

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "add_column", index=2, header="mid")   # inside the H2 colspan
regs = merge_regions(live_table(doc_id))
check("insert column inside a colspan: colspan widens to 3", (0, 1, 0, 3) in regs, regs)
check("insert column inside the 2x2 block: block widens", (3, 1, 4, 3) in regs, regs)
check("insert column: rowspan untouched", (1, 0, 3, 0) in regs, regs)
check("insert column: plain rows get the new empty cell at index 2",
      [r[2] for r in rows_of(doc_id)[1:3]] == ["", ""], rows_of(doc_id))
assert_valid("insert column inside colspan", doc_id)

doc_id = upload(merged_doc())
grid_w = lambda: sum(int(gc.get(W("w"))) for gc in
                     live_table(doc_id)._tbl.findall(f"{W('tblGrid')}/{W('gridCol')}"))
width_before = grid_w()
table_op(doc_id, table_id(doc_id), "add_column", header="x")
table_op(doc_id, table_id(doc_id), "add_column", header="y")
check("adding columns keeps the table's total width (stays inside the page)",
      abs(grid_w() - width_before) <= 1, (width_before, grid_w()))
tcw_ok = all(
    int(tc.find(f"{W('tcPr')}/{W('tcW')}").get(W("w"))) ==
    sum(int(gc.get(W("w"))) for gc in live_table(doc_id)._tbl.findall(f"{W('tblGrid')}/{W('gridCol')}")[s:s + k])
    for tr in live_table(doc_id)._tbl.findall(W("tr")) for s, k, _, tc in row_layout(tr))
check("every cell's width equals the grid columns it spans", tcw_ok)

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "add_column", index=0, header="first")
regs = merge_regions(live_table(doc_id))
check("insert column at 0: everything shifts right",
      regs == {(0, 2, 0, 3), (1, 1, 3, 1), (3, 2, 4, 3)}, regs)
check("insert column at 0: new column is NOT vertically merged",
      [r[0] for r in rows_of(doc_id)] == ["first", "", "", "", ""], [r[0] for r in rows_of(doc_id)])
assert_valid("insert column at 0", doc_id)

# ===========================================================================
section("delete_column")

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "delete_column", index=2)   # covered by H2 colspan + 2x2 block
regs = merge_regions(live_table(doc_id))
rows = rows_of(doc_id)
check("delete column under a colspan: colspan shrinks to a single cell", (0, 1, 0, 1) not in regs
      and rows[0] == ["H1", "H2", "H4"], rows[0])
check("delete column under a 2x2 block: block becomes a rowspan", (3, 1, 4, 1) in regs, regs)
check("delete column: plain column gone", [r[2] for r in rows[1:3]] == ["d1", "d2"], rows)
assert_valid("delete column under spans", doc_id)

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "delete_column", index=0)   # the rowspan column
check("delete the rowspan's column", rows_of(doc_id)[1] == ["b1", "c1", "d1"], rows_of(doc_id)[1])
check("delete the rowspan's column: other merges shift left",
      merge_regions(live_table(doc_id)) == {(0, 0, 0, 1), (3, 0, 4, 1)}, merge_regions(live_table(doc_id)))
assert_valid("delete rowspan column", doc_id)

doc_id = upload(merged_doc())
table_op(doc_id, table_id(doc_id), "delete_column", index=1)   # START column of the colspan
check("delete the first column of a colspan keeps the header text",
      rows_of(doc_id)[0] == ["H1", "H2", "H4"], rows_of(doc_id)[0])
assert_valid("delete colspan start column", doc_id)

status, body = table_op(doc_id, table_id(doc_id), "delete_column", index=-1)
check("delete_column negative index refused", status == 400, body)
d = Document()
make_table(d, [["a"], ["b"]])
onec = upload(d)
status, body = table_op(onec, table_id(onec), "delete_column", index=0)
check("deleting the only column refused", status == 400, body)

# ===========================================================================
section("sort")

d = Document()
t = make_table(d, [["Item", "Qty", "Note"], ["banana", "10", "x"], ["apple", "2", "y"],
                   ["cherry", "1,000", "wide"]], merges=[(3, 1, 3, 2)])
# shade the banana row so we can see rows move as units
for c in t.rows[1].cells:
    c._tc.get_or_add_tcPr().append(parse_xml(
        f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="FFFF00"/>'))
sid = upload(d)
status, body = table_op(sid, table_id(sid), "sort", column=0)
check("sort with a horizontal merge: 200", status == 200, body)
check("sort: alphabetical order", [r[0] for r in rows_of(sid)] == ["Item", "apple", "banana", "cherry"])
check("sort: the colspan travelled with its row", (3, 1, 3, 2) in merge_regions(live_table(sid)),
      merge_regions(live_table(sid)))
banana_tr = live_table(sid)._tbl.findall(W("tr"))[2]
check("sort: row shading travelled with its row",
      banana_tr.find(f".//{W('shd')}") is not None)
status, body = table_op(sid, table_id(sid), "sort", column=1)
check("sort numeric with thousands separator", [r[1] for r in rows_of(sid)[1:]] == ["2", "10", "1,000"],
      rows_of(sid))
table_op(sid, table_id(sid), "sort", column=1, descending=True)
check("sort descending", [r[1] for r in rows_of(sid)[1:]] == ["1,000", "10", "2"])
assert_valid("sort", sid)

doc_id = upload(merged_doc())
undo_before = len(entry(doc_id)["undo"])
status, body = table_op(doc_id, table_id(doc_id), "sort", column=3)
check("sort across a rowspan refused (would tear the merge)", status == 400, body)
check("refused sort leaves table unchanged", merge_regions(live_table(doc_id)) == EXPECTED_REGIONS)
check("refused sort consumes no undo slot", len(entry(doc_id)["undo"]) == undo_before)
status, body = table_op(doc_id, table_id(doc_id), "sort", column=9)
check("sort column out of range refused", status == 400)

# ===========================================================================
section("transpose")

d = Document()
make_table(d, [["a", "b", "c"], ["1", "2", "3"]])
tid = upload(d)
table_op(tid, table_id(tid), "transpose")
check("transpose plain 2x3 -> 3x2", rows_of(tid) == [["a", "1"], ["b", "2"], ["c", "3"]], rows_of(tid))
table_op(tid, table_id(tid), "transpose")
check("transpose twice is identity", rows_of(tid) == [["a", "b", "c"], ["1", "2", "3"]])
assert_valid("transpose plain", tid)

doc_id = upload(merged_doc())
status, body = table_op(doc_id, table_id(doc_id), "transpose")
check("transpose merged table: 200", status == 200, body)
rows = rows_of(doc_id)
check("transpose merged table: 4 x 5 rectangle", rect(rows) and len(rows) == 4 and len(rows[0]) == 5, rows)
check("transpose merged table: merges move with the transpose",
      merge_regions(live_table(doc_id)) == {(1, 0, 2, 0), (0, 1, 0, 3), (1, 3, 2, 4)},
      merge_regions(live_table(doc_id)))
assert_valid("transpose merged", doc_id)

# ===========================================================================
section("merge two tables")

d = Document()
make_table(d, [["Name", "Team", "Role"], ["ann", "Core", "dev"]])
d.add_paragraph("between")
make_table(d, [["Name", "Team", "Role"], ["bob", "Infra", "ops"], ["cy", "Infra", "sre"]],
           merges=[(1, 1, 2, 1)])
mid = upload(d)
status, body = table_op(mid, table_id(mid, 0), "merge", other_block_id=table_id(mid, 1))
check("merge tables: 200", status == 200, body)
rows = rows_of(mid)
check("merge tables: duplicate header dropped, rows appended",
      [r[0] for r in rows] == ["Name", "ann", "bob", "cy"], rows)
check("merge tables: second table removed", len(tables(mid)) == 1)
check("merge tables: the source's rowspan survives",
      (2, 1, 3, 1) in merge_regions(live_table(mid)), merge_regions(live_table(mid)))
assert_valid("merge tables", mid)

d = Document()
make_table(d, [["a", "b"]])
make_table(d, [["a", "b", "c"]])
mm = upload(d)
status, body = table_op(mm, table_id(mm, 0), "merge", other_block_id=table_id(mm, 1))
check("merge with different column counts refused", status == 400 and "column counts" in body["error"], body)

# ===========================================================================
section("set_content with shape changes (the table editor popup)")

doc_id = upload(merged_doc())
rows = [r[:] for r in rows_of(doc_id)]
del rows[2]                                   # user deleted the 'b2' row in the popup
table_op(doc_id, table_id(doc_id), "set_content", rows=rows)
regs = merge_regions(live_table(doc_id))
check("popup row delete: header colspan kept", (0, 1, 0, 2) in regs, regs)
check("popup row delete: 2x2 block kept", (2, 1, 3, 2) in regs, regs)
check("popup row delete: rowspan shrank", (1, 0, 2, 0) in regs, regs)
assert_valid("popup row delete", doc_id)

doc_id = upload(merged_doc())
rows = [r[:] for r in rows_of(doc_id)]
rows.insert(5, ["new", "row", "at", "end"])
rows[1][3] = "d1!"
table_op(doc_id, table_id(doc_id), "set_content", rows=rows)
check("popup row add + edit: merges kept", EXPECTED_REGIONS <= merge_regions(live_table(doc_id)),
      merge_regions(live_table(doc_id)))
check("popup row add + edit: values", rows_of(doc_id)[5] == ["new", "row", "at", "end"]
      and rows_of(doc_id)[1][3] == "d1!")
assert_valid("popup row add", doc_id)

doc_id = upload(merged_doc())
rows = [r[:] for r in rows_of(doc_id)]
rows = [r[:3] for r in rows]                  # user deleted the last column
table_op(doc_id, table_id(doc_id), "set_content", rows=rows)
check("popup column delete: merges kept", merge_regions(live_table(doc_id)) == EXPECTED_REGIONS,
      merge_regions(live_table(doc_id)))
check("popup column delete: 3 columns", all(len(r) == 3 for r in rows_of(doc_id)))
assert_valid("popup column delete", doc_id)

d = Document()
t = make_table(d, [["k", "v"], ["one", "1"], ["two", "2"], ["three", "3"]])
t.rows[1].cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xC0, 0, 0)
pid = upload(d)
rows = [r[:] for r in rows_of(pid)]
rows[1], rows[3] = rows[3], rows[1]           # user swapped rows with the popup arrows
table_op(pid, table_id(pid), "set_content", rows=rows)
check("popup row reorder: order", [r[0] for r in rows_of(pid)] == ["k", "three", "two", "one"])
moved = live_table(pid).rows[3].cells[0].paragraphs[0].runs[0]
check("popup row reorder: the red 'one' row moved as a unit (formatting kept)",
      moved.text == "one" and moved.font.color.rgb == RGBColor(0xC0, 0, 0))
assert_valid("popup reorder", pid)

# ===========================================================================
section("replace a table with a merged table (cross-document)")

src = Document()
make_table(src, [["S1", "S2", "S3"], ["x", "y", "z"], ["p", "q", "r"]], merges=[(0, 0, 0, 1), (1, 2, 2, 2)])
dst = Document()
make_table(dst, [["T1", "T2"], ["t", "u"]])
sid_, did_ = upload(src, "src.docx"), upload(dst, "dst.docx")
status, body = call("/replace", base_doc_id=did_, base_block_id=table_id(did_),
                    source_doc_id=sid_, source_block_id=table_id(sid_))
check("replace with merged table: 200", status == 200, body)
check("replace with merged table: values", rows_of(did_) == rows_of(sid_), rows_of(did_))
check("replace with merged table: source merges reproduced",
      merge_regions(live_table(did_)) == {(0, 0, 0, 1), (1, 2, 2, 2)}, merge_regions(live_table(did_)))
assert_valid("replace with merged", did_)

# ===========================================================================
section("HTML tables with rowspan/colspan, then edited")

html = """
<table>
  <tr><th rowspan="2">Region</th><th colspan="2">Sales</th></tr>
  <tr><th>Q1</th><th>Q2</th></tr>
  <tr><td>North</td><td>10</td><td>12</td></tr>
  <tr><td rowspan="2">South</td><td colspan="2">n/a</td></tr>
  <tr><td>7</td><td>9</td></tr>
</table>"""
status, body = call("/upload_html", html=html)
hid = body["doc_id"]
check("HTML merged table loads", status == 200 and len(tables(hid)) == 1)
regs = merge_regions(live_table(hid))
check("HTML merges become real merges", {(0, 0, 1, 0), (0, 1, 0, 2), (3, 0, 4, 0), (3, 1, 3, 2)} <= regs, regs)
table_op(hid, table_id(hid), "add_row", values=["West", "1", "2"])
table_op(hid, table_id(hid), "delete_row", index=2)
table_op(hid, table_id(hid), "add_column", header="Q3")
check("HTML table edited: rows", [r[0] for r in rows_of(hid)] == ["Region", "Region", "South", "South", "West"],
      rows_of(hid))
assert_valid("HTML table edited", hid)

# ===========================================================================
section("odd table structures")

# nested table inside a cell
d = Document()
outer = make_table(d, [["outer", "cell"], ["a", "b"]])
outer.cell(1, 1).add_table(rows=1, cols=2).cell(0, 0).text = "inner"
nid = upload(d)
check("nested table: parses", len(tables(nid)) == 1)
status, body = table_op(nid, table_id(nid), "add_row", values=["c", "d"])
check("nested table: add_row works", status == 200 and rows_of(nid)[-1] == ["c", "d"], body)
check("nested table: inner table kept",
      live_table(nid)._tbl.find(f".//{W('tc')}//{W('tbl')}") is not None)
assert_valid("nested table", nid)

# a row that starts late (gridBefore)
d = Document()
t = make_table(d, [["a", "b", "c"], ["d", "e", "f"]])
tr = t.rows[1]._tr
tr.remove(tr.findall(W("tc"))[0])
tr.insert(0, parse_xml(f'<w:trPr {nsdecls("w")}><w:gridBefore w:val="1"/></w:trPr>'))
gbid = upload(d)
check("gridBefore row: parses rectangular", rect(rows_of(gbid)), rows_of(gbid))
status, body = table_op(gbid, table_id(gbid), "add_row", values=["1", "2", "3"])
check("gridBefore row: add_row works", status == 200, body)
assert_valid("gridBefore", gbid)

# 1x1 table
d = Document()
make_table(d, [["solo"]])
sid1 = upload(d)
table_op(sid1, table_id(sid1), "add_column", header="two")
table_op(sid1, table_id(sid1), "add_row", values=["3", "4"])
check("1x1 grows to 2x2", rows_of(sid1) == [["solo", "two"], ["3", "4"]], rows_of(sid1))
assert_valid("1x1 grown", sid1)

# whole-row merge (one cell spanning every column)
d = Document()
make_table(d, [["Title", "", ""], ["a", "b", "c"]], merges=[(0, 0, 0, 2)])
wid = upload(d)
table_op(wid, table_id(wid), "delete_column", index=1)
check("full-width colspan survives a column delete",
      rows_of(wid)[0] == ["Title", "Title"] and (0, 0, 0, 1) in merge_regions(live_table(wid)),
      (rows_of(wid), merge_regions(live_table(wid))))
assert_valid("full-width colspan", wid)

# a rowspan covering the whole table height
d = Document()
make_table(d, [["side", "1"], ["", "2"], ["", "3"]], merges=[(0, 0, 2, 0)])
fh = upload(d)
table_op(fh, table_id(fh), "delete_row", index=0)
table_op(fh, table_id(fh), "delete_row", index=0)
check("full-height rowspan reduced to one row", rows_of(fh) == [["side", "3"]], rows_of(fh))
assert_valid("full-height rowspan", fh)

# ===========================================================================
section("undo across table edits")

doc_id = upload(merged_doc())
before = rows_of(doc_id)
for op, params in [("add_row", {"index": 2}), ("delete_column", {"index": 1}),
                   ("add_column", {"index": 0}), ("delete_row", {"index": 0})]:
    table_op(doc_id, table_id(doc_id), op, **params)
for _ in range(4):
    call("/undo", doc_id=doc_id)
check("4 table edits then 4 undos restore the exact table", rows_of(doc_id) == before)
check("undo restores the merges too", merge_regions(live_table(doc_id)) == EXPECTED_REGIONS)

# ===========================================================================
section("tracked-changes export of merged-table edits")

doc_id = upload(merged_doc())
rows = [r[:] for r in rows_of(doc_id)]
rows[1][1] = "b1 changed"
table_op(doc_id, table_id(doc_id), "set_content", rows=rows)
table_op(doc_id, table_id(doc_id), "add_row", values=["z0", "z1", "z2", "z3"])
r = client.get(f"/export/{doc_id}?track=1")
out = Document(io.BytesIO(r.data))
check("tracked export of merged table: opens", r.status_code == 200)
dels = "".join(t.text or "" for t in out.element.body.iter(W("delText")))
ins = "".join(t.text for tag in ("ins",) for el in out.element.body.iter(W(tag))
              for t in el.iter(W("t")) if t.text)
check("tracked export: edited cell shows old and new", "b1" in ins + dels and "changed" in ins, (dels, ins))
check("tracked export: new row is an inserted row",
      any(tr.find(f"{W('trPr')}/{W('ins')}") is not None for tr in out.element.body.iter(W("tr"))))
SAVED.append(("tracked merged", r.data))

# ===========================================================================
soffice = next((p for p in ("/usr/bin/soffice", "/usr/bin/libreoffice") if os.path.exists(p)), None)
section("open every result in LibreOffice")
if soffice:
    with tempfile.TemporaryDirectory() as tmp:
        paths = []
        for i, (name, data) in enumerate(SAVED):
            path = os.path.join(tmp, f"doc{i:03d}.docx")
            with open(path, "wb") as f:
                f.write(data)
            paths.append((name, path))
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp]
                       + [p for _, p in paths], capture_output=True, timeout=600)
        bad = [n for n, p in paths if not os.path.exists(p[:-5] + ".pdf")]
        check(f"LibreOffice opened all {len(paths)} result documents", not bad, bad)
else:
    print("  [SKIP] LibreOffice not installed")

print(f"\n=== {passed} passed, {failed} failed ===")
sys.exit(1 if failed else 0)
