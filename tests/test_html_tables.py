"""
Regression tests for HTML table parsing, focused on rowspan/colspan.

Run from the project root:  python tests/test_html_tables.py

These guard the bug where rowspan was ignored, which shifted cells out of
alignment and produced ragged (non-rectangular) grids that then errored when
building the Word table. Every case here asserts the grid is rectangular and
that spanned cells land in the right positions.
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bs4 import BeautifulSoup
from app import _html_table_to_rows, html_to_document


passed = 0
failed = 0


def check(name, cond):
    global passed, failed
    if cond:
        passed += 1
        print(f"  [PASS] {name}")
    else:
        failed += 1
        print(f"  [FAIL] {name}")


def rows_of(html):
    soup = BeautifulSoup(html, "lxml")
    matrix, _merges = _html_table_to_rows(soup.find("table"))
    return matrix


def merges_of(html):
    soup = BeautifulSoup(html, "lxml")
    _matrix, merges = _html_table_to_rows(soup.find("table"))
    return merges


def is_rectangular(rows):
    return len(rows) > 0 and len({len(r) for r in rows}) == 1


print("\n=== HTML table rowspan/colspan ===\n")

# 1. rowspan: anchor holds the value, covered cell is empty, merge recorded
r = rows_of(
    "<table>"
    "<tr><th>Region</th><th>Q1</th><th>Q2</th></tr>"
    "<tr><td rowspan='2'>North</td><td>10</td><td>20</td></tr>"
    "<tr><td>30</td><td>40</td></tr>"
    "</table>"
)
m = merges_of(
    "<table>"
    "<tr><th>Region</th><th>Q1</th><th>Q2</th></tr>"
    "<tr><td rowspan='2'>North</td><td>10</td><td>20</td></tr>"
    "<tr><td>30</td><td>40</td></tr>"
    "</table>"
)
check("rowspan: rectangular grid", is_rectangular(r))
check("rowspan: anchor holds value", r[1][0] == "North")
check("rowspan: covered cell left empty (no duplication)", r[2][0] == "")
check("rowspan: following cells aligned", r[2][1] == "30" and r[2][2] == "40")
check("rowspan: merge region recorded", (1, 0, 2, 0) in m)

# 2. colspan: anchor + empty, merge recorded
r = rows_of("<table><tr><th colspan='2'>Wide</th><th>C</th></tr><tr><td>a</td><td>b</td><td>c</td></tr></table>")
m = merges_of("<table><tr><th colspan='2'>Wide</th><th>C</th></tr><tr><td>a</td><td>b</td><td>c</td></tr></table>")
check("colspan: rectangular grid", is_rectangular(r))
check("colspan: anchor holds value, next cell empty", r[0][0] == "Wide" and r[0][1] == "")
check("colspan: merge region recorded", (0, 0, 0, 1) in m)

# 3. the user's exact table: 5 cols, col1 + cols 3-5 rowspan 2, col2 splits
USER = (
    "<table>"
    "<tr><th>Column 1</th><th>Column 2</th><th>Column 3</th><th>Column 4</th><th>Column 5</th></tr>"
    "<tr><td rowspan='2'>Value 1</td><td>Column 2 - Row 1</td>"
    "<td rowspan='2'>Value 3</td><td rowspan='2'>Value 4</td><td rowspan='2'>Value 5</td></tr>"
    "<tr><td>Column 2 - Row 2</td></tr>"
    "</table>"
)
r = rows_of(USER)
m = merges_of(USER)
check("user table: 3 rows x 5 cols", len(r) == 3 and all(len(x) == 5 for x in r))
check("user table: Value 1 anchored once", r[1][0] == "Value 1" and r[2][0] == "")
check("user table: Column 2 splits into two rows",
      r[1][1] == "Column 2 - Row 1" and r[2][1] == "Column 2 - Row 2")
check("user table: 4 vertical merges (cols 0,2,3,4)",
      set(m) == {(1, 0, 2, 0), (1, 2, 2, 2), (1, 3, 2, 3), (1, 4, 2, 4)})
# and the built docx actually merges them
doc = html_to_document(USER)
t = doc.tables[0]
check("user table: (1,0)/(2,0) are the SAME merged cell", t.cell(1, 0)._tc is t.cell(2, 0)._tc)
check("user table: Column 2 cells NOT merged", t.cell(1, 1)._tc is not t.cell(2, 1)._tc)

# 4. rowspan across 3 rows plus a colspan mid-table
r = rows_of(
    "<table>"
    "<tr><td rowspan='3'>Tall</td><td>a1</td><td>a2</td></tr>"
    "<tr><td colspan='2'>wide</td></tr>"
    "<tr><td>c1</td><td>c2</td></tr>"
    "</table>"
)
check("triple rowspan: rectangular", is_rectangular(r))
check("triple rowspan: anchored once at top", r[0][0] == "Tall" and r[1][0] == "" and r[2][0] == "")

# 5. oversized rowspan (exceeds remaining rows) must not error
r = rows_of("<table><tr><td rowspan='9'>X</td><td>a</td></tr><tr><td>b</td></tr></table>")
check("oversized rowspan: rectangular, clamped", is_rectangular(r) and len(r) == 2)

# 6. thead/tbody wrappers (Froala emits these) still parse spans + merge
doc = html_to_document(
    "<table><thead><tr><th>H1</th><th>H2</th></tr></thead>"
    "<tbody><tr><td rowspan='2'>m</td><td>x</td></tr><tr><td>y</td></tr></tbody></table>"
)
built_t = doc.tables[0]
check("thead/tbody: (1,0)/(2,0) merged", built_t.cell(1, 0)._tc is built_t.cell(2, 0)._tc)

# 7. plain table unaffected, no merges
r = rows_of("<table><tr><th>N</th><th>V</th></tr><tr><td>a</td><td>1</td></tr></table>")
m = merges_of("<table><tr><th>N</th><th>V</th></tr><tr><td>a</td><td>1</td></tr></table>")
check("plain table still parses", r == [["N", "V"], ["a", "1"]])
check("plain table has no merges", m == [])

print(f"\n=== {passed} passed, {failed} failed ===")
sys.exit(1 if failed else 0)
