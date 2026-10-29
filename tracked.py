"""
Tracked-changes export: show every edit as a Word revision.

`build_tracked_document(original_bytes, current_document)` returns a NEW
python-docx Document that looks like the current document, but with every
difference from the original marked as a Word tracked change:

  - deleted text / paragraphs / table rows  -> <w:del> (struck through)
  - inserted text / paragraphs / table rows -> <w:ins> (underlined)
  - a paragraph whose style changed (e.g. heading level) -> <w:pPrChange>

so a reviewer can open it in Word and Accept / Reject each change. "Accept
all" yields exactly the current document; "Reject all" yields the original.

How the diff works
------------------
It compares *content*, not the edit history, so it is correct no matter what
sequence of operations (or undos) produced the current document:

1. Body level: both documents' top-level elements are reduced to signatures
   (paragraph text + style, table cell texts) and aligned with difflib.
2. Changed paragraphs that line up are diffed word-by-word, so a one-word fix
   shows as one struck word + one inserted word, not a whole rewritten
   paragraph. The current paragraph's own runs (and their formatting) are
   used for kept and inserted text; deleted text keeps the original run's
   formatting.
3. Changed tables with the same column count are diffed row-by-row, and
   changed rows cell-by-cell using the same paragraph logic.

Deliberate scope limits: a moved block shows as delete + insert (Word's
move-tracking isn't generated), and formatting-only changes to the same text
are not marked.
"""
import io
import re
from copy import deepcopy
from datetime import datetime, timezone
from difflib import SequenceMatcher

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from fidelity import REL_ATTRS

W = lambda tag: qn("w:" + tag)

# paragraphs pairs less similar than this are shown as delete + insert of the
# whole paragraph - a word-level diff of two unrelated sentences is unreadable
MIN_INLINE_RATIO = 0.3

_TOKEN_RE = re.compile(r"\w+|\s+|[^\w\s]", re.UNICODE)

# run children that are "atoms" in the diff: each one is one placeholder char
_ATOM_CHARS = {
    W("tab"): "\t", W("br"): "\n", W("cr"): "\n", W("noBreakHyphen"): "-",
    W("softHyphen"): "­", W("sym"): "￼",
}
_IGNORABLE_RUN_CHILDREN = {W("rPr"), W("lastRenderedPageBreak")}
_SIMPLE_P_CHILDREN = {W("pPr"), W("r"), W("bookmarkStart"), W("bookmarkEnd"), W("proofErr")}


class _Rev:
    """Revision id/author/date factory."""

    def __init__(self, start_id, author, date):
        self.next_id = start_id
        self.author = author
        self.date = date

    def mark(self, tag):
        el = OxmlElement("w:" + tag)
        el.set(W("id"), str(self.next_id))
        el.set(W("author"), self.author)
        el.set(W("date"), self.date)
        self.next_id += 1
        return el


def build_tracked_document(original_bytes, current_document, author="docagent", date=None):
    buf = io.BytesIO()
    current_document.save(buf)
    out = Document(io.BytesIO(buf.getvalue()))      # never touch the live doc
    orig = Document(io.BytesIO(original_bytes))

    date = date or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev = _Rev(_max_w_id(out.element.body) + 1000, author, date)
    valid_rids = set(out.part.rels.keys())

    out_body = out.element.body
    sect = out_body.find(W("sectPr"))
    orig_children = [c for c in orig.element.body if c.tag != W("sectPr")]
    out_children = [c for c in out_body if c.tag != W("sectPr")]

    ctx = _Ctx(rev, valid_rids)
    merged = ctx.diff_sequence(orig_children, out_children)

    for c in out_children:
        out_body.remove(c)
    for el in merged:
        if sect is not None:
            sect.addprevious(el)
        else:
            out_body.append(el)
    return out


class _Ctx:
    def __init__(self, rev, valid_rids):
        self.rev = rev
        self.valid_rids = valid_rids

    # -- sequence alignment (body, or a table cell's content) ----------------
    def diff_sequence(self, old, new):
        sm = SequenceMatcher(None, [_sig(e) for e in old], [_sig(e) for e in new],
                             autojunk=False)
        result = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                result.extend(new[j1:j2])
            elif tag == "delete":
                result.extend(self.deleted(e) for e in old[i1:i2])
            elif tag == "insert":
                result.extend(self.inserted(e) for e in new[j1:j2])
            else:
                result.extend(self._replace_block(old[i1:i2], new[j1:j2]))
        return result

    def _replace_block(self, old, new):
        """
        A run of changed elements. Align them by KIND first (paragraph vs
        table), and within a kind by text similarity, so an edited paragraph is
        paired with its own earlier version rather than with whatever happens
        to sit at the same position. Unpaired elements are plain deletes /
        inserts; each pair is diffed in detail.
        """
        out = []
        for o, n in _align(old, new):
            if n is None:
                out.append(self.deleted(o))
            elif o is None:
                out.append(self.inserted(n))
            elif o.tag == W("p"):
                out.extend(self.diff_paragraph(o, n))
            elif o.tag == W("tbl"):
                out.extend(self.diff_table(o, n))
            else:
                out.extend([self.deleted(o), self.inserted(n)])
        return out

    # -- whole-element marking --------------------------------------------
    def inserted(self, el):
        _mark_container(el, "ins", self.rev)
        return el

    def deleted(self, el):
        el = deepcopy(el)
        _strip_for_deleted_copy(el, self.valid_rids)
        _mark_container(el, "del", self.rev)
        return el

    # -- paragraph word diff ------------------------------------------------
    def diff_paragraph(self, old_p, new_p, force_inline=False):
        if not (_simple_paragraph(old_p) and _simple_paragraph(new_p)):
            return [self.deleted(old_p), self.inserted(new_p)]
        old_atoms, new_atoms = _atoms(old_p), _atoms(new_p)
        old_text = "".join(a[1] for a in old_atoms)
        new_text = "".join(a[1] for a in new_atoms)
        old_toks, new_toks = _tokens(old_text), _tokens(new_text)
        sm = SequenceMatcher(None, [t for t, _, _ in old_toks],
                             [t for t, _, _ in new_toks], autojunk=False)
        if not force_inline and old_text.strip() and new_text.strip() \
                and sm.ratio() < MIN_INLINE_RATIO:
            return [self.deleted(old_p), self.inserted(new_p)]

        p = OxmlElement("w:p")
        new_ppr = new_p.find(W("pPr"))
        if new_ppr is not None:
            p.append(deepcopy(new_ppr))
        self._style_change(old_p, p)
        for b in new_p:
            if b.tag in (W("bookmarkStart"), W("bookmarkEnd")):
                p.append(deepcopy(b))

        def span(toks, a, b):
            return (toks[a][1], toks[b - 1][2]) if b > a else (0, 0)

        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            o0, o1 = span(old_toks, i1, i2)
            n0, n1 = span(new_toks, j1, j2)
            if tag == "equal":
                _emit_runs(p, new_atoms[n0:n1], None, self.rev)
            if tag in ("delete", "replace"):
                _emit_runs(p, old_atoms[o0:o1], "del", self.rev)
            if tag in ("insert", "replace"):
                _emit_runs(p, new_atoms[n0:n1], "ins", self.rev)
        return [p]

    def _style_change(self, old_p, p):
        """Record a paragraph style change (e.g. Heading 2 -> Heading 1)."""
        if _pstyle(old_p) == _pstyle(p):
            return
        ppr = p.find(W("pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            p.insert(0, ppr)
        change = self.rev.mark("pPrChange")
        old_ppr = deepcopy(old_p.find(W("pPr"))) if old_p.find(W("pPr")) is not None \
            else OxmlElement("w:pPr")
        for c in list(old_ppr):
            if c.tag in (W("rPr"), W("sectPr"), W("pPrChange")):
                old_ppr.remove(c)
        change.append(old_ppr)
        ppr.append(change)              # pPrChange is always pPr's last child

    # -- table row/cell diff ----------------------------------------------
    def diff_table(self, old_t, new_t):
        old_rows, new_rows = old_t.findall(W("tr")), new_t.findall(W("tr"))
        widths = {len(r.findall(W("tc"))) for r in old_rows + new_rows}
        if len(widths) != 1:
            # column structure changed - show it as old table out, new table in
            return [self.deleted(old_t), self.inserted(new_t)]

        sm = SequenceMatcher(None, [_row_sig(r) for r in old_rows],
                             [_row_sig(r) for r in new_rows], autojunk=False)
        rows = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                rows.extend(new_rows[j1:j2])
                continue
            old_part, new_part = old_rows[i1:i2], new_rows[j1:j2]
            k = min(len(old_part), len(new_part)) if tag == "replace" else 0
            for o, n in zip(old_part[:k], new_part[:k]):
                rows.append(self._diff_row(o, n))
            rows.extend(self.deleted(o) for o in old_part[k:])
            rows.extend(self.inserted(n) for n in new_part[k:])

        for r in new_rows:
            new_t.remove(r)
        for r in rows:
            new_t.append(r)
        return [new_t]

    def _diff_row(self, old_r, new_r):
        for old_tc, new_tc in zip(old_r.findall(W("tc")), new_r.findall(W("tc"))):
            old_content = [c for c in old_tc if c.tag != W("tcPr")]
            new_content = [c for c in new_tc if c.tag != W("tcPr")]
            if [_sig(e) for e in old_content] == [_sig(e) for e in new_content]:
                continue
            if len(old_content) == len(new_content) == 1 and \
                    old_content[0].tag == new_content[0].tag == W("p"):
                # a single-value cell ("3" -> "5"): show the edit inline in
                # the cell, however different the values are
                merged = self.diff_paragraph(old_content[0], new_content[0], force_inline=True)
            else:
                merged = self.diff_sequence(old_content, new_content)
            for c in new_content:
                new_tc.remove(c)
            for c in merged:
                new_tc.append(c)
            if merged and merged[-1].tag != W("p"):
                new_tc.append(OxmlElement("w:p"))   # a cell must end with a paragraph
        return new_r


# ---------------------------------------------------------------------------
# similarity alignment of a changed run of elements
# ---------------------------------------------------------------------------

# minimum similarity for two changed elements to count as "the same block,
# edited" rather than one deleted and an unrelated one inserted
_PAIR_MIN = {W("p"): 0.5, W("tbl"): 0.3}
_ALIGN_MAX_CELLS = 40_000     # beyond this, fall back to cheaper kind-only pairing


def _similarity(o, n):
    if o.tag != n.tag or o.tag not in _PAIR_MIN:
        return None
    a, b = _text(o), _text(n)
    if not a and not b:
        score = 1.0 if _sig(o) == _sig(n) else 0.0
    else:
        score = SequenceMatcher(None, a, b, autojunk=False).quick_ratio()
    return score if score >= _PAIR_MIN[o.tag] else None


def _align(old, new):
    """
    Order-preserving alignment maximising total similarity of paired
    elements (a small Needleman-Wunsch with zero gap cost). Returns a list of
    (old_or_None, new_or_None) in document order, deletions before insertions.
    """
    n_old, n_new = len(old), len(new)
    if n_old * n_new > _ALIGN_MAX_CELLS:
        sim = lambda i, j: 1.0 if old[i].tag == new[j].tag else None
    else:
        cache = {}

        def sim(i, j):
            if (i, j) not in cache:
                cache[(i, j)] = _similarity(old[i], new[j])
            return cache[(i, j)]

    # best[i][j] = best score aligning old[i:] with new[j:]
    best = [[0.0] * (n_new + 1) for _ in range(n_old + 1)]
    for i in range(n_old - 1, -1, -1):
        for j in range(n_new - 1, -1, -1):
            s = sim(i, j)
            pair = (s + best[i + 1][j + 1]) if s is not None else -1.0
            best[i][j] = max(pair, best[i + 1][j], best[i][j + 1])

    out, i, j = [], 0, 0
    while i < n_old and j < n_new:
        s = sim(i, j)
        if s is not None and best[i][j] == s + best[i + 1][j + 1]:
            out.append((old[i], new[j]))
            i += 1
            j += 1
        elif best[i][j] == best[i + 1][j]:
            out.append((old[i], None))
            i += 1
        else:
            out.append((None, new[j]))
            j += 1
    out.extend((o, None) for o in old[i:])
    out.extend((None, n) for n in new[j:])
    return out


# ---------------------------------------------------------------------------
# signatures
# ---------------------------------------------------------------------------

def _text(el):
    parts = []
    for node in el.iter():
        if node.tag == W("t"):
            parts.append(node.text or "")
        elif node.tag in _ATOM_CHARS:
            parts.append(_ATOM_CHARS[node.tag])
        elif node.tag == W("p"):
            parts.append(" ")
    return "".join(parts)


def _pstyle(p):
    s = p.find(W("pPr") + "/" + W("pStyle"))
    return s.get(W("val")) if s is not None else None


def _sig(el):
    if el.tag == W("p"):
        rids = tuple(n.get(a) for n in el.iter() for a in REL_ATTRS if n.get(a))
        return ("p", _text(el), _pstyle(el), rids)
    if el.tag == W("tbl"):
        return ("tbl", tuple(_row_sig(r) for r in el.findall(W("tr"))))
    return ("x", el.tag, _text(el))


def _row_sig(tr):
    return tuple(_text(tc) for tc in tr.findall(W("tc")))


# ---------------------------------------------------------------------------
# paragraph atoms (char-level view of runs, keeping each char's formatting)
# ---------------------------------------------------------------------------

def _simple_paragraph(p):
    """Only plain text runs - no images, fields, hyperlinks, comments..."""
    for child in p:
        if child.tag not in _SIMPLE_P_CHILDREN:
            return False
        if child.tag == W("r"):
            for rc in child:
                if rc.tag not in _IGNORABLE_RUN_CHILDREN and rc.tag != W("t") \
                        and rc.tag not in _ATOM_CHARS:
                    return False
    return True


def _atoms(p):
    """[(rPr element or None, char, element-or-None)] for every char in p."""
    atoms = []
    for r in p.findall(W("r")):
        rpr = r.find(W("rPr"))
        for rc in r:
            if rc.tag == W("t"):
                atoms.extend((rpr, ch, None) for ch in (rc.text or ""))
            elif rc.tag in _ATOM_CHARS:
                atoms.append((rpr, _ATOM_CHARS[rc.tag], rc))
    return atoms


def _tokens(text):
    return [(m.group(), m.start(), m.end()) for m in _TOKEN_RE.finditer(text)]


def _emit_runs(p, atoms, kind, rev):
    """Append atoms to p as runs (grouped by formatting), wrapped in w:ins/w:del."""
    if not atoms:
        return
    wrapper = rev.mark(kind) if kind else None
    target = wrapper if wrapper is not None else p
    run, cur_rpr, buf = None, object(), []

    def flush_text():
        if buf and run is not None:
            t = OxmlElement("w:delText" if kind == "del" else "w:t")
            t.text = "".join(buf)
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            run.append(t)
            buf.clear()

    for rpr, ch, el in atoms:
        if rpr is not cur_rpr or run is None:
            flush_text()
            run = OxmlElement("w:r")
            if rpr is not None:
                run.append(deepcopy(rpr))
            target.append(run)
            cur_rpr = rpr
        if el is None:
            buf.append(ch)
        else:
            flush_text()
            run.append(deepcopy(el))
    flush_text()
    if wrapper is not None:
        p.append(wrapper)


# ---------------------------------------------------------------------------
# whole-element marking
# ---------------------------------------------------------------------------

_RUN_CONTAINERS = {W("hyperlink"), W("smartTag"), W("fldSimple"), W("customXml")}


def _mark_container(el, kind, rev):
    """Mark every paragraph (and table row) inside el as inserted/deleted."""
    if el.tag == W("tr"):
        _mark_row(el, kind, rev)
    for tr in el.iter(W("tr")):
        if tr is not el:
            _mark_row(tr, kind, rev)
    paras = [el] if el.tag == W("p") else list(el.iter(W("p")))
    for p in paras:
        _mark_paragraph(p, kind, rev)


def _mark_row(tr, kind, rev):
    trpr = tr.find(W("trPr"))
    if trpr is None:
        trpr = OxmlElement("w:trPr")
        prex = tr.find(W("tblPrEx"))
        if prex is not None:
            prex.addnext(trpr)
        else:
            tr.insert(0, trpr)
    change = trpr.find(W("trPrChange"))
    mark = rev.mark(kind)
    if change is not None:
        change.addprevious(mark)
    else:
        trpr.append(mark)


def _mark_paragraph(p, kind, rev):
    _wrap_runs(p, kind, rev)
    # the paragraph mark itself
    ppr = p.find(W("pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    rpr = ppr.find(W("rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        anchor = ppr.find(W("sectPr"))
        if anchor is None:
            anchor = ppr.find(W("pPrChange"))
        if anchor is not None:
            anchor.addprevious(rpr)
        else:
            ppr.append(rpr)
    rpr.insert(0, rev.mark(kind))


def _wrap_runs(container, kind, rev):
    """Group consecutive w:r children into one w:ins/w:del; recurse into links."""
    group = None
    for child in list(container):
        if child.tag == W("r"):
            if kind == "del":
                _to_deleted_text(child)
            if group is None:
                group = rev.mark(kind)
                child.addprevious(group)
            group.append(child)
        else:
            group = None
            if child.tag in _RUN_CONTAINERS:
                _wrap_runs(child, kind, rev)


def _to_deleted_text(r):
    for t in r.iter(W("t")):
        t.tag = W("delText")
    for t in r.iter(W("instrText")):
        t.tag = W("delInstrText")


def _strip_for_deleted_copy(el, valid_rids):
    """A deleted copy re-enters the document next to the live content: drop
    its bookmarks (names must be unique) and anything pointing at a part the
    output no longer has."""
    for tag in ("bookmarkStart", "bookmarkEnd", "commentRangeStart",
                "commentRangeEnd"):
        for b in list(el.iter(W(tag))):
            b.getparent().remove(b)
    for node in list(el.iter()):
        if any(node.get(a) and node.get(a) not in valid_rids for a in REL_ATTRS):
            r = node
            while r is not None and r.tag != W("r"):
                r = r.getparent()
            target = r if r is not None else node
            if target.getparent() is not None:
                target.getparent().remove(target)
    for sect in list(el.iter(W("sectPr"))):
        sect.getparent().remove(sect)


def _max_w_id(root):
    best = 0
    for node in root.iter():
        v = node.get(W("id"))
        if v and v.lstrip("-").isdigit():
            best = max(best, int(v))
    return best
