"""
Formatting-fidelity helpers for docagent.

Two jobs live here, both about not losing formatting when content moves:

1. CROSS-DOCUMENT IMPORT (`Importer`)
   A paragraph or table's XML is not self-contained. It points outside itself:
     - images and hyperlinks by relationship id (r:embed="rId7"),
     - list bullets/numbers by numbering id (<w:numId w:val="3"/>),
     - styles by style id (<w:pStyle w:val="ListBullet"/>).
   Those ids only mean something inside the document they came from. A plain
   deepcopy into another document leaves them pointing at nothing - or worse,
   at a *different* image / list / style that happens to share the id. The
   Importer clones an element and re-points every such reference at an
   equivalent definition in the target document, copying the definition over
   when the target doesn't have it.

2. RICH CONTENT COPY (`copy_paragraph_content`)
   A same-type replace used to write the source's plain text into the target's
   first run, so a bold word mid-sentence in the source came out plain. Now
   the source's runs are copied across (with their images/links imported),
   and each run is restyled as: the TARGET's base run formatting (font, size,
   colour) + the SOURCE's emphasis (bold, italic, underline, strike,
   super/subscript, caps, highlight, character style). The target block still
   looks like it belongs in the target document, but the meaning-carrying
   emphasis survives.
"""
import io
import logging
import secrets
from copy import deepcopy

from lxml import etree

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

log = logging.getLogger("docagent")

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_ATTRS = (f"{{{R_NS}}}embed", f"{{{R_NS}}}link", f"{{{R_NS}}}id")

# Elements that reference document-local parts we deliberately do NOT carry
# across (comments, footnotes/endnotes belong to the source document). They are
# dropped from imported content rather than left dangling, which Word reports
# as a corrupt file.
_DROP_ON_IMPORT = (
    "w:commentRangeStart", "w:commentRangeEnd", "w:commentReference",
    "w:footnoteReference", "w:endnoteReference",
)


class ImportError_(ValueError):
    """Content can't be carried into the target document safely."""


class Importer:
    """
    Clone XML from `source` (a python-docx Document) for insertion into
    `target`, re-pointing images, hyperlinks, list numbering and styles.

    One Importer per operation: it caches mappings so importing ten list items
    from the same source list maps them all onto ONE new list in the target
    (so numbering continues 1, 2, 3... instead of restarting per item).
    """

    def __init__(self, source, target):
        self.source = source
        self.target = target
        self.same = source.part is target.part
        self._rid_map = {}
        self._num_map = {}
        self._styles_done = set()

    # -- public ------------------------------------------------------------
    def import_element(self, element):
        clone = deepcopy(element)
        if self.same:
            return clone
        self._drop_unportable(clone)
        self._remap_rels(clone)
        self._remap_numbering(clone)
        self._ensure_styles(clone)
        return clone

    # -- relationships: images, hyperlinks -----------------------------------
    def _remap_rels(self, root):
        for node in root.iter():
            for attr in REL_ATTRS:
                rid = node.get(attr)
                if rid:
                    node.set(attr, self._map_rid(rid))

    def _map_rid(self, rid):
        if rid in self._rid_map:
            return self._rid_map[rid]
        src_part, tgt_part = self.source.part, self.target.part
        rel = src_part.rels.get(rid)
        if rel is None:
            raise ImportError_(f"source content references a missing relationship ({rid})")
        if rel.is_external:
            new = tgt_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        elif rel.reltype == RT.IMAGE:
            # get_or_add_image de-duplicates by content hash and gives the image
            # a fresh part name in the target package - no name collisions.
            new, _ = tgt_part.get_or_add_image(io.BytesIO(rel.target_part.blob))
        else:
            kind = rel.reltype.rsplit("/", 1)[-1]
            raise ImportError_(
                f"the selected content contains a '{kind}' object (e.g. a chart or "
                f"embedded file) that can't be copied between documents yet"
            )
        self._rid_map[rid] = new
        return new

    def _drop_unportable(self, root):
        # a section break inside a paragraph carries header/footer references
        # to the source's own parts - the target keeps its own layout instead
        for sect in root.iter(qn("w:sectPr")):
            sect.getparent().remove(sect)
        for tag in _DROP_ON_IMPORT:
            for el in list(root.iter(qn(tag))):
                parent = el.getparent()
                # a footnote/comment reference lives alone in its own run
                if parent is not None and parent.tag == qn("w:r") and \
                        len([c for c in parent if c.tag != qn("w:rPr")]) == 1:
                    gp = parent.getparent()
                    if gp is not None:
                        gp.remove(parent)
                        continue
                if parent is not None:
                    parent.remove(el)

    # -- numbering: bullets / numbered lists ----------------------------------
    def _remap_numbering(self, root):
        for num_id_el in root.iter(qn("w:numId")):
            val = num_id_el.get(qn("w:val"))
            if val and val != "0":
                num_id_el.set(qn("w:val"), self._map_num(val))

    def _src_numbering(self):
        try:
            return self.source.part.numbering_part.element
        except (KeyError, NotImplementedError):
            return None

    def _map_num(self, src_num_id):
        if src_num_id in self._num_map:
            return self._num_map[src_num_id]
        src = self._src_numbering()
        num = None if src is None else _first(src, f'./w:num[@w:numId="{src_num_id}"]')
        if num is None:
            # dangling in the source too - neutralise rather than point at a
            # random list in the target
            self._num_map[src_num_id] = "0"
            return "0"
        abs_id = _first(num, "./w:abstractNumId").get(qn("w:val"))
        abstract = _first(src, f'./w:abstractNum[@w:abstractNumId="{abs_id}"]')

        tgt = self.target.part.numbering_part.element   # created if missing
        new_abs_id = str(_max_attr(tgt, "./w:abstractNum", "w:abstractNumId") + 1)
        new_num_id = str(_max_attr(tgt, "./w:num", "w:numId") + 1)

        if abstract is not None:
            abs_clone = deepcopy(abstract)
            abs_clone.set(qn("w:abstractNumId"), new_abs_id)
            # nsid is Word's identity for a list definition - a fresh one stops
            # Word from merging this list with an unrelated one in the target
            nsid = _first(abs_clone, "./w:nsid")
            if nsid is not None:
                nsid.set(qn("w:val"), secrets.token_hex(4).upper())
            # schema order: every abstractNum precedes every num
            first_num = _first(tgt, "./w:num")
            if first_num is not None:
                first_num.addprevious(abs_clone)
            else:
                _insert_before_cleanup(tgt, abs_clone)
        num_clone = deepcopy(num)
        num_clone.set(qn("w:numId"), new_num_id)
        _first(num_clone, "./w:abstractNumId").set(qn("w:val"), new_abs_id)
        _insert_before_cleanup(tgt, num_clone)

        self._num_map[src_num_id] = new_num_id
        return new_num_id

    # -- styles ------------------------------------------------------------
    def _ensure_styles(self, root):
        for tag in ("w:pStyle", "w:rStyle", "w:tblStyle"):
            for el in root.iter(qn(tag)):
                self._ensure_style(el.get(qn("w:val")))

    def _ensure_style(self, style_id):
        """
        Copy a style definition the target lacks. When the target already has
        a style with that id, the TARGET's definition wins (Word's "use
        destination styles" paste) so imported headings match the base doc.
        """
        if not style_id or style_id in self._styles_done:
            return
        self._styles_done.add(style_id)
        tgt_styles = self.target.styles.element
        if _first(tgt_styles, f'./w:style[@w:styleId="{style_id}"]') is not None:
            return
        src_style = _first(self.source.styles.element, f'./w:style[@w:styleId="{style_id}"]')
        if src_style is None:
            return
        clone = deepcopy(src_style)
        self._remap_numbering(clone)      # list styles carry their own numPr
        tgt_styles.append(clone)
        for tag in ("w:basedOn", "w:link", "w:next"):
            ref = _first(clone, f"./{tag}")
            if ref is not None:
                self._ensure_style(ref.get(qn("w:val")))


# ---------------------------------------------------------------------------
# Rich paragraph copy (replace that keeps per-word emphasis)
# ---------------------------------------------------------------------------

# emphasis = the formatting that carries meaning; taken from the SOURCE run
EMPHASIS_TAGS = ("w:rStyle", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
                 "w:strike", "w:dstrike", "w:highlight", "w:u", "w:vertAlign")

# canonical child order of <w:rPr> (CT_RPr is a sequence; Word is strict)
_RPR_ORDER = ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
              "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
              "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
              "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
              "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
              "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
              "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange"]
_RPR_RANK = {qn(t): i for i, t in enumerate(_RPR_ORDER)}

# children of a run that are "just text" - a run made only of these can be
# rewritten freely; anything else (drawing, field char, footnote ref...) is
# content that plain-text editing must leave alone
TEXTUAL_RUN_CHILDREN = {qn(t) for t in (
    "w:rPr", "w:t", "w:tab", "w:br", "w:cr", "w:noBreakHyphen", "w:softHyphen",
    "w:lastRenderedPageBreak", "w:sym")}


def is_text_run(r):
    return all(child.tag in TEXTUAL_RUN_CHILDREN for child in r)


def sort_rpr(rpr):
    children = sorted(list(rpr), key=lambda c: _RPR_RANK.get(c.tag, len(_RPR_RANK) - 1))
    for c in list(rpr):
        rpr.remove(c)
    for c in children:
        rpr.append(c)


def _base_rpr(p_el):
    """Target's base run formatting: rPr of its first text-bearing run."""
    for r in p_el.iter(qn("w:r")):
        if r.find(qn("w:t")) is not None:
            rpr = r.find(qn("w:rPr"))
            return deepcopy(rpr) if rpr is not None else OxmlElement("w:rPr")
    return None


def _restyle_run(r, base_rpr):
    if base_rpr is None:
        return                          # nothing to adopt - keep source formatting
    src_rpr = r.find(qn("w:rPr"))
    new = deepcopy(base_rpr)
    for tag in EMPHASIS_TAGS:
        for el in new.findall(qn(tag)):
            new.remove(el)
    if src_rpr is not None:
        for tag in EMPHASIS_TAGS:
            el = src_rpr.find(qn(tag))
            if el is not None:
                new.append(deepcopy(el))
        r.remove(src_rpr)
    sort_rpr(new)
    if len(new):
        r.insert(0, new)


def copy_paragraph_content(target_p, source_p, importer):
    """
    Replace target_p's inline content with source_p's, keeping target_p's
    paragraph properties (style, spacing, numbering) and base run look, but
    the source's per-run emphasis, inline images and hyperlinks.
    """
    t_el, s_el = target_p._p if hasattr(target_p, "_p") else target_p, \
        source_p._p if hasattr(source_p, "_p") else source_p
    base = _base_rpr(t_el)
    for child in list(t_el):
        if child.tag != qn("w:pPr"):
            t_el.remove(child)
    for child in s_el:
        if child.tag in (qn("w:pPr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
            continue                     # bookmarks would duplicate names
        clone = importer.import_element(child)
        runs = [clone] if clone.tag == qn("w:r") else list(clone.iter(qn("w:r")))
        for r in runs:
            _restyle_run(r, base)
        t_el.append(clone)


def copy_table_content(target_table, source_table, importer):
    """
    After the target table has been resized to the source's shape (and filled
    with plain text), upgrade each cell to the source cell's rich content.
    Merged cells are visited once.
    """
    seen = set()
    t_rows, s_rows = target_table.rows, source_table.rows
    for r_idx in range(min(len(t_rows), len(s_rows))):
        t_cells, s_cells = t_rows[r_idx].cells, s_rows[r_idx].cells
        for c_idx in range(min(len(t_cells), len(s_cells))):
            tc, sc = t_cells[c_idx], s_cells[c_idx]
            if tc._tc in seen:
                continue
            seen.add(tc._tc)
            if sc._tc.find(qn("w:tbl")) is not None:
                continue                 # nested tables: keep the plain-text copy
            s_paras = sc.paragraphs
            t_paras = tc.paragraphs
            template = t_paras[0]._p
            for extra in t_paras[1:]:
                extra._p.getparent().remove(extra._p)
            prev = template
            for i, sp in enumerate(s_paras):
                if i == 0:
                    tp_el = template
                else:
                    tp_el = OxmlElement("w:p")
                    ppr = template.find(qn("w:pPr"))
                    if ppr is not None:
                        tp_el.append(deepcopy(ppr))
                    # keep the template's base run look for the new paragraph
                    tr = template.find(qn("w:r"))
                    if tr is not None:
                        stub = deepcopy(tr)
                        for c in list(stub):
                            if c.tag != qn("w:rPr"):
                                stub.remove(c)
                        stub.append(OxmlElement("w:t"))
                        tp_el.append(stub)
                    prev.addnext(tp_el)
                copy_paragraph_content(tp_el, sp, importer)
                prev = tp_el


def set_plain_text(paragraph, new_text):
    """
    Set a paragraph's text for plain-text editing, keeping its formatting and
    any non-text content.

    The first text run keeps its formatting and takes the new text. Other
    pure-text runs and hyperlinks (their text is part of what the user just
    rewrote) are removed. Runs holding anything else - an inline image, a
    field, a footnote reference - are left exactly where they are, so editing
    a caption next to an image no longer deletes the image.
    """
    p = paragraph._p
    text_runs = [r for r in p.findall(qn("w:r")) if is_text_run(r)]
    keeper = next((r for r in text_runs if r.find(qn("w:t")) is not None),
                  text_runs[0] if text_runs else None)
    for r in text_runs:
        if r is not keeper:
            p.remove(r)
    for h in p.findall(qn("w:hyperlink")):
        p.remove(h)
    if keeper is None:
        paragraph.add_run(new_text)
        return
    from docx.text.run import Run
    Run(keeper, paragraph).text = new_text


# ---------------------------------------------------------------------------
# small XML helpers
# ---------------------------------------------------------------------------

_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _xpath(el, xpath):
    """xpath with the w: prefix bound, for oxml and plain lxml elements alike
    (python-docx's own xpath() refuses a namespaces= argument)."""
    return etree._Element.xpath(el, xpath, namespaces=_NS)


def _first(el, xpath):
    found = _xpath(el, xpath)
    return found[0] if found else None


def _max_attr(el, xpath, attr):
    vals = []
    for node in _xpath(el, xpath):
        try:
            vals.append(int(node.get(qn(attr))))
        except (TypeError, ValueError):
            pass
    return max(vals) if vals else 0


def _insert_before_cleanup(numbering, el):
    cleanup = _first(numbering, "./w:numIdMacAtCleanup")
    if cleanup is not None:
        cleanup.addprevious(el)
    else:
        numbering.append(el)
