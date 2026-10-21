"""docx-master OOXML extension channel — JSON-driven post-processing.

`apply` (the Node engine) owns structure and content. This CLI owns the
features it deliberately doesn't write: section breaks, field-driven lists
(TOC / figure list / index), note parts, watermarks, page decoration,
drawings, charts, embedded objects, custom XML.

Usage:
    python extend.py config.json
    python extend.py config.json --dry-run

Config shape:
    {
      "source": "in.docx",
      "output": "out.docx",
      "operations": [
        {"op": "bookmark", "at": {"afterText": "第二章"}, "name": "chap2"},
        {"op": "bookmark", "name": "chap2-part2", "from": {"afterText": "2.1"},
         "to": {"afterText": "2.4"}},
        {"op": "sectionBreak", "at": {"afterText": "第二章"}, "kind": "nextPage",
         "columns": [2, 425, true]},
        {"op": "watermark", "header": 1, "text": "内部资料", "color": "C8C8C8"},
        {"op": "toc", "at": {"afterText": "目录"}, "levels": "1-3"},
        {"op": "footnote", "at": {"afterText": "脚注在此"}, "text": "脚注正文"},
        {"op": "smartArt", "at": {"afterText": "⟦图形落点⟧"}, "kind": "process",
         "nodes": ["需求", "设计", "实现", "验收"], "width": 14},
        {"op": "drawing", "at": {"afterText": "⟦图形落点⟧"}, "textBox": true,
         "paragraphs": [{"text": "标题", "size": 22, "bold": true},
                        {"text": "正文", "size": 18}],
         "width": 10, "height": 2.4},
        {"op": "canvas", "at": {"afterText": "⟦图形落点⟧"}, "width": 11, "height": 3,
         "children": [{"x": 0.2, "y": 0.2, "text": "A"},
                      {"x": 4.0, "y": 1.5, "text": "B"}]}
      ]
    }

Every operation is idempotent-ish in the sense that it appends rather than
search-and-replaces — run it once per document. Anchors accept
`afterText` / `beforeText` (substring match on paragraph text),
`paragraphIndex` (0-based body index) or `bookmark`.

Always follow with the engine's validator and the PRO `validate.py`:
    node scripts/validate.js out.docx
    python pro/scripts/validate.py out.docx
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import re
import sys

from lxml import etree

from _common import (IGNORABLE_PREFIXES, NS, PPR_ORDER, RPR_ORDER, SECTPR_ORDER,
                     DocxPackage, W, ensure_root_declarations, esc, insert_ordered,
                     local, para_text)
import capabilities as cap

DOCUMENT = "word/document.xml"
DOC_RELS = "word/_rels/document.xml.rels"
CT = "[Content_Types].xml"


# --------------------------------------------------------------------------
# Anchors
# --------------------------------------------------------------------------

class AnchorError(RuntimeError):
    pass


# Text anchors match on "paragraph text contains the substring" and take
# the first hit. That is fine for a unique sentinel, but a substring that
# appears in a heading *and* in body text (or in a TOC entry) silently
# lands on the wrong paragraph. Warning instead of guessing keeps the
# failure visible; each distinct needle only warns once per run.
_ANCHOR_WARNED = set()


def warn(message: str) -> None:
    print("[extend] warning: %s" % message, file=sys.stderr)


def warn_ambiguous_anchor(key, needle, hits) -> None:
    """A text anchor matched more than one paragraph and we took #1."""
    marker = (key, needle, hits)
    if marker in _ANCHOR_WARNED:
        return
    _ANCHOR_WARNED.add(marker)
    warn("%s %r matched %d paragraphs — using the first. Prefer a "
         "`bookmark` anchor (a Block `anchor`, a `format` op's `anchor`, or "
         "an `extend bookmark` op): it names one exact target, while a text "
         "needle can also hit a heading or a generated TOC entry."
         % (key, needle, hits))


def spec_flag(spec, *names, default=False):
    """Read a boolean that has more than one accepted spelling.

    Two spellings of the same switch (`flipV` in the reference docs vs
    `flipVertical` in the implementation) used to mean the documented one
    was read by nobody and failed *silently* — the shape just came out
    unflipped. Accepting both fixes the silent half; raising on a
    disagreement fixes the other half, because two contradicting values
    are a config bug, not a preference.
    """
    present = [(n, spec[n]) for n in names if n in spec]
    if not present:
        return default
    if len({bool(v) for _, v in present}) > 1:
        raise RuntimeError(
            "conflicting values for %s: %s — they are aliases, keep one"
            % (" / ".join(n for n, _ in present),
               ", ".join("%s=%r" % (n, v) for n, v in present)))
    return bool(present[0][1])


def body_of(root):
    body = root.find(W("body"))
    if body is None:
        raise AnchorError("document.xml has no <w:body>")
    return body


def body_paragraphs(root):
    """Top-level body paragraphs (table cells excluded)."""
    return [p for p in body_of(root) if p.tag == W("p")]


def resolve_anchor(root, spec, nth: int = 0):
    """Find the paragraph a spec points at. Returns the element."""
    if spec is None:
        raise AnchorError("operation needs an `at` anchor")
    if "paragraphIndex" in spec:
        paras = body_paragraphs(root)
        idx = int(spec["paragraphIndex"])
        if idx < 0:
            idx += len(paras)
        if not (0 <= idx < len(paras)):
            raise AnchorError("paragraphIndex %s out of range (0..%d)"
                              % (spec["paragraphIndex"], len(paras) - 1))
        return paras[idx]
    if "bookmark" in spec:
        name = spec["bookmark"]
        for el in root.iter(W("bookmarkStart")):
            if el.get(W("name")) == name:
                p = el.getparent()
                while p is not None and p.tag != W("p"):
                    p = p.getparent()
                if p is not None:
                    return p
        raise AnchorError("bookmark %r not found" % name)
    for key in ("afterText", "beforeText", "text"):
        if key in spec:
            needle = spec[key]
            hits = [p for p in body_paragraphs(root) if needle in para_text(p)]
            if len(hits) <= nth:
                raise AnchorError("%s %r matched %d paragraph(s), need #%d"
                                  % (key, needle, len(hits), nth + 1))
            if len(hits) > 1 and nth == 0:
                warn_ambiguous_anchor(key, needle, len(hits))
            return hits[nth]
    raise AnchorError("anchor spec needs one of paragraphIndex / bookmark / afterText")


def parse(xml_str: str):
    """Parse a full part. lxml rejects a `str` that carries an encoding
    declaration, so everything goes through bytes."""
    return etree.fromstring(xml_str.encode("utf-8"))


def parse_fragment(xml: str):
    """Parse a fragment that may contain several top-level elements."""
    wrapped = '<root xmlns:w="%s" xmlns:r="%s" xmlns:a="%s" xmlns:wp="%s" ' \
              'xmlns:pic="%s" xmlns:c="%s" xmlns:v="%s" xmlns:o="%s" xmlns:wps="%s" ' \
              'xmlns:wpg="%s" xmlns:wpc="%s" xmlns:w10="%s" xmlns:mc="%s" ' \
              'xmlns:w14="%s" xmlns:w15="%s" xmlns:wp14="%s" xmlns:b="%s">%s</root>' % (
                  NS["w"], NS["r"], NS["a"], NS["wp"], NS["pic"], NS["c"], NS["v"],
                  NS["o"], NS["wps"], NS["wpg"], NS["wpc"], NS["w10"], NS["mc"],
                  NS["w14"], NS["w15"], NS["wp14"], cap.BIB_NS, xml)
    return list(etree.fromstring(wrapped))


def insert_fragment(anchor, xml: str, where: str = "after"):
    nodes = parse_fragment(xml)
    parent = anchor.getparent()
    idx = parent.index(anchor)
    if where == "before":
        for offset, node in enumerate(nodes):
            parent.insert(idx + offset, node)
    else:
        for offset, node in enumerate(nodes):
            parent.insert(idx + 1 + offset, node)
    return nodes


def find_sectpr(root, include_breaks: bool = False):
    """SectPr list: paragraph-embedded ones (in order) plus the body-trailing
    one. Matches how Word sections map to document order."""
    found = []
    if include_breaks:
        for p in body_paragraphs(root):
            ppr = p.find(W("pPr"))
            if ppr is not None:
                sp = ppr.find(W("sectPr"))
                if sp is not None:
                    found.append(sp)
    trailing = root.find("%s/%s" % (W("body"), W("sectPr")))
    if trailing is not None:
        found.append(trailing)
    return found


def upsert_sectpr_child(sectpr, tag: str, xml_fragment: str):
    """Replace an existing child of the same name, else insert in order."""
    for child in list(sectpr):
        if local(child.tag) == tag:
            sectpr.remove(child)
    el = parse_fragment(xml_fragment)[0]
    insert_ordered(sectpr, el, SECTPR_ORDER)
    return el


# --------------------------------------------------------------------------
# Operations
# --------------------------------------------------------------------------

# Word's bookmark-name rules: ≤40 chars, no spaces, must not start with a digit.
# CJK is allowed (Word accepts it; the limit is on characters, not bytes).
BOOKMARK_NAME_RE = re.compile(r"^[A-Za-z_\u4e00-\u9fff][A-Za-z0-9_\u4e00-\u9fff]{0,39}$")


def next_bookmark_id(root) -> int:
    """Lowest `w:id` free across both bookmarkStarts and bookmarkEnds.

    Word requires the id to be unique package-wide (a duplicate makes Word
    renumber on open, and can pair a start with the wrong end). Scanning ends
    too matters: a stray end holding an id would otherwise be reused.
    """
    used = set()
    for tag in ("bookmarkStart", "bookmarkEnd"):
        for el in root.iter(W(tag)):
            raw = el.get(W("id"))
            if raw is None:
                continue
            try:
                used.add(int(raw))
            except ValueError:
                continue
    i = 0
    while i in used:
        i += 1
    return i


def op_bookmark(pkg, root, spec):
    """Wrap a paragraph — or a paragraph range — in a named bookmark.

    Word's own bookmarks (`_Toc*`, `_Ref*`, `_GoBack`) are produced by the TOC /
    REF machinery and rewritten on every field update. This op is for *authored*
    anchors: the name a REF field cites, a `HYPERLINK \\l` jump lands on, or a
    `TOC \\b` scope encloses. Nothing in a document points at a bookmark until
    something is written that names it, so create the bookmark first, then add
    the REF / link / field-scope that references it.
    """
    name = str(spec.get("name") or "").strip()
    if not name:
        raise AnchorError("bookmark needs a `name`")
    if not BOOKMARK_NAME_RE.match(name):
        raise AnchorError(
            "bookmark name %r is not a legal Word bookmark name — letters / CJK / "
            "digits / underscore, must not start with a digit, no spaces, max 40 chars" % name)
    for el in root.iter(W("bookmarkStart")):
        if el.get(W("name")) == name:
            raise AnchorError(
                "bookmark %r already exists — names must be unique, and a re-use would "
                "make every existing reference ambiguous" % name)

    spec_from = spec.get("from")
    if spec_from is not None:
        start_para = resolve_anchor(root, spec_from)
        end_para = resolve_anchor(root, spec["to"]) if spec.get("to") is not None else start_para
        paras = body_paragraphs(root)
        if paras.index(end_para) < paras.index(start_para):
            raise AnchorError("bookmark `from` anchor comes after its `to` anchor — swap them")
    else:
        start_para = end_para = resolve_anchor(root, spec.get("at") or spec, nth=spec.get("nth", 0))

    bid = next_bookmark_id(root)
    start_el = parse_fragment('<w:bookmarkStart w:id="%d" w:name="%s"/>' % (bid, esc(name)))[0]
    end_el = parse_fragment('<w:bookmarkEnd w:id="%d"/>' % bid)[0]

    point = spec.get("mark") == "point"
    pPr = start_para.find(W("pPr"))
    # bookmarkStart belongs with the paragraph content, after pPr — a start
    # placed before pPr is still schema-valid but Word moves it on save.
    if pPr is not None:
        pPr.addnext(start_el)
    else:
        start_para.insert(0, start_el)
    if point:
        start_el.addnext(end_el)
    else:
        end_para.append(end_el)

    where = "point" if point else (
        "paragraph" if start_para is end_para else "paragraphs")
    return "bookmark %r added (%s, w:id=%d)" % (name, where, bid)


def op_section_break(pkg, root, spec):
    """Insert a paragraph carrying a sectPr, which ends the current section
    and starts a new one with the declared properties."""
    anchor = resolve_anchor(root, spec.get("at"))
    trailing = root.find("%s/%s" % (W("body"), W("sectPr")))
    base = trailing if trailing is not None else root.find(".//%s" % W("sectPr"))
    sect_xml = cap.section_properties(
        kind=spec.get("kind", "nextPage"),
        pg_sz=spec.get("pgSz"),
        pg_mar=spec.get("pgMar"),
        pg_num=tuple(spec["pgNum"]) if spec.get("pgNum") else None,
        columns=tuple(spec["columns"]) if spec.get("columns") else None,
        direction=spec.get("textDirection"),
    )
    if base is not None:
        # Page geometry is inherited key by key: declaring only `pgSz` (e.g. a
        # landscape section) must not drop `pgMar` — a sectPr missing pgMar
        # falls back to Word's defaults and the text reflows unexpectedly.
        # Both are emitted in CT_SectPr order (pgSz before pgMar) right after
        # <w:type/>, so declared + inherited values are re-serialised together
        # rather than concatenated in declaration order.
        geometry = ""
        for tag in ("pgSz", "pgMar"):
            m = re.search(r"<w:%s[^>]*/>" % tag, sect_xml)
            if m:
                geometry += m.group(0)
                sect_xml = sect_xml[:m.start()] + sect_xml[m.end():]
                continue
            existing = base.find(W(tag))
            if existing is not None:
                geometry += etree.tostring(existing, encoding="unicode")
        if geometry:
            m = re.search(r"<w:type[^>]*/>", sect_xml)
            if m:
                sect_xml = sect_xml[:m.end()] + geometry + sect_xml[m.end():]
            else:
                sect_xml = sect_xml.replace("<w:sectPr>", "<w:sectPr>" + geometry, 1)
        for tag in ("docGrid",):
            existing = base.find(W(tag))
            if existing is not None and "<w:%s" % tag not in sect_xml:
                sect_xml = sect_xml.replace(
                    "</w:sectPr>", etree.tostring(existing, encoding="unicode") + "</w:sectPr>")
            # `cols` is inherited too, but CT_SectPr orders it after pgNumType —
            # so it can't ride along with pgSz / pgMar above. A bare break in a
            # multi-column document would otherwise drop back to one column.
            existing_cols = base.find(W("cols"))
            if existing_cols is not None and "<w:cols" not in sect_xml:
                cols_xml = etree.tostring(existing_cols, encoding="unicode")
                anchor_m = None
                for tag in ("pgNumType", "pgMar", "pgSz", "type"):
                    m3 = re.search(r"<w:%s[^>]*/>" % tag, sect_xml)
                    if m3:
                        anchor_m = m3
                        break
                if anchor_m is not None:
                    sect_xml = sect_xml[:anchor_m.end()] + cols_xml + sect_xml[anchor_m.end():]
                else:
                    sect_xml = sect_xml.replace("<w:sectPr>", "<w:sectPr>" + cols_xml, 1)
        # Header / footer references are positional, not inherited by a bare
        # sectPr: Word's "link to previous" is implemented by copying the
        # references down. Without this the sections created by this op render
        # headerless while the trailing body sectPr keeps its headers — the
        # single most visible way a hand-written break differs from one Word
        # inserts. CT_SectPr puts EG_HdrFtrReferences first, so they slot in
        # right after the opening tag, ahead of <w:type/>.
        if spec.get("inheritHeaders", True):
            refs = ""
            for tag in ("headerReference", "footerReference"):
                for el in base.findall(W(tag)):
                    refs += etree.tostring(el, encoding="unicode")
            if refs:
                sect_xml = sect_xml.replace("<w:sectPr>", "<w:sectPr>" + refs, 1)
            # `titlePg` rides along with the references — it is what makes the
            # first-page reference take effect. CT_SectPr orders it after
            # cols / vAlign and before textDirection / docGrid, so slot it
            # ahead of whichever of those is present.
            if base.find(W("titlePg")) is not None and "<w:titlePg" not in sect_xml:
                pos = None
                for tag in ("textDirection", "docGrid"):
                    m2 = re.search(r"<w:%s[^>]*/>" % tag, sect_xml)
                    if m2:
                        pos = m2.start()
                        break
                if pos is not None:
                    sect_xml = sect_xml[:pos] + "<w:titlePg/>" + sect_xml[pos:]
                else:
                    sect_xml = sect_xml.replace("</w:sectPr>", "<w:titlePg/></w:sectPr>", 1)
    nodes = insert_fragment(anchor, "<w:p><w:pPr>%s</w:pPr></w:p>" % sect_xml)
    return "section break inserted after anchor (%d node)" % len(nodes)


def op_toc(pkg, root, spec):
    """TOC / list-of-figures / index field, inserted as its own paragraph."""
    anchor = resolve_anchor(root, spec.get("at"))
    style = spec.get("style", "TOCHeading")
    if spec.get("identifier"):
        runs = cap.toc_of_field(spec["identifier"], spec.get("label", "图"))
        heading = spec.get("title", "%s目录" % spec.get("label", "图"))
    elif spec.get("index"):
        runs = cap.index_field(spec.get("columns", 1))
        heading = spec.get("title", "索引")
    else:
        runs = cap.toc_field(levels=spec.get("levels", "1-3"),
                             hyperlinks=spec.get("hyperlinks", True))
        heading = spec.get("title", "目录")
    title_para = ('<w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>'
                  '<w:r><w:t>%s</w:t></w:r></w:p>' % (style, esc(heading)))
    field_para = "<w:p>%s</w:p>" % runs
    where = spec.get("where", "after")
    insert_fragment(anchor, title_para + field_para, where)
    return "field paragraph inserted"


def op_field(pkg, root, spec):
    anchor = resolve_anchor(root, spec.get("at"))
    runs = cap.field_runs(spec["instruction"], spec.get("placeholder", ""))
    if spec.get("inline"):
        # Inline fields must land INSIDE the anchor paragraph, otherwise the
        # run sequence becomes a sibling of <w:p> at body level, which is not
        # valid OOXML (body only allows block-level content plus a trailing
        # sectPr). Climb to the nearest ancestor paragraph first, so anchors
        # pointing at a bookmark/run still work.
        para = anchor
        while para is not None and local(para.tag) != "p":
            para = para.getparent()
        if para is None:
            raise AnchorError("inline field anchor is not inside a paragraph")
        for node in parse_fragment("<w:r><w:t xml:space='preserve'> </w:t></w:r>" + runs):
            para.append(node)
        return "inline field appended into anchor paragraph"
    insert_fragment(anchor, "<w:p>%s</w:p>" % runs, spec.get("where", "after"))
    return "field paragraph inserted"


def op_note(pkg, root, spec):
    """Add a footnote / endnote: create the part if needed, register it,
    append the note body and drop the reference mark into the anchor."""
    # `op: "endnote"` must not silently land in footnotes.xml — the op name
    # is the default, an explicit `kind` overrides it.
    kind = spec.get("kind") or ("endnote" if spec.get("_op") == "endnote" else "footnote")
    is_footnote = kind == "footnote"
    part = "word/footnotes.xml" if is_footnote else "word/endnotes.xml"
    root_tag = "w:footnotes" if is_footnote else "w:endnotes"
    content_type = cap.FOOTNOTES_CT if is_footnote else cap.ENDNOTES_CT
    rel_type = cap.FOOTNOTES_REL if is_footnote else cap.ENDNOTES_REL

    xml = pkg.read(part)
    if xml is None:
        xml = cap.note_part(root_tag, separator_text="separator",
                            continuation_text="continuation")
        pkg.write(part, xml)
        pkg.add_content_type_override(part, content_type)
        pkg.add_relationship(DOC_RELS, rel_type, part.split("/", 1)[1])
        tree = parse(ensure_root_declarations(xml))
    else:
        tree = parse(ensure_root_declarations(xml))

    ids = [int(e.get(W("id"))) for e in tree if e.get(W("id")) is not None]
    note_id = max([i for i in ids if i > 0] or [0]) + 1
    entry = parse_fragment(cap.note_entry(kind, note_id, spec.get("text", "")))[0]
    tree.append(entry)
    pkg.write(part, etree.tostring(tree, encoding="unicode", xml_declaration=False))

    anchor = resolve_anchor(root, spec.get("at"))
    run = parse_fragment(cap.note_reference_run(kind, note_id))[0]
    anchor.append(run)
    return "%s #%d added (%s)" % (kind, note_id, part)


def op_watermark(pkg, root, spec):
    """Insert a watermark paragraph into a header part."""
    index = spec.get("header", 1)
    part = "word/header%d.xml" % index
    body = cap.watermark_paragraph(
        text=spec.get("text", "内部资料"),
        color=spec.get("color", "C8C8C8"),
        font=spec.get("font", "宋体"),
        rotation=int(spec.get("rotation", 315)),
        opacity=str(spec.get("opacity", ".5")),
    )
    xml = pkg.read(part)
    if xml is None:
        raise AnchorError("%s not found — declare a header in apply's "
                          "headerFooter config first" % part)
    tree = parse(ensure_root_declarations(xml))
    hdr_root = tree
    for node in parse_fragment(body):
        hdr_root.append(node)
    pkg.write(part, etree.tostring(tree, encoding="unicode", xml_declaration=False))
    return "watermark written to %s" % part


def op_page_property(pkg, root, spec):
    """Mutate section properties across a section selector: borders,
    background colour, page numbering, columns, text direction."""
    selectors = spec.get("sections", "all")
    sectprs = find_sectpr(root, include_breaks=True)
    if not sectprs:
        raise AnchorError("no sectPr found")
    if selectors == "all":
        targets = list(range(len(sectprs)))
    elif isinstance(selectors, list):
        targets = [int(s) - 1 for s in selectors]
    else:
        m = re.match(r"^(\d+)-(\d+)$", str(selectors))
        targets = (list(range(int(m.group(1)) - 1, int(m.group(2)))) if m
                   else [int(selectors) - 1])
    touched = 0
    for idx in targets:
        if not (0 <= idx < len(sectprs)):
            continue
        sp = sectprs[idx]
        if spec.get("borders"):
            upsert_sectpr_child(sp, "pgBorders", cap.pg_borders(**spec["borders"]))
        if spec.get("pgNum"):
            fmt, start = spec["pgNum"]
            upsert_sectpr_child(sp, "pgNumType", cap.pg_num_type(fmt, int(start)))
        if spec.get("columns"):
            num, space, sep = spec["columns"]
            upsert_sectpr_child(sp, "cols", cap.cols(int(num), int(space), bool(sep)))
        if spec.get("textDirection"):
            upsert_sectpr_child(sp, "textDirection", cap.text_direction(spec["textDirection"]))
        touched += 1
    return "section properties applied to %d section(s)" % touched


def op_page_background(pkg, root, spec):
    """`<w:background>` is the first child of `<w:document>`."""
    for existing in root.findall(W("background")):
        root.remove(existing)
    if not spec.get("color"):
        return "background removed"
    frag = cap.background_shape(spec["color"]) if spec.get("drawing") else \
        '<w:background w:color="%s"/>' % spec["color"]
    node = parse_fragment(frag)[0]
    root.insert(0, node)
    return "page background set to %s" % spec["color"]


def op_drawing(pkg, root, spec):
    """Shapes / text boxes / WordArt / canvases / connectors / ink / symbols."""
    anchor = resolve_anchor(root, spec.get("at"))
    which = spec.get("shape", "roundRect")
    if spec.get("symbols"):
        frag = cap.symbol_paragraph(spec["symbols"])
    elif spec.get("ink"):
        frag = cap.ink_run(spec["ink"], **{k: v for k, v in spec.items()
                                           if k in ("color", "width_pt", "cx_cm", "cy_cm")})
    elif spec.get("textBox"):
        # A real text box: `txBox="1"`, no default fill/outline, multi-paragraph.
        paragraphs = cap.rich_paragraphs(spec.get("paragraphs") or spec.get("text", ""))
        inner = cap.text_box(
            paragraphs,
            cx_cm=float(spec.get("width", 6.0)),
            cy_cm=float(spec.get("height", 2.0)),
            fill=spec.get("fill"), line=spec.get("line"),
            anchor=spec.get("anchor", "t"),
            wrap=spec.get("wrapText", "square"),
            rot_deg=float(spec.get("rotation", 0)),
            body_insets=(tuple(spec["insets"]) if spec.get("insets")
                         else (91440, 45720, 91440, 45720)),
        )
        frag = cap.wrap_inline(inner, float(spec.get("width", 6.0)),
                               float(spec.get("height", 2.0)),
                               spec.get("name", "文本框"))
    elif spec.get("connector"):
        inner = cap.connector(
            cx_cm=float(spec.get("width", 3.0)),
            cy_cm=float(spec.get("height", 1.5)),
            flip_vertical=spec_flag(spec, "flipVertical", "flipV"),
            flip_horizontal=spec_flag(spec, "flipHorizontal", "flipH"),
            arrow=spec_flag(spec, "arrow", default=True),
            color=spec.get("line", "4A6FA5"),
            width=int(spec.get("lineWidthEmu", 19050)),
        )
        frag = cap.wrap_inline(inner, float(spec.get("width", 3.0)),
                               float(spec.get("height", 1.5)),
                               spec.get("name", "连接线"))
    else:
        paragraphs = (cap.rich_paragraphs(spec["paragraphs"])
                      if spec.get("paragraphs") else None)
        inner = cap.shape(
            preset=which,
            cx_cm=float(spec.get("width", 3.4)),
            cy_cm=float(spec.get("height", 1.6)),
            text=spec.get("text", ""),
            fill=spec.get("fill", "E8F0FE"),
            line=spec.get("line", "4A6FA5"),
            text_warp=spec.get("warp"),
            paragraphs=paragraphs,
            anchor=spec.get("anchor", "ctr"),
            body_insets=tuple(spec["insets"]) if spec.get("insets") else None,
        )
        frag = cap.wrap_inline(inner, float(spec.get("width", 3.4)),
                               float(spec.get("height", 1.6)),
                               spec.get("name", "形状"))
    insert_fragment(anchor, frag, spec.get("where", "after"))
    return "drawing inserted"


def op_canvas(pkg, root, spec):
    """Canvas of positioned children. Each child takes `x` / `y` (cm,
    relative to the canvas top-left), `shape`, `width`, `height`, `text`."""
    anchor = resolve_anchor(root, spec.get("at"))
    inner = []
    for item in spec.get("children", []):
        inner.append(cap.shape(
            preset=item.get("shape", "roundRect"),
            cx_cm=float(item.get("width", 2.4)),
            cy_cm=float(item.get("height", 1.2)),
            text=item.get("text", ""),
            fill=item.get("fill", "E8F0FE"),
            line=item.get("line", "4A6FA5"),
            offset_cm=(float(item.get("x", 0.0)), float(item.get("y", 0.0))),
            paragraphs=(cap.rich_paragraphs(
                [{"text": item.get("text", ""),
                  "size": int(item.get("size", 18)),
                  "color": item.get("textColor", "1F3A5F"),
                  "align": item.get("align", "center")}])
                if item.get("text") else None),
        ))
    frag = cap.wrap_inline(cap.canvas(inner, float(spec.get("width", 8.0)),
                                      float(spec.get("height", 3.0))),
                           float(spec.get("width", 8.0)), float(spec.get("height", 3.0)),
                           spec.get("name", "绘图画布"))
    insert_fragment(anchor, frag, spec.get("where", "after"))
    return "canvas inserted with %d shape(s)" % len(inner)


def op_smart_art(pkg, root, spec):
    """SmartArt-equivalent diagram (auto-laid-out shapes + connectors).

    Spec: {at, kind: process|list|cycle|hierarchy, nodes: [str | {text}],
    width?, height?, fill?, line?, textColor?, font?, size?, arrow?}
    """
    anchor = resolve_anchor(root, spec.get("at"))
    nodes = spec.get("nodes") or []
    if not nodes:
        raise RuntimeError("smartArt needs `nodes` — a non-empty list of "
                           "labels (strings or {text})")
    inner, w_cm, h_cm = cap.smart_art(
        spec.get("kind", "process"), nodes,
        width_cm=float(spec.get("width", 12.0)),
        height_cm=float(spec["height"]) if spec.get("height") else None,
        fill=spec.get("fill", "DCE9F9"), line=spec.get("line", "4A6FA5"),
        text_color=spec.get("textColor", "1F3A5F"),
        font=spec.get("font", "微软雅黑"),
        size=int(spec.get("size", 18)),
        arrow=bool(spec.get("arrow", True)),
    )
    frag = cap.wrap_inline(inner, w_cm, h_cm, spec.get("name", "SmartArt 图示"))
    insert_fragment(anchor, frag, spec.get("where", "after"))
    return "smartArt (%s) inserted: %d node(s), %.1f×%.1f cm" % (
        spec.get("kind", "process"), len(nodes), w_cm, h_cm)


def op_image(pkg, root, spec):
    """Inline or floating picture. Requires the image to already exist as a
    part in the package (apply's `image` block does that)."""
    anchor = resolve_anchor(root, spec.get("at"))
    part = spec.get("part")
    local_file = spec.get("file") or spec.get("src")
    if local_file:
        # Copy a local image in rather than requiring apply to have imported
        # it — extend-only documents (and extra figures) had no way to add
        # media at all before this.
        if not os.path.isfile(local_file):
            raise AnchorError("image file not found: %s" % local_file)
        ext = (os.path.splitext(local_file)[1].lstrip(".") or "png").lower()
        part = "word/media/%s.%s" % (os.path.splitext(os.path.basename(local_file))[0], ext)
        with open(local_file, "rb") as fh:
            pkg.write(part, fh.read())
        pkg.add_content_type_default(
            ext, "image/jpeg" if ext in ("jpg", "jpeg") else "image/%s" % ext)
    elif part:
        if pkg.read_bytes(part) is None:
            raise AnchorError("image part %r is not in the package — pass `file` "
                              "(a local image path) to copy one in" % part)
    else:
        raise AnchorError("image needs `part` (a media part already in the "
                          "package) or `file` (a local image to copy in)")
    rid = spec.get("rid")
    if not rid:
        rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        rid = pkg.add_relationship(DOC_RELS, rel_type, part.replace("word/", "", 1))
    frag = cap.picture_run(
        rid, float(spec.get("width", 8.0)), float(spec.get("height", 5.0)),
        spec.get("name", "图片"),
        crop=tuple(spec["crop"]) if spec.get("crop") else None,
        lum=tuple(spec["lum"]) if spec.get("lum") else None,
        effects=spec.get("effects", ""),
        floating=spec_flag(spec, "floating"),
        wrap_text=spec.get("wrapText", "bothSides"),
        offset_x_cm=float(spec.get("offsetX", 0.0)),
        offset_y_cm=float(spec.get("offsetY", 0.0)),
        allow_overlap=spec_flag(spec, "allowOverlap", default=True),
        behind_doc=spec_flag(spec, "behindDoc", "behindDocument"),
    )
    insert_fragment(anchor, frag, spec.get("where", "after"))
    if spec.get("floating"):
        return ("picture inserted (floating at %.1f,%.1f cm from the anchor "
                "line%s)"
                % (float(spec.get("offsetX", 0.0)),
                   float(spec.get("offsetY", 0.0)),
                   "" if spec_flag(spec, "allowOverlap", default=True)
                   else ", overlap forbidden"))
    return "picture inserted (inline)"


def op_chart(pkg, root, spec):
    """Chart part + workbook data source + the drawing that references it."""
    anchor = resolve_anchor(root, spec.get("at"))
    chart_name = spec.get("name", "chart1")
    chart_part = "word/charts/%s.xml" % chart_name
    chart_rels_part = "word/charts/_rels/%s.xml.rels" % chart_name
    book = "word/embeddings/%s-data.xlsx" % chart_name

    pkg.write(book, cap.minimal_xlsx(spec["rows"]))
    pkg.add_content_type_override(book, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    data_rid = pkg.add_relationship(chart_rels_part if pkg.has(chart_rels_part)
                                    else _seed_chart_rels(pkg, chart_rels_part),
                                    cap.CHART_COLOR_REL,
                                    "/" + book)
    pkg.write(chart_part, cap.chart_part(
        spec["categories"], spec["series"], title=spec.get("title", ""),
        relation_id=data_rid, kind=spec.get("kind", "bar")))
    pkg.add_content_type_override(chart_part, cap.CHART_CT)
    chart_rid = pkg.add_relationship(DOC_RELS, cap.CHART_REL,
                                     chart_part.replace("word/", "", 1))
    insert_fragment(anchor, cap.chart_reference_run(chart_rid,
                                                    float(spec.get("width", 11.0)),
                                                    float(spec.get("height", 5.5)),
                                                    spec.get("title", "图表")),
                    spec.get("where", "after"))
    return "chart inserted (%s + %s)" % (chart_part, book)


def _seed_chart_rels(pkg, path):
    if not pkg.has(path):
        # Full open+close tags — `add_relationship` patches the XML via a
        # string replace on "</Relationships>", which never matches a
        # self-closed element.
        pkg.write(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                         '<Relationships xmlns="%s"></Relationships>' % NS["rel"])
    return path


def op_ole(pkg, root, spec):
    """Embed a workbook as an OLE object (VML preview + o:OLEObject)."""
    anchor = resolve_anchor(root, spec.get("at"))
    payload = cap.minimal_xlsx(spec["rows"])
    cfb = cap.ole_compound_file(spec.get("stream", "Package"), payload)
    bin_part = "word/embeddings/oleObject%d.bin" % pkg.next_id()
    pkg.write(bin_part, cfb)
    pkg.add_content_type_default("bin", cap.OLE_CT)
    ole_rid = pkg.add_relationship(DOC_RELS, cap.OLE_REL, bin_part.replace("word/", "", 1))

    image_part = spec.get("previewPart")
    preview_file = spec.get("previewFile")
    if preview_file:
        # Copy a local image in as the preview. Without this the caller has to
        # arrange for the media part to already exist, which is impossible from
        # a pure extend config.
        if not os.path.isfile(preview_file):
            raise AnchorError("ole previewFile not found: %s" % preview_file)
        ext = (os.path.splitext(preview_file)[1].lstrip(".") or "png").lower()
        image_part = "word/media/ole_preview.%s" % ext
        with open(preview_file, "rb") as fh:
            pkg.write(image_part, fh.read())
        pkg.add_content_type_default(
            ext, "image/jpeg" if ext in ("jpg", "jpeg") else "image/%s" % ext)
    elif image_part:
        # A dangling relationship makes Word report a repair prompt, so fail
        # loudly here instead of emitting a rel to a part that isn't there.
        if pkg.read(image_part) is None:
            raise AnchorError(
                "ole previewPart %r is not in the package — pass `previewFile` "
                "(a local image to copy in) or point `previewPart` at an "
                "existing media part" % image_part)
    else:
        raise AnchorError("ole needs `previewPart` (a media part already in the "
                          "package) or `previewFile` (a local image to copy in)")
    image_rid = pkg.add_relationship(
        DOC_RELS,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        image_part.replace("word/", "", 1))
    insert_fragment(anchor, cap.ole_object_paragraph(
        ole_rid, image_rid, program_id=spec.get("progId", "Package"),
        name=spec.get("name", "嵌入的对象")), spec.get("where", "after"))
    return "OLE object embedded (%s, %d bytes CFB)" % (bin_part, len(cfb))


def op_xe(pkg, root, spec):
    """Append a hidden `XE` index-entry field to the anchor paragraph."""
    anchor = resolve_anchor(root, spec.get("at"))
    run = parse_fragment(cap.xe_runs(spec["term"]))
    for r in run:
        anchor.append(r)
    return "XE entry %r added" % spec["term"]


def op_sdt(pkg, root, spec):
    """Insert a block-level rich-text content control (SDT)."""
    anchor = resolve_anchor(root, spec.get("at"))
    frag = cap.sdt_paragraph(spec.get("text", ""),
                             alias=spec.get("alias", ""),
                             tag=spec.get("tag", ""),
                             style=spec.get("style", "BodyText"),
                             dropdown=spec.get("dropdown"),
                             doc_part_gallery=spec.get("docPartGallery"))
    nodes = insert_fragment(anchor, frag, spec.get("where", "after"))
    kind = ("docPartObj(%s)" % spec["docPartGallery"]
            if spec.get("docPartGallery")
            else ("dropdown" if spec.get("dropdown") else "richText"))
    return "SDT inserted (%s; %d node)" % (kind, len(nodes))


def op_source(pkg, root, spec):
    """Register a bibliography source (word/sources.xml) and optionally
    drop an inline CITATION field at an anchor.

    Field spellings: `sourceType` (canonical) also accepts `type`, and
    `citeAt` also accepts `cite` — a `{at, text}` object as documented, or
    a bare anchor. `cite.text` becomes the field's placeholder text.
    """
    tag = spec["tag"]
    source_type = spec.get("sourceType")
    if source_type is None:
        source_type = spec.get("type")
    if source_type is None:
        raise RuntimeError(
            "source op needs `sourceType` (a b:SourceType such as Book / "
            "JournalArticle / Report / Misc). `type` is accepted as an "
            "alias.")
    if (spec.get("sourceType") is not None and spec.get("type") is not None
            and spec["sourceType"] != spec["type"]):
        raise RuntimeError(
            "source op: `sourceType` (%r) and `type` (%r) disagree — they "
            "are aliases, keep one."
            % (spec["sourceType"], spec["type"]))
    entry = cap.source_entry(tag, source_type, spec.get("fields", {}),
                             authors=spec.get("authors"))
    xml = pkg.read("word/sources.xml")
    if xml is None:
        pkg.write("word/sources.xml", cap.sources_part(entry))
        pkg.add_content_type_override("word/sources.xml", cap.SOURCES_CT)
        pkg.add_relationship(DOC_RELS, cap.SOURCES_REL, "sources.xml")
    else:
        tree = parse(ensure_root_declarations(xml))
        tree.append(parse_fragment(entry)[0])
        pkg.write("word/sources.xml",
                  etree.tostring(tree, encoding="unicode", xml_declaration=False))
    cited = ""
    cite_at = spec.get("citeAt")
    cite_text = None
    if cite_at is None and spec.get("cite") is not None:
        # Documented as `cite: {at, text}`; a bare anchor works too.
        cite = spec["cite"]
        if isinstance(cite, dict):
            cite_at = cite.get("at")
            cite_text = cite.get("text")
        else:
            cite_at = cite
    if cite_at:
        anchor = resolve_anchor(root, cite_at)
        locale = spec.get("localeId", 2052)
        runs = parse_fragment(cap.field_runs("CITATION %s \\l %d" % (tag, locale),
                                             cite_text or ("[%s]" % tag)))
        for r in runs:
            anchor.append(r)
        cited = " + CITATION field"
    return "source %r registered%s" % (tag, cited)


def op_drop_cap(pkg, root, spec):
    """Turn the anchor paragraph's first character into a drop cap
    (`w:framePr dropCap`) and enlarge that character's run."""
    anchor = resolve_anchor(root, spec.get("at"))
    ppr = anchor.find(W("pPr"))
    if ppr is None:
        ppr = etree.Element(W("pPr"))
        anchor.insert(0, ppr)
    if ppr.find(W("framePr")) is not None:
        raise RuntimeError("paragraph already has a framePr — drop cap already set?")
    fp = etree.Element(W("framePr"))
    fp.set(W("dropCap"), spec.get("kind", "drop"))  # drop | margin | none
    fp.set(W("lines"), str(int(spec.get("lines", 2))))
    fp.set(W("wrap"), "around")
    fp.set(W("vAnchor"), "text")
    fp.set(W("hAnchor"), "text")
    insert_ordered(ppr, fp, PPR_ORDER)

    half_pts = int(float(spec.get("size", "21pt").rstrip("pt")) * 2) \
        if isinstance(spec.get("size"), str) and spec.get("size", "").endswith("pt") \
        else int(spec.get("size", 42))
    runs = [r for r in anchor if r.tag == W("r")]
    if not runs:
        return "framePr set (no runs to enlarge)"
    first = runs[0]
    t = first.find(W("t"))
    if t is None or not t.text:
        return "framePr set (first run has no text)"
    ch, rest = t.text[0], t.text[1:]
    t.text = ch
    rpr = first.find(W("rPr"))
    if rpr is None:
        rpr = etree.Element(W("rPr"))
        first.insert(0, rpr)
    for tag in ("sz", "szCs"):
        el = etree.Element(W(tag))
        el.set(W("val"), str(half_pts))
        insert_ordered(rpr, el, RPR_ORDER)
    if rest:
        new_r = copy.deepcopy(first)
        rpr2 = new_r.find(W("rPr"))
        for tag in ("sz", "szCs"):
            el = rpr2.find(W(tag)) if rpr2 is not None else None
            if el is not None:
                rpr2.remove(el)
        t2 = new_r.find(W("t"))
        t2.text = rest
        first.addnext(new_r)
    return "drop cap %r x%d lines, size %d half-points" % (ch, int(spec.get("lines", 2)), half_pts)


def op_nested_table(pkg, root, spec):
    """Insert a nested table into the table cell whose text matches
    `at: {cellText: "...", nth?: N}`. Lands before the cell's last
    paragraph so the `w:tc`-must-end-with-`w:p` invariant holds."""
    needle = spec["at"]["cellText"]
    nth = int(spec["at"].get("nth", 0))
    hits = [tc for tc in root.iter(W("tc"))
            if needle in "".join(t.text or "" for t in tc.iter(W("t")))]
    if len(hits) <= nth:
        raise AnchorError("cellText %r matched %d cell(s), need #%d"
                          % (needle, len(hits), nth + 1))
    tc = hits[nth]
    paras = [c for c in tc if c.tag == W("p")]
    if not paras:
        raise AnchorError("target cell has no paragraph")
    frag = cap.nested_table(spec["rows"], widths=spec.get("widths"),
                            border_color=spec.get("borderColor", "808080"),
                            header_fill=spec.get("headerFill", ""),
                            header_bold=spec.get("headerBold", False))
    paras[-1].addprevious(parse_fragment(frag)[0])
    return "nested table (%dx%d) inserted into cell %r" % (
        len(spec["rows"]), len(spec["rows"][0]), needle)


def op_custom_props(pkg, root, spec):
    """Write `docProps/custom.xml` (custom document properties) and register
    it in the package root rels + content types."""
    if not spec.get("properties"):
        raise RuntimeError("customProps needs a `properties` object")
    pkg.write("docProps/custom.xml", cap.custom_props_xml(spec["properties"]))
    pkg.add_content_type_override("docProps/custom.xml", cap.CUSTOM_PROPS_CT)
    rid = pkg.add_relationship("_rels/.rels", cap.CUSTOM_PROPS_REL,
                               "docProps/custom.xml")
    return "docProps/custom.xml written (%d properties, relationship %s)" % (
        len(spec["properties"]), rid)


def op_custom_xml(pkg, root, spec):
    index = 1
    while pkg.has("customXml/item%d.xml" % index):
        index += 1
    item = "customXml/item%d.xml" % index
    props = "customXml/itemProps%d.xml" % index
    pkg.write(item, cap.custom_xml_item(spec.get("fields", {}),
                                        root_name=spec.get("root", "project")))
    pkg.write(props, cap.custom_xml_item_props())
    pkg.write("customXml/item%d.xml.rels" % index,
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="%s">'
              '<Relationship Id="rId1" Type="%s" Target="itemProps%d.xml"/>'
              '</Relationships>' % (NS["rel"], cap.CUSTOM_XML_PROPS_REL, index))
    pkg.add_content_type_override(props, cap.CUSTOM_XML_CT)
    rid = pkg.add_relationship(DOC_RELS, cap.CUSTOM_XML_REL, "../" + item)
    return "custom XML part %s registered (relationship %s)" % (item, rid)


def op_glossary(pkg, root, spec):
    """Write `word/glossary/document.xml` — Building Blocks / 自动图文集.

    Each entry: {name, paragraphs|xml, gallery?, category?, autoInsert?}.
    `paragraphs` items are plain text or objects: {text, style?, align?,
    bold?, italic?, size?, color?, font?, before?, after?, pageBreakBefore?}.
    Alignment/size/color make cover-page (`coverPg`) templates look like
    real covers instead of a stack of plain paragraphs. Entries show up
    in Word's Building Blocks Organizer (插入 → 文档部件 / 封面) under
    the declared gallery (autoTxt = 自动图文集, coverPg = 封面页).
    Repeat calls append.
    """
    entries = spec.get("entries")
    if not entries:
        raise RuntimeError("glossary needs `entries` — a list of "
                           "{name, paragraphs|xml, gallery?, category?}")
    parts = []
    for e in entries:
        blocks = ""
        for para in e.get("paragraphs", []):
            if isinstance(para, dict):
                text = str(para.get("text", ""))
                ppr = ""
                if para.get("style"):
                    ppr += '<w:pStyle w:val="%s"/>' % esc(para["style"])
                if para.get("pageBreakBefore"):
                    ppr += "<w:pageBreakBefore/>"
                spacing = ""

                def _twips(v):
                    m = re.match(r"^([0-9.]+)\s*(cm|mm|pt|in)?$", str(v))
                    if not m:
                        raise RuntimeError(
                            "glossary %s: %r is not a length (cm/mm/pt/in)"
                            % (e.get("name"), v))
                    val = float(m.group(1))
                    unit = m.group(2) or "pt"
                    tw = {"cm": 567, "mm": 56.7, "pt": 20, "in": 1440}[unit]
                    return str(int(round(val * tw)))

                if para.get("before") is not None or para.get("after") is not None:
                    if para.get("before") is not None:
                        spacing += ' w:before="%s"' % _twips(para["before"])
                    if para.get("after") is not None:
                        spacing += ' w:after="%s"' % _twips(para["after"])
                    ppr += "<w:spacing%s/>" % spacing
                if para.get("align"):
                    jc = {"left": "left", "center": "center", "right": "right",
                          "both": "both", "justify": "both",
                          "distribute": "distribute"}.get(str(para["align"]).lower())
                    if not jc:
                        raise RuntimeError(
                            "glossary align: %r is not left/center/right/"
                            "both/justify/distribute" % para["align"])
                    ppr += '<w:jc w:val="%s"/>' % jc
                rpr = ""
                if para.get("font"):
                    rpr += ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
                            % (esc(para["font"]), esc(para["font"]), esc(para["font"])))
                if para.get("bold"):
                    rpr += "<w:b/><w:bCs/>"
                if para.get("italic"):
                    rpr += "<w:i/><w:iCs/>"
                if para.get("color"):
                    rpr += '<w:color w:val="%s"/>' % esc(para["color"])
                if para.get("size") is not None:
                    # size is a font size in pt; half-points = pt * 2
                    m = re.match(r"^([0-9.]+)\s*(pt)?$", str(para["size"]))
                    if not m:
                        raise RuntimeError(
                            "glossary size: %r is not a pt font size "
                            "(e.g. 48 or \"48pt\")" % para["size"])
                    half = str(int(round(float(m.group(1)) * 2)))
                    rpr += ('<w:sz w:val="%s"/><w:szCs w:val="%s"/>' % (half, half))
                run_pr = "<w:rPr>%s</w:rPr>" % rpr if rpr else ""
                blocks += ('<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
                           '</w:p>' % (("<w:pPr>%s</w:pPr>" % ppr) if ppr else "",
                                       run_pr, esc(text)))
            else:
                blocks += ('<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r>'
                           '</w:p>' % esc(str(para)))
        blocks += e.get("xml", "")
        if not blocks:
            raise RuntimeError("glossary entry %r is empty — give it "
                               "`paragraphs` or `xml`" % e.get("name"))
        parts.append(cap.glossary_entry(
            e["name"], blocks, gallery=e.get("gallery", "autoTxt"),
            category=e.get("category", "General"),
            auto_insert=bool(e.get("autoInsert"))))
    part = "word/glossary/document.xml"
    xml = pkg.read(part)
    if xml is None:
        pkg.write(part, cap.glossary_part("".join(parts)))
        pkg.add_content_type_override(part, cap.GLOSSARY_CT)
        pkg.add_relationship(DOC_RELS, cap.GLOSSARY_REL, "glossary/document.xml")
    else:
        tree = parse(ensure_root_declarations(xml))
        doc_parts = tree.find(W("docParts"))
        if doc_parts is None:
            doc_parts = etree.SubElement(tree, W("docParts"))
        for node in parse_fragment("".join(parts)):
            doc_parts.append(node)
        pkg.write(part, etree.tostring(tree, encoding="unicode",
                                       xml_declaration=False))
    return "glossary: %d building block(s) written to %s" % (len(entries), part)


def op_cover(pkg, root, spec):
    """Prepend a cover page: title + subtitle lines + a section break that
    vertically centers the cover section (`vAlign=center`) and enables
    first-page-different (`titlePg`) so cover pages show no running
    header/footer unless first-page ones are declared.

    Spec: {title, subtitle?: [..], titleStyle?: "Title",
    subtitleStyle?: "Subtitle"}. Styles must exist in the document —
    apply's default style set includes Title / Subtitle.
    """
    title = spec.get("title")
    if not title:
        raise RuntimeError("cover needs `title`")
    tstyle = spec.get("titleStyle", "Title")
    sstyle = spec.get("subtitleStyle", "Subtitle")
    frag = ('<w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>'
            '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (esc(tstyle), esc(str(title))))
    for sub in spec.get("subtitle", []):
        frag += ('<w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>'
                 '<w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                 % (esc(sstyle), esc(str(sub))))
    sect = cap.section_properties(kind="nextPage", valign="center",
                                  title_pg=bool(spec.get("firstPageDifferent", True)))
    frag += "<w:p><w:pPr>%s</w:pPr></w:p>" % sect
    body = body_of(root)
    first = next((c for c in body if c.tag == W("p")), None)
    nodes = parse_fragment(frag)
    if first is not None:
        idx = body.index(first)
    else:
        idx = 0
    for offset, node in enumerate(nodes):
        body.insert(idx + offset, node)
    return "cover page prepended (%d paragraph(s), vAlign=center, titlePg)" % len(nodes)


OPERATIONS = {
    "bookmark": op_bookmark,
    "sectionBreak": op_section_break,
    "toc": op_toc,
    "field": op_field,
    "footnote": op_note,
    "endnote": op_note,
    "pageProperty": op_page_property,
    "pageBackground": op_page_background,
    "image": op_image,
    "xe": op_xe,
    "sdt": op_sdt,
    "source": op_source,
    "dropCap": op_drop_cap,
    "nestedTable": op_nested_table,
    "customProps": op_custom_props,
    "cover": op_cover,
}

# 标准版（standard）保留日常办公 / 论文 / 实训报告所需的全部基础能力。
# 以下 op 仅在 pro 版提供，见 `git checkout pro` 或 README「两个版本」：
#   watermark（水印）、drawing / canvas / smartArt（形状、文本框、画布、智能图形）、
#   chart（图表）、ole（OLE 嵌入）、glossary（构建基块）、customXml（自定义 XML）
UNSUPPORTED_OPS = {
    "watermark": "watermark（水印）",
    "drawing": "drawing（形状 / 文本框）",
    "canvas": "canvas（绘图画布）",
    "smartArt": "smartArt（智能图形）",
    "chart": "chart（图表）",
    "ole": "ole（OLE 嵌入）",
    "glossary": "glossary（构建基块）",
    "customXml": "customXml（自定义 XML）",
}


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------

def run(config_path: str, dry_run: bool = False) -> int:
    config = json.load(open(config_path, encoding="utf-8"))
    source = config["source"]
    output = config.get("output", source)
    ops = config.get("operations", [])

    pkg = DocxPackage(source)
    try:
        raw = ensure_root_declarations(pkg.read(DOCUMENT))
        root = parse(raw)

        results = []
        for i, spec in enumerate(ops):
            name = spec.get("op")
            handler = OPERATIONS.get(name)
            if handler is None:
                if name in UNSUPPORTED_OPS:
                    raise SystemExit(
                        "operations[%d]: %r 仅在 pro 版提供（标准版不含 %s）。"
                        "完整功能版：git checkout pro 或见 README「两个版本」；"
                        "标准版可用操作：%s"
                        % (i, name, UNSUPPORTED_OPS[name],
                           ", ".join(sorted(OPERATIONS))))
                raise SystemExit("unknown operation %r (known: %s)"
                                 % (name, ", ".join(sorted(OPERATIONS))))
            try:
                results.append("%-14s %s" % (name, handler(pkg, root, {**spec, "_op": name})))
            except AnchorError as exc:
                raise SystemExit("operations[%d] (%s): %s" % (i, name, exc))

        pkg.write(DOCUMENT, etree.tostring(root, encoding="unicode",
                                           xml_declaration=False))
        # One pass over every word/*.xml part. declare_all_body_parts()
        # already covers document/header/footer/notes; the wider sweep
        # catches parts an operation fabricated.
        for part in list(pkg._parts):
            if part.startswith("word/") and part.endswith(".xml"):
                pkg.declare_root(part)

        if not dry_run:
            pkg.save(output)

        print("=== OOXML extension report ===")
        print("Source: %s" % source)
        print("Output: %s%s" % (output, "  (dry-run, not written)" if dry_run else ""))
        for line in results:
            print("  " + line)
        print("  parts in package: %d" % len(pkg._order))
        return 0
    finally:
        pkg.close()


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("config")
    parser.add_argument("--dry-run", action="store_true",
                        help="run the operations but don't write the package")
    args = parser.parse_args(argv)
    if not os.path.exists(args.config):
        raise SystemExit("config not found: %s" % args.config)
    return run(args.config, args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
