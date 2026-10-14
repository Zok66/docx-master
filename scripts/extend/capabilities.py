"""XML / binary builders for the Word features the `apply` engine doesn't
write itself.

Everything here returns a fragment (a `<w:p>` paragraph, a run sequence, a
whole part) that the caller splices into the package. Nothing touches the
filesystem or the package — that keeps the builders testable and lets
`extend.py` compose them.

Grouped by where the result goes:

  * paragraph / run fragments  -> document.xml, headerN.xml, footerN.xml
  * sectPr children            -> the trailing `<w:sectPr>` or a break
  * whole parts                -> footnotes.xml, endnotes.xml, chart1.xml,
                                  embeddings/oleObject1.bin, customXml/item1.xml

Element orders come from `_common` (SECTPR_ORDER / PPR_ORDER) because a
sectPr with misordered children makes Word show a repair prompt.
"""

from __future__ import annotations

import math
import os
import re
import struct
import zipfile

from _common import NS, esc

CM = 360000          # EMU per centimetre
PT = 12700           # EMU per point
TWIP_CM = 567        # twips per centimetre

_IDS = [900]


def next_id() -> int:
    _IDS[0] += 1
    return _IDS[0]


# --------------------------------------------------------------------------
# Fields
# --------------------------------------------------------------------------

def field_runs(instruction: str, placeholder: str = "", rpr: str = "") -> str:
    """A complete field as a run sequence.

    `instruction` is the raw field code without the outer braces, e.g.
    `DATE \\@ "yyyy年M月d日"`. `placeholder` is the cached result Word shows
    before the field is updated (and after, if the field is never updated).
    """
    rpr_xml = "<w:rPr>%s</w:rPr>" % rpr if rpr else ""
    return (
        '<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    ) % (rpr_xml, instruction, esc(placeholder))


def simple_field(instruction: str, placeholder: str = "") -> str:
    """A field written in the single-run `<w:fldSimple>` form. Use for
    fields whose result is a plain string (PAGE, REF, SEQ, DOCPROPERTY)."""
    return '<w:fldSimple w:instr=" %s ">%s</w:fldSimple>' % (
        esc(instruction),
        '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % esc(placeholder) if placeholder else "",
    )


def toc_field(levels: str = "1-3", use_outline_levels: bool = True,
              hyperlinks: bool = True, hide_tab_leader: bool = True) -> str:
    """Table-of-contents field.

    The field is generated empty — Word fills it on update (Ctrl+A → F9, or
    automatically because the engine sets `w:updateFields`). An empty
    result here is correct: the data only exists in Word's own layout pass.
    """
    switches = "\\o \"%s\"" % levels if use_outline_levels else "\\t \"%s\"" % levels
    if hyperlinks:
        switches += " \\h"
    if hide_tab_leader:
        switches += " \\z"
    return field_runs("TOC %s" % switches, "（请在此处右键更新域以生成目录）")


def toc_of_field(caption_identifier: str, label: str = "图") -> str:
    """A list-of-figures / list-of-tables field (`TOC \\c "<identifier>"`)."""
    return field_runs(
        'TOC \\h \\z \\c "%s"' % caption_identifier,
        "（右键更新域以生成%s目录）" % label,
    )


def index_field(columns: int = 1) -> str:
    """Alphabetical index field. Pair with `index_entry()` markers."""
    return field_runs("INDEX \\c \"%d\"" % columns if columns else "INDEX",
                      "（右键更新域以生成索引）")


def index_entry(term: str) -> str:
    """A hidden `XE` marker paragraph carrying one index term."""
    return (
        '<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:rPr><w:vanish/></w:rPr><w:instrText xml:space="preserve"> XE "%s" </w:instrText></w:r>'
        '<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:p>'
    ) % esc(term)


def page_number_field(fmt: str = "PAGE") -> str:
    """PAGE / NUMPAGES / SECTIONPAGES field as inline runs."""
    return field_runs("%s" % fmt, "1")


# --------------------------------------------------------------------------
# Section properties
# --------------------------------------------------------------------------

def pg_num_type(fmt: str, start: int) -> str:
    return '<w:pgNumType w:fmt="%s" w:start="%d"/>' % (fmt, start)


def cols(num: int, space: int = 425, separator: bool = False) -> str:
    sep = ' w:sep="1"' if separator else ""
    return '<w:cols w:num="%d" w:space="%d"%s/>' % (num, space, sep)


def text_direction(value: str = "tbRl") -> str:
    """`tbRl` = vertical CJK (top-to-bottom, right-to-left columns)."""
    return '<w:textDirection w:val="%s"/>' % value


def pg_borders(val: str = "single", sz: int = 12, space: int = 24,
               color: str = "4472C4", offset_from: str = "page") -> str:
    sides = "".join(
        '<w:%s w:val="%s" w:sz="%d" w:space="%d" w:color="%s"/>' % (side, val, sz, space, color)
        for side in ("top", "left", "bottom", "right")
    )
    return '<w:pgBorders w:offsetFrom="%s">%s</w:pgBorders>' % (offset_from, sides)


def section_properties(kind: str = "nextPage", pg_sz=None, pg_mar=None,
                       pg_num=None, columns=None, direction=None,
                       valign=None, title_pg=False, doc_grid=True) -> str:
    """Assemble a `<w:sectPr>`.

    `pg_sz` / `pg_mar` are dicts of attribute name -> value (`w:w`, `w:h`,
    `w:orient` / `w:top`, `w:right`, `w:bottom`, `w:left`, `w:gutter`).
    `pg_num` is `(format, start)`; `columns` is `(num, space, separator)`.
    `valign` sets section vertical alignment (`"center"` for cover pages);
    `title_pg` writes `<w:titlePg/>` (first-page-different header/footer —
    with no first-page headerReference declared, Word renders the first
    page header/footer empty, which is exactly what a cover wants).

    Child order follows CT_SectPr: type, pgSz, pgMar, pgNumType, cols,
    vAlign, titlePg, textDirection, docGrid.
    """
    parts = []
    if kind:
        parts.append('<w:type w:val="%s"/>' % kind)
    if pg_sz:
        parts.append("<w:pgSz %s/>" % " ".join('w:%s="%s"' % (k, v) for k, v in pg_sz.items()))
    if pg_mar:
        parts.append("<w:pgMar %s/>" % " ".join('w:%s="%s"' % (k, v) for k, v in pg_mar.items()))
    if pg_num:
        parts.append(pg_num_type(pg_num[0], pg_num[1]))
    if columns:
        parts.append(cols(columns[0], columns[1], columns[2]))
    if valign:
        parts.append('<w:vAlign w:val="%s"/>' % valign)
    if title_pg:
        parts.append("<w:titlePg/>")
    if direction:
        parts.append(text_direction(direction))
    if doc_grid:
        parts.append('<w:docGrid w:linePitch="312"/>')
    return "<w:sectPr>%s</w:sectPr>" % "".join(parts)


# --------------------------------------------------------------------------
# Page-level decoration in settings.xml
# --------------------------------------------------------------------------

def background_shape(fill: str = "F5F7FA") -> str:
    """`<w:background>` element for settings.xml (serialised into
    document.xml as the first child of `<w:document>`)."""
    return ('<w:background w:color="%s"><w:drawing><wp:anchor behindDoc="1" '
            'relativeHeight="0" allowOverlap="1" layoutInCell="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="Background"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
            '<a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
            '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:background>'
            ) % (fill, 12240 * 635, 15840 * 635, next_id(), 12240 * 635, 15840 * 635, fill)


# --------------------------------------------------------------------------
# Watermark
# --------------------------------------------------------------------------

def watermark_paragraph(text: str = "内部资料", color: str = "C8C8C8",
                        font: str = "宋体", rotation: int = 315,
                        size_pt: float = 1.0, opacity: str = ".5") -> str:
    """A Word-style diagonal text watermark (VML `v:shape` inside `w:pict`).

    Rectangles inherited from the PowerPlus watermark Word itself writes,
    parameterised. `size_pt` stays tiny on purpose — VML `textpath`
    stretches the glyphs to the shape, so the nominal size barely matters.
    """
    return (
        '<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:noProof/></w:rPr><w:pict>'
        '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" '
        'path="m@7,l@8,m@5,21600l@6,21600e">'
        '<v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/>'
        '<v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>'
        '<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/>'
        '<v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>'
        '<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas>'
        '<v:path textpathok="t" o:connecttype="custom" '
        'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/>'
        '<v:textpath on="t" fitshape="t"/></v:shapetype>'
        '<v:shape id="PowerPlusWaterMarkObject%d" o:spid="_x0000_s%d" type="#_x0000_t136" '
        'style="position:absolute;margin-left:0;margin-top:0;width:414.05pt;height:103.5pt;'
        'rotation:%d;z-index:-251658752;mso-position-horizontal:center;'
        'mso-position-horizontal-relative:margin;mso-position-vertical:center;'
        'mso-position-vertical-relative:margin" o:allowincell="f" '
        'fillcolor="#%s" stroked="f">'
        '<v:fill opacity="%s"/>'
        '<v:textpath style="font-family:&quot;%s&quot;;font-size:%.1fpt" string="%s"/>'
        '</v:shape></w:pict></w:r></w:p>'
    ) % (next_id(), next_id(), rotation, color, opacity, font, size_pt, esc(text))


# --------------------------------------------------------------------------
# Notes
# --------------------------------------------------------------------------

FOOTNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
ENDNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
FOOTNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
ENDNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"


def note_part(root_tag: str = "w:footnotes", separator_text: str = "",
              continuation_text: str = "") -> str:
    """Skeleton for a notes part.

    Word requires ids -1 (separator) and 0 (continuationSeparator) to exist
    before any real note; a part with only real notes renders but shows an
    odd separator line.
    """
    children = ""
    if separator_text:
        children += ('<w:%s w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" '
                     'w:line="240" w:lineRule="auto"/></w:pPr><w:r>'
                     '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>'
                     '<w:%s/>'
                     '</w:r></w:p></w:%s>'
                     % ("footnote" if root_tag == "w:footnotes" else "endnote",
                        "separator" if root_tag == "w:footnotes" else "separator",
                        "footnote" if root_tag == "w:footnotes" else "endnote"))
    if continuation_text:
        children += ('<w:%s w:type="continuationSeparator" w:id="0"><w:p><w:pPr>'
                     '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>'
                     '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>'
                     '<w:%s/></w:r></w:p></w:%s>'
                     % ("footnote" if root_tag == "w:footnotes" else "endnote",
                        "continuationSeparator", "footnote" if root_tag == "w:footnotes" else "endnote"))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<%s xmlns:w="%s" xmlns:r="%s" xmlns:v="%s" xmlns:o="%s" xmlns:mc="%s" '
            'xmlns:w14="%s" xmlns:w15="%s" xmlns:wps="%s" xmlns:wp14="%s" '
            'mc:Ignorable="w14 w15 wp14">%s</%s>'
            % (root_tag, NS["w"], NS["r"], NS["v"], NS["o"], NS["mc"], NS["w14"],
               NS["w15"], NS["wps"], NS["wp14"], children, root_tag))


def note_entry(kind: str, note_id: int, text: str) -> str:
    """One `<w:footnote>` / `<w:endnote>` body."""
    tag = "footnote" if kind == "footnote" else "endnote"
    ref = "footnoteRef" if kind == "footnote" else "endnoteRef"
    return (
        '<w:%s w:id="%d"><w:p><w:pPr><w:pStyle w:val="%s"/>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="%s"/></w:rPr><w:%s/></w:r>'
        '<w:r><w:t xml:space="preserve"> %s</w:t></w:r></w:p></w:%s>'
    ) % (tag, note_id, "FootnoteText" if kind == "footnote" else "EndnoteText",
         "FootnoteReference" if kind == "footnote" else "EndnoteReference",
         ref, esc(text), tag)


def note_reference_run(kind: str, note_id: int) -> str:
    """The in-body reference mark that points at a note."""
    ref = "footnoteReference" if kind == "footnote" else "endnoteReference"
    style = "FootnoteReference" if kind == "footnote" else "EndnoteReference"
    return ('<w:r><w:rPr><w:rStyle w:val="%s"/><w:vertAlign w:val="%s"/></w:rPr>'
            '<w:%s w:id="%d"/></w:r>' % (style, "superscript", ref, note_id))


# --------------------------------------------------------------------------
# Drawings
# --------------------------------------------------------------------------

_URI_BY_FRAGMENT = [
    ("<wpc:wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"),
    ("<wpg:wgp", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
    ("<pic:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
    ("<c:chart", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
]


def graphic_data_uri(fragment: str) -> str:
    for needle, uri in _URI_BY_FRAGMENT:
        if needle in fragment:
            return uri
    return "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def wrap_inline(fragment: str, cx_cm: float, cy_cm: float, name: str,
                centered: bool = True) -> str:
    """Wrap a `wps` / `wpc` / `pic` fragment as an inline drawing in a
    paragraph."""
    cx, cy = int(cx_cm * CM), int(cy_cm * CM)
    jc = '<w:pPr><w:jc w:val="center"/></w:pPr>' if centered else ""
    return (
        '<w:p>%s<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%d" name="%s"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="%s">%s</a:graphicData></a:graphic>'
        '</wp:inline></w:drawing></w:r></w:p>'
    ) % (jc, cx, cy, next_id(), esc(name), graphic_data_uri(fragment), fragment)


def textbox_paragraphs(text: str, size: int = 18, color: str = "1F3A5F",
                       bold: bool = False, font: str = "微软雅黑") -> str:
    """Paragraphs for use inside a shape's `w:txbxContent`."""
    return (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr>'
        '<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>%s'
        '<w:color w:val="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/>'
        '</w:rPr><w:t>%s</w:t></w:r></w:p>'
    ) % (font, font, font, "<w:b/>" if bold else "", color, size, size, esc(text))


def shape(preset: str = "roundRect", cx_cm: float = 3.4, cy_cm: float = 1.6,
          text: str = "", fill: str = "E8F0FE", line: str = "4A6FA5",
          line_width: int = 12700, inner_inset: int = 45720,
          text_warp: str = None, offset_cm=(0.0, 0.0),
          paragraphs: str = None, anchor: str = "ctr",
          body_insets=None) -> str:
    """A DrawingML preset-geometry shape whose text stays editable.

    `text_warp` adds `<a:prstTxWarp>` — the WordArt deformations
    (`textWave1`, `textArchUp`, `textInflate`…). Same element tree as a
    plain shape, so it remains selectable and re-typable in Word.

    `offset_cm` positions the shape inside a canvas / group coordinate
    space (absolute `a:off`). `paragraphs` replaces the single-text
    paragraph template with pre-built `<w:p>` runs; `body_insets` is
    `(l, t, r, b)` in EMU (defaults to 45720 = 0.36 mm ≈ 0.05")."""
    cx, cy = int(cx_cm * CM), int(cy_cm * CM)
    ox, oy = int(offset_cm[0] * CM), int(offset_cm[1] * CM)
    warp = ("" if not text_warp else
            '<a:prstTxWarp prst="%s"><a:avLst/></a:prstTxWarp>' % text_warp)
    if body_insets is None:
        body_insets = (inner_inset, 45720, inner_inset, 45720)
    l_ins, t_ins, r_ins, b_ins = [int(v) for v in body_insets]
    inner = paragraphs if paragraphs is not None else textbox_paragraphs(text)
    return (
        '<wps:wsp><wps:cNvSpPr/>'
        '<wps:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
        '</wps:spPr>'
        '<wps:txbx><w:txbxContent>%s</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" wrap="square" lIns="%d" tIns="%d" rIns="%d" '
        'bIns="%d" anchor="%s" anchorCtr="0">%s<a:noAutofit/></wps:bodyPr>'
        '</wps:wsp>'
    ) % (ox, oy, cx, cy, preset, fill, line_width, line,
         inner, l_ins, t_ins, r_ins, b_ins, anchor, warp)


def text_box(paragraphs_xml: str, cx_cm: float = 5.0, cy_cm: float = 2.0,
             fill: str = None, line: str = None, line_width: int = 12700,
             body_insets=(91440, 45720, 91440, 45720), anchor: str = "t",
             wrap: str = "square", offset_cm=(0.0, 0.0),
             rot_deg: float = 0.0) -> str:
    """A real Word text box — `wps:cNvSpPr txBox="1"` + `w:txbxContent`.

    Unlike a filled shape, a text box carries no default fill/outline
    (pass `fill` / `line` to add them) and defaults to top-anchored,
    auto-wrapping text with 0.1" side insets. `paragraphs_xml` is the
    pre-built `<w:p>` sequence (see `rich_paragraphs`)."""
    cx, cy = int(cx_cm * CM), int(cy_cm * CM)
    ox, oy = int(offset_cm[0] * CM), int(offset_cm[1] * CM)
    l_ins, t_ins, r_ins, b_ins = [int(v) for v in body_insets]
    fill_xml = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill
                if fill else "<a:noFill/>")
    line_xml = ('<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
                % (line_width, line) if line else "<a:ln><a:noFill/></a:ln>")
    rot = (' rot="%d"' % int(rot_deg * 60000)) if rot_deg else ""
    return (
        '<wps:wsp><wps:cNvSpPr txBox="1"/>'
        '<wps:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>%s%s</wps:spPr>'
        '<wps:txbx><w:txbxContent>%s</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0"%s wrap="%s" lIns="%d" tIns="%d" rIns="%d" bIns="%d" '
        'anchor="%s" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        '</wps:wsp>'
    ) % (ox, oy, cx, cy, fill_xml, line_xml, paragraphs_xml, rot, wrap,
         l_ins, t_ins, r_ins, b_ins, anchor)


def rich_paragraphs(items) -> str:
    """Build `<w:p>` runs for a text box / shape body.

    `items` is a list of dicts:
      { text, size?, bold?, italic?, color?, font?, align? }
    `size` is half-points (21 = 10.5pt), `align` ∈ left|center|right|both.
    A plain string is accepted as shorthand for `{text: ...}`."""
    out = []
    for it in items:
        if isinstance(it, str):
            it = {"text": it}
        align = it.get("align", "left")
        size = int(it.get("size", 21))
        font = it.get("font", "微软雅黑")
        color = it.get("color", "000000")
        rpr = ('<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s"/>'
               '%s%s<w:color w:val="%s"/><w:sz w:val="%d"/><w:szCs w:val="%d"/>'
               '</w:rPr>' % (font, font, font,
                             "<w:b/>" if it.get("bold") else "",
                             "<w:i/>" if it.get("italic") else "",
                             color, size, size))
        out.append('<w:p><w:pPr><w:jc w:val="%s"/></w:pPr><w:r>%s'
                   '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                   % (align, rpr, esc(str(it.get("text", "")))))
    return "".join(out)


def connector(cx_cm: float, cy_cm: float, flip_vertical: bool = False,
              arrow: bool = True, color: str = "4A6FA5", width: int = 19050,
              offset_cm=(0.0, 0.0), flip_horizontal: bool = False) -> str:
    cx, cy = int(cx_cm * CM), int(cy_cm * CM)
    ox, oy = int(offset_cm[0] * CM), int(offset_cm[1] * CM)
    flips = (' flipH="1"' if flip_horizontal else "") + \
            (' flipV="1"' if flip_vertical else "")
    tail = '<a:tailEnd type="triangle" w="med" len="med"/>' if arrow else ""
    return (
        '<wps:wsp><wps:cNvCnPr/>'
        '<wps:spPr><a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</a:ln>'
        '</wps:spPr><wps:bodyPr/></wps:wsp>'
    ) % (flips, ox, oy, cx, cy, width, color, tail)


def line_between(x1_cm: float, y1_cm: float, x2_cm: float, y2_cm: float,
                 color: str = "4A6FA5", arrow: bool = True,
                 width: int = 19050) -> str:
    """A straight connector from (x1,y1) to (x2,y2) — arbitrary direction
    via `flipH` / `flipV` (the shape's bare geometry is always drawn from
    its top-left corner before flipping)."""
    return connector(abs(x2_cm - x1_cm), abs(y2_cm - y1_cm),
                     flip_vertical=y2_cm < y1_cm,
                     flip_horizontal=x2_cm < x1_cm,
                     arrow=arrow, color=color, width=width,
                     offset_cm=(min(x1_cm, x2_cm), min(y1_cm, y2_cm)))


def smart_art(kind: str, nodes, width_cm: float = 12.0, height_cm: float = None,
              fill: str = "DCE9F9", line: str = "4A6FA5",
              text_color: str = "1F3A5F", font: str = "微软雅黑",
              size: int = 18, arrow: bool = True, name: str = None) -> str:
    """Layout a SmartArt-equivalent diagram as a `wpc` canvas of vector
    shapes + connectors.

    Semantics (`kind`):
      * `process`   — left-to-right chain of rounded boxes with arrows
      * `list`      — vertical stack with a connecting spine
      * `cycle`     — nodes on a circle, adjacent nodes linked
      * `hierarchy` — one root on top, children evenly spread below

    Every node keeps editable text, so the result behaves like a real
    SmartArt for editing purposes; it is *not* a `dgm` SmartArt object
    (see references/extend.md for why that is not attempted).
    """
    labels = [n if isinstance(n, str) else str(n.get("text", ""))
              for n in nodes]
    n = len(labels)
    if n == 0:
        raise ValueError("smart_art needs at least one node")
    kind = (kind or "process").lower()
    if kind not in ("process", "list", "cycle", "hierarchy"):
        raise ValueError("smart_art kind must be process / list / cycle / "
                         "hierarchy (got %r)" % kind)

    def box(x, y, w, h, text, preset="roundRect"):
        body = rich_paragraphs([{"text": text, "size": size, "color": text_color,
                                 "font": font, "align": "center"}])
        return shape(preset, w, h, fill=fill, line=line,
                     offset_cm=(x, y), paragraphs=body)

    W = float(width_cm)
    gap = 0.5
    children = []

    if kind == "process":
        h = float(height_cm) if height_cm else 1.6
        H = h
        w = (W - gap * (n - 1)) / n
        y = (H - h) / 2.0
        for i, label in enumerate(labels):
            x = i * (w + gap)
            children.append(box(x, y, w, h, label))
            if i < n - 1 and arrow:
                children.append(line_between(x + w + 0.06, H / 2.0,
                                             x + w + gap - 0.06, H / 2.0))

    elif kind == "list":
        h, v_gap = 0.95, 0.35
        H = float(height_cm) if height_cm else (n * h + (n - 1) * v_gap)
        spine_x = 0.34
        for i, label in enumerate(labels):
            y = i * (h + v_gap)
            children.append(box(spine_x, y, W - spine_x, h, label,
                                preset="roundRect"))
            if i < n - 1 and arrow:
                children.append(line_between(spine_x + 0.34, y + h,
                                             spine_x + 0.34,
                                             y + h + v_gap, arrow=False))

    elif kind == "cycle":
        H = float(height_cm) if height_cm else min(W, 6.2)
        d = min(W, H) / 2.55
        R = (min(W, H) - d) / 2.0
        centers = []
        for i in range(n):
            ang = -math.pi / 2 + 2 * math.pi * i / n
            cx = W / 2.0 + R * math.cos(ang) - d / 2.0
            cy = H / 2.0 + R * math.sin(ang) - d / 2.0
            centers.append((cx + d / 2.0, cy + d / 2.0))
            children.append(box(cx, cy, d, d, labels[i], preset="ellipse"))
        if arrow:
            for i in range(n):
                x1, y1 = centers[i]
                x2, y2 = centers[(i + 1) % n]
                if (i + 1) % n == 0 and n > 2:
                    break  # closing segment would cut across the circle
                children.append(line_between(x1, y1, x2, y2))

    else:  # hierarchy
        H = float(height_cm) if height_cm else 2.9
        h = 0.95
        root_w = min(W * 0.5, 3.6)
        children.append(box((W - root_w) / 2.0, 0.0, root_w, h, labels[0]))
        kids = labels[1:]
        if kids:
            kw = (W - gap * (len(kids) - 1)) / len(kids)
            bus_y = h + (H - 2 * h) / 2.0
            top_bus = h + 0.12
            bot_bus = H - h - 0.12
            bus_y = min(max(bus_y, top_bus), bot_bus) if bot_bus > top_bus else H / 2.0
            for i, label in enumerate(kids):
                x = i * (kw + gap)
                children.append(box(x, H - h, kw, h, label))
                cx_kid = x + kw / 2.0
                children.append(line_between(cx_kid, bus_y, cx_kid, H - h,
                                             arrow=False))
            children.append(line_between(W / 2.0, h, W / 2.0, bus_y,
                                         arrow=False))
            children.append(line_between(kw / 2.0, bus_y,
                                         W - kw / 2.0, bus_y, arrow=False))

    return canvas(children, W, H, name=name or ("SmartArt · %s" % kind)), W, H


def canvas(inner_shapes, cx_cm: float = 8.0, cy_cm: float = 3.0,
           name: str = "绘图画布") -> str:
    """A `wpc:wpc` canvas. Children live in the canvas coordinate space
    (chOff / chExt), so they move as a group — position each child with
    its own `offset_cm`."""
    inner = "".join(inner_shapes)
    cx, cy = int(cx_cm * CM), int(cy_cm * CM)
    return (
        '<wpc:wpc><wpc:cNvPr id="%d" name="%s"/><wpc:cNvGrpSpPr/><wpc:grpSpPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="%d" cy="%d"/></a:xfrm>'
        '</wpc:grpSpPr>%s</wpc:wpc>'
    ) % (next_id(), esc(name), cx, cy, cx, cy, inner)


def _split_top_level(xml: str):
    """Split a fragment into its top-level element chunks.

    Small regex scanner instead of lxml — capabilities.py stays dependency
    free so the builders can run anywhere `apply` runs.
    """
    chunks = []
    i, n = 0, len(xml)
    while i < n:
        if xml[i] != "<":
            i += 1
            continue
        m = re.match(r"<([A-Za-z_][\w:.\-]*)([^>]*)>", xml[i:])
        if not m:
            i += 1
            continue
        name = m.group(1)
        start = i
        if m.group(2).rstrip().endswith("/"):
            i += m.end()
        else:
            cm = re.compile(r"</%s\s*>" % re.escape(name)).search(xml, i + m.end())
            i = cm.end() if cm else i + m.end()
        chunks.append(xml[start:i])
    return chunks


# CT_EffectList is an xsd:sequence — order is enforced, not optional.
_EFFECT_ORDER = ("blur", "fillOverlay", "glow", "innerShdw", "outerShdw",
                 "prstShdw", "reflection", "softEdge")


def _order_effect_lst(body_xml: str) -> str:
    """Sort <a:effectLst> children into schema order.

    Word itself is lenient, but any XSD validator rejects e.g. a glow placed
    after an outerShdw, so normalise here rather than making callers know the
    sequence.
    """
    chunks = _split_top_level(body_xml)
    if len(chunks) < 2:
        return body_xml
    strays = [c for c in chunks if not c.lstrip().startswith("<a:")]
    if strays:
        return body_xml
    def rank(chunk):
        m = re.match(r"<a:(\w+)", chunk.lstrip())
        name = m.group(1) if m else ""
        return _EFFECT_ORDER.index(name) if name in _EFFECT_ORDER else len(_EFFECT_ORDER)
    return "".join(sorted(chunks, key=rank))


def picture_run(relationship_id: str, cx_cm: float, cy_cm: float, name: str,
                crop=None, lum=None, effects: str = "", floating: bool = False,
                wrap_text: str = "bothSides", offset_x_cm: float = 0.0,
                offset_y_cm: float = 0.0, allow_overlap: bool = True,
                behind_doc: bool = False) -> str:
    """A picture as a paragraph.

    `crop` is a 4-tuple of per-mille insets `(left, top, right, bottom)`;
    `lum` is `(brightness, contrast)`; `effects` is raw `<a:effectLst>`
    content (glow, shadow, softEdge). `floating=True` produces a
    `<wp:anchor>` with a wrap element instead of an inline drawing.

    Floating placement (`offset_x_cm` / `offset_y_cm` /
    `allow_overlap` / `behind_doc`): the anchor is positioned from the
    column and the line, so a zero offset pins the picture to the anchor
    paragraph's left edge. Two floating objects anchored in the same
    paragraph therefore land on the very same spot — and because
    `allowOverlap` defaults to on, Word stacks them there instead of
    pushing one aside. Give each floating object its own offset, or pass
    `allow_overlap=False` to forbid stacking.

    A floating picture's wrap also governs how *following* content flows:
    with `wrapSquare` a picture that eats most of the column width leaves
    too little room for wide inline drawings on the same page, and Word
    squeezes those over the right margin. Prefer inline drawings for
    layout-critical figures and reserve floating for a deliberate
    text-wrapping effect.
    """
    cx, cy = int(cx_cm * CM), int(cy_cm * CM)
    off_x, off_y = int(offset_x_cm * CM), int(offset_y_cm * CM)
    src_rect = '<a:srcRect l="%d" t="%d" r="%d" b="%d"/>' % crop if crop else ""
    lum_xml = '<a:lum bright="%d" contrast="%d"/>' % lum if lum else ""
    # `effects` is documented as the inner content of <a:effectLst>, but a
    # full wrapper is a natural thing to paste in — strip it rather than
    # emit a nested <a:effectLst> (which fails XSD).
    body_xml = effects.strip()
    if body_xml.startswith("<a:effectLst"):
        body_xml = body_xml[body_xml.find(">") + 1:]
        tail = "</a:effectLst>"
        if body_xml.endswith(tail):
            body_xml = body_xml[:-len(tail)]
    body_xml = _order_effect_lst(body_xml)
    effect_xml = "<a:effectLst>%s</a:effectLst>" % body_xml if body_xml else ""
    pic = (
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s">%s</a:blip>%s'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>%s</pic:spPr></pic:pic>'
    ) % (next_id(), esc(name), relationship_id, lum_xml, src_rect, cx, cy, effect_xml)
    if not floating:
        return (
            '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
            '<wp:inline distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:docPr id="%d" name="%s"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '%s</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
        ) % (cx, cy, next_id(), esc(name), pic)
    # ST_WrapText is {bothSides, left, right, largest} — "square" is the
    # *shape* of the wrap, not the text-flow value. Accept the colloquial
    # "square" (Word's 四周型环绕) and switch the element to wrapSquare.
    wrap_style = "tight"
    if wrap_text in ("square", "wrapSquare"):
        wrap_style, wrap_text = "square", "bothSides"
    elif wrap_text not in ("bothSides", "left", "right", "largest"):
        wrap_text = "bothSides"
    if wrap_style == "square":
        wrap_xml = '<wp:wrapSquare wrapText="%s"/>' % wrap_text
    else:
        wrap_xml = ('<wp:wrapTight wrapText="%s"><wp:wrapPolygon edited="0">'
                    '<wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>'
                    '<wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>'
                    '<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>') % wrap_text
    return (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        'relativeHeight="251658240" behindDoc="%d" locked="0" layoutInCell="1" allowOverlap="%d">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>%d</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="line"><wp:posOffset>%d</wp:posOffset></wp:positionV>'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '%s'
        '<wp:docPr id="%d" name="%s"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '%s</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'
    ) % (1 if behind_doc else 0, 1 if allow_overlap else 0, off_x, off_y,
         cx, cy, wrap_xml, next_id(), esc(name), pic)


def chart_reference_run(relationship_id: str, cx_cm: float = 11.0,
                        cy_cm: float = 5.5, name: str = "图表") -> str:
    """A `<c:chart>` reference wrapped as an inline drawing. The chart part
    itself comes from `chart_part()`."""
    return wrap_inline(
        '<c:chart xmlns:c="%s" xmlns:r="%s" r:id="%s"/>' % (NS["c"], NS["r"], relationship_id),
        cx_cm, cy_cm, name,
    )


def ink_run(points, color: str = "1F3A5F", width_pt: float = 2.0,
            cx_cm: float = 5.0, cy_cm: float = 2.0) -> str:
    """Freehand ink as a VML polyline (`v:shape` with a `v:path`).

    Word's native ink uses InkML plus a `<w14:ink>` extension; a VML
    polyline renders identically for static output and stays in the parts
    that basic schema covers, at the cost of losing stroke pressure /
    replay. `points` is a list of (x, y) in a 0..1000 space. JSON configs
    deliver lists, so coerce each point to a tuple for `%` formatting.
    """
    points = [tuple(p) for p in points]
    # VML style wants points, the API takes centimetres — 72/2.54, not 72
    # (assuming inches here inflated every stroke box by 2.54x).
    box_w_pt = cx_cm * 72.0 / 2.54
    box_h_pt = cy_cm * 72.0 / 2.54
    # An explicit coordsize keeps the path 1:1 with the point space; without
    # it VML falls back to its own default grid and the stroke is rescaled.
    coord_x = max([p[0] for p in points] + [1])
    coord_y = max([p[1] for p in points] + [1])
    d = "m%d,%d" % (points[0][0], points[0][1]) + "".join(
        " l%d,%d" % (x, y) for x, y in points[1:]
    )
    return (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>'
        '<w:pict><v:shape id="inkStroke%d" type="#_x0000_t201" '
        'style="width:%.1fpt;height:%.1fpt" coordsize="%d,%d" '
        'strokecolor="#%s" strokeweight="%.1fpt" '
        'fillcolor="none"><v:path shadowok="f" strokeok="t" fillok="f" '
        'gradientshapeok="f" textpathok="f" o:connecttype="custom" '
        'o:connectlocs="%s" v="%s"/></v:shape></w:pict></w:r></w:p>'
    ) % (next_id(), box_w_pt, box_h_pt, coord_x, coord_y, color, width_pt,
         ";".join("%d,%d" % p for p in points), d)


def symbol_char_code(value) -> str:
    """Normalise a `w:sym@w:char` value to `ST_ShortHexNumber`.

    The attribute is a 1-4 digit **hex** token, so the literal glyph
    ("→"), a `U+` spelling ("U+2192") and a decimal code point are all
    invalid as written and fail XSD validation with an unhelpful message.
    Accept every reasonable spelling and emit the canonical one:

        "→"  ·  "U+2192"  ·  "0x2192"  ·  "2192"  →  "2192"
        8594 (int)                                 →  "2192"
        167                                        →  "00A7"

    Rules that resolve the ambiguity between hex and decimal strings:
    4 characters or fewer made of hex digits is *already* a code — that is
    exactly the shape ST_ShortHexNumber allows, and how every symbol table
    (Wingdings "F0FC", "2192") writes them. Five digits or more of pure
    digits cannot be a code, so it is read as a decimal code point.
    """
    if isinstance(value, bool):
        raise ValueError("symbol char must be a hex code, U+ code, decimal "
                         "code point or the glyph itself — got %r" % (value,))
    if isinstance(value, int):
        cp = value
    else:
        text = str(value).strip()
        if not text:
            raise ValueError("symbol char is empty")
        low = text.lower()
        if low.startswith("u+"):
            cp = int(low[2:], 16)
        elif low.startswith("0x"):
            cp = int(low[2:], 16)
        elif len(text) <= 4 and all(c in "0123456789abcdefABCDEF" for c in text):
            return text.upper()
        elif text.isdigit():
            cp = int(text)
        elif len(text) == 1:
            cp = ord(text)
        else:
            raise ValueError("symbol char %r is not usable: pass a single "
                             "glyph, its hex code (\"2192\"), a U+ spelling "
                             "(\"U+2192\") or a decimal code point" % (text,))
    if cp < 0 or cp > 0xFFFF:
        raise ValueError("symbol char U+%04X is outside ST_ShortHexNumber — "
                         "the attribute only carries one UTF-16 code unit "
                         "(BMP, U+0000–U+FFFF)" % cp)
    return "%04X" % cp if cp > 0xFF else "%02X" % cp


def symbol_run(font: str = "Wingdings", char: str = "F0FC", size_half_pt: int = 28) -> str:
    """A `w:sym` glyph run (Windings / Webdings / Symbol).

    `char` accepts the glyph itself, a hex code, a `U+` spelling or a
    decimal code point — see `symbol_char_code()` for the rules.

    Note the trailing `\u3000` sits in its own run in `symbol_paragraph` —
    a full-width space after the last run gets parsed as character data of
    `<w:p>` and fails the element-only content model.
    """
    code = symbol_char_code(char)
    return ('<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/>'
            '<w:sz w:val="%d"/></w:rPr><w:sym w:font="%s" w:char="%s"/></w:r>'
            ) % (font, font, size_half_pt, font, code)


def symbol_paragraph(symbols, style: str = "BodyText") -> str:
    """A paragraph of symbol runs; spacing is carried by dedicated runs so
    no bare text lands directly inside `<w:p>`."""
    runs = ""
    for i, (font, char) in enumerate(symbols):
        runs += symbol_run(font, char)
        if i < len(symbols) - 1:
            runs += '<w:r><w:t xml:space="preserve">　</w:t></w:r>'
    return ('<w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>%s</w:p>' % (style, runs))


# --------------------------------------------------------------------------
# Charts (part + data source)
# --------------------------------------------------------------------------

CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
CHART_COLOR_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"


def chart_part(categories, series, title: str = "", relation_id: str = "rIdChartData",
               kind: str = "bar") -> str:
    """A minimal single-plot chart part with cached values.

    Cached values matter: without them Word shows an empty plot until the
    embedded workbook is loaded. Child order inside `c:ser` /
    `c:dLbls` is schema-relevant — `c:txPr` is NOT a valid `c:dLbls`
    child, so text formatting goes on the plot area instead.
    """
    cats = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, esc(c))
                   for i, c in enumerate(categories))
    sers = ""
    for s_i, s in enumerate(series):
        vals = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v)
                       for i, v in enumerate(s["values"]))
        sers += (
            '<c:ser><c:idx val="%d"/><c:order val="%d"/>'
            '<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/>'
            '<c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>'
            '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$%d</c:f><c:strCache>'
            '<c:ptCount val="%d"/>%s</c:strCache></c:strRef></c:cat>'
            '<c:val><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f><c:numCache>'
            '<c:formatCode>General</c:formatCode><c:ptCount val="%d"/>%s</c:numCache>'
            '</c:numRef></c:val>'
            '</c:ser>'
        ) % (s_i, s_i, esc(s["name"]), len(categories) + 1, len(categories), cats,
             chr(ord("B") + s_i), chr(ord("B") + s_i), len(categories) + 1,
             len(categories), vals)
    title_xml = ""
    if title:
        title_xml = ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>'
                     '<a:t>%s</a:t></a:r></a:p></c:rich></c:tx>'
                     '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
                     % esc(title))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="%s" xmlns:a="%s" xmlns:r="%s">'
        '<c:chart>%s<c:plotArea><c:layout/><c:%sChart>'
        '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
        '%s<c:dLbls><c:showLegendKey val="0"/>'
        '<c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/>'
        '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
        '<c:gapWidth val="150"/>'
        '<c:axId val="111111111"/><c:axId val="222222222"/></c:%sChart>'
        '<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>'
        '<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>'
        '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/></c:chart><c:externalData r:id="%s">'
        '<c:autoUpdate val="0"/></c:externalData></c:chartSpace>'
    ) % (NS["c"], NS["a"], NS["r"], title_xml, kind, sers, kind, relation_id)


def minimal_xlsx(sheet_rows) -> bytes:
    """A minimal but valid .xlsx whose sheet holds `sheet_rows` (list of
    lists; strings become inline strings).

    Used as a chart's external data and as an OLE payload — Word and Excel
    both open it, and `zipfile` padding pushes the archive past 4 KB so the
    OLE writer can store it in the regular stream instead of the mini
    stream.
    """
    import io

    rows_xml = []
    for r_i, row in enumerate(sheet_rows, start=1):
        cells = []
        for c_i, val in enumerate(row):
            col = chr(ord("A") + c_i)
            if isinstance(val, str):
                cells.append('<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>'
                             % (col, r_i, esc(val)))
            else:
                cells.append('<c r="%s%d"><v>%s</v></c>' % (col, r_i, val))
        rows_xml.append('<row r="%d">%s</row>' % (r_i, "".join(cells)))
    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             '<sheetData>%s</sheetData></worksheet>' % "".join(rows_xml))
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          '</Types>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="%s">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>' % NS["rel"])
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="%s"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
          '</workbook>' % NS["r"])
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="%s">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
               '</Relationships>' % NS["rel"])
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
        z.comment = b"\x00" * 4200
    return buf.getvalue()


# --------------------------------------------------------------------------
# OLE2 compound file (for w:object → o:OLEObject)
# --------------------------------------------------------------------------

OLE_CT = "application/vnd.openxmlformats-officedocument.oleObject"
OLE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"


def ole_compound_file(stream_name: str, payload: bytes, pad_to: int = 4096) -> bytes:
    """Wrap `payload` in a one-stream OLE2 (CFB) container.

    Word's `<w:object>` points at a `oleObject.bin` that must be a real
    compound file; anything else and the double-click activation path can't
    find the payload. Single-stream is enough for an embedded workbook —
    only the FAT and directory are hand-built, everything else is the
    payload verbatim.
    """
    SECTOR = 512
    data = payload + b"\x00" * max(0, pad_to - len(payload))
    n_data = max((len(data) + SECTOR - 1) // SECTOR, 1)
    data = data.ljust(n_data * SECTOR, b"\x00")

    # sector 0 = FAT, sector 1 = directory, sectors 2.. = data
    fat = [0xFFFFFFFF] * 128
    fat[0] = 0xFFFFFFFD          # FATSECT
    fat[1] = 0xFFFFFFFE          # end of directory chain
    for i in range(n_data):
        fat[2 + i] = (3 + i) if i < n_data - 1 else 0xFFFFFFFE
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)

    def dir_entry(name, obj_type, start, size, color=1, child=0xFFFFFFFF):
        raw = name.encode("utf-16-le") + b"\x00\x00"
        entry = raw.ljust(64, b"\x00")
        entry += struct.pack("<H", len(raw))
        entry += struct.pack("<B", obj_type)
        entry += struct.pack("<B", color)
        entry += struct.pack("<III", 0xFFFFFFFF, 0xFFFFFFFF, child)
        entry += b"\x00" * 16                      # CLSID
        entry += struct.pack("<I", 0)              # state bits
        entry += struct.pack("<QQ", 0, 0)          # timestamps
        entry += struct.pack("<I", start)
        entry += struct.pack("<Q", size)
        assert len(entry) == 128, len(entry)
        return entry

    directory = (dir_entry("Root Entry", 5, 0xFFFFFFFE, 0, child=1)
                 + dir_entry(stream_name, 2, 2, len(payload))
                 + dir_entry("", 0, 0, 0, color=0)
                 + dir_entry("", 0, 0, 0, color=0))

    header = bytearray(512)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<H", header, 24, 0x003E)     # minor version
    struct.pack_into("<H", header, 26, 0x0003)     # major version (512-byte sectors)
    struct.pack_into("<H", header, 28, 0xFFFE)     # little-endian
    struct.pack_into("<H", header, 30, 0x0009)     # sector shift
    struct.pack_into("<H", header, 32, 0x0006)     # mini sector shift
    struct.pack_into("<I", header, 40, 0)          # directory sector count (v3: 0)
    struct.pack_into("<I", header, 44, 1)          # FAT sector count
    struct.pack_into("<I", header, 48, 1)          # first directory sector
    struct.pack_into("<I", header, 52, 0)          # transaction signature
    struct.pack_into("<I", header, 56, 4096)       # mini stream cutoff
    struct.pack_into("<I", header, 60, 0xFFFFFFFE) # first mini FAT sector
    struct.pack_into("<I", header, 64, 0)          # mini FAT count
    struct.pack_into("<I", header, 68, 0xFFFFFFFE) # first DIFAT sector
    struct.pack_into("<I", header, 72, 0)          # DIFAT count
    difat = [0] + [0xFFFFFFFF] * 108
    for i, v in enumerate(difat):
        struct.pack_into("<I", header, 76 + i * 4, v)

    return bytes(header) + fat_bytes + directory + data


def ole_object_paragraph(ole_rel_id: str, image_rel_id: str, program_id: str = "Package",
                         shape_id: str = "_x0000_i1031", name: str = "嵌入的对象",
                         width_pt: float = 216.0, height_pt: float = 72.0) -> str:
    """`<w:object>` holding an `o:OLEObject`.

    Two constraints drove this shape: ISO `CT_Object` only allows VML /
    `office` children plus `objectEmbed` — Word's own EMBED-field wrapping
    relies on a Microsoft extension schema and fails XSD validation, so the
    VML-preview form is used; and the VML shape must carry the same
    `ShapeID` the `o:OLEObject` references, or double-click activation
    can't find its target.
    """
    return (
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>'
        '<w:object w:dxaOrig="4320" w:dyaOrig="1440">'
        '<v:shape id="%s" type="#_x0000_t75" style="width:%.1fpt;height:%.1fpt">'
        '<v:imagedata r:id="%s" o:title="%s"/></v:shape>'
        '<o:OLEObject Type="Embed" ProgID="%s" ShapeID="%s" DrawAspect="Icon" '
        'ObjectID="_%d" r:id="%s"/>'
        '</w:object></w:r></w:p>'
    ) % (shape_id, width_pt, height_pt, image_rel_id, esc(name), program_id,
         shape_id, next_id(), ole_rel_id)


# --------------------------------------------------------------------------
# Custom XML part
# --------------------------------------------------------------------------

CUSTOM_XML_CT = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
CUSTOM_XML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
CUSTOM_XML_PROPS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"


def custom_xml_item(data_fields: dict, root_name: str = "project",
                    namespace: str = "urn:docx-master:customXml") -> str:
    """A business-data part stored alongside the document.

    Distinct from `docProps` (user-visible metadata): custom XML is the
    designed home for bound content controls and app data.
    """
    body = "".join("<%s>%s</%s>" % (k, esc(v), k) for k, v in data_fields.items())
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<%s xmlns="%s">%s</%s>' % (root_name, namespace, body, root_name))


def custom_xml_item_props(item_id: str = "{00000000-0000-0000-0000-000000000001}") -> str:
    paths = ('<ds:datastoreItem ds:itemID="%s" '
             'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
             '<ds:schemaRefs/></ds:datastoreItem>' % item_id)
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + paths


# --------------------------------------------------------------------------
# Glossary document (自动图文集 / Building Blocks / AutoText)
# --------------------------------------------------------------------------

GLOSSARY_CT = ("application/vnd.openxmlformats-officedocument."
               "wordprocessingml.document.glossary+xml")
GLOSSARY_REL = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                "relationships/glossaryDocument")


def glossary_entry(name: str, blocks_xml: str, gallery: str = "autoTxt",
                   category: str = "General", auto_insert: bool = False) -> str:
    """One `<w:docPart>` building block (ISO 29500 transitional markup:
    `w:name` + `w:category/w:gallery` + `w:docPartBody`).

    `gallery` is the Building Blocks Organizer category the entry shows
    under — `autoTxt` (自动图文集/AutoText), `coverPg`, `tbls`, `pgNum`,
    `hdrs`, `ftrs`, `watermarks`, `txtBox`, `docParts`, ... `auto_insert`
    adds `<w:type w:val="autoExp"/>` (word-level auto-insert on typing);
    the default `normal` type is insert-on-demand via the gallery.
    """
    pr = '<w:docPartPr><w:name w:val="%s"/>' % esc(name)
    if category:
        pr += ('<w:category><w:name w:val="%s"/><w:gallery w:val="%s"/>'
               '</w:category>' % (esc(category), esc(gallery)))
    pr += '<w:types><w:type w:val="%s"/></w:types></w:docPartPr>' % (
        "autoExp" if auto_insert else "normal")
    return '<w:docPart>%s<w:docPartBody>%s</w:docPartBody></w:docPart>' % (
        pr, blocks_xml)


def glossary_part(entries_xml: str) -> str:
    """The whole `word/glossary/document.xml` part (CT_GlossaryDocument)."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:glossaryDocument xmlns:w="%s" xmlns:r="%s" xmlns:mc="%s" '
            'xmlns:w14="%s"><w:docParts>%s</w:docParts></w:glossaryDocument>'
            % (NS["w"], NS["r"], NS["mc"], NS["w14"], entries_xml))


# --------------------------------------------------------------------------
# Index entry (XE), content control (SDT), bibliography (sources.xml)
# --------------------------------------------------------------------------

def xe_runs(term: str) -> str:
    """An `XE` index-entry field as a hidden complex-field run sequence.

    Appended to a body paragraph; the INDEX field picks these up on F9.
    Runs carry `w:vanish` so nothing renders even if a result run leaks in.
    """
    return ('<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:rPr><w:vanish/></w:rPr>'
            '<w:instrText xml:space="preserve"> XE "%s" </w:instrText></w:r>'
            '<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
            % esc(term))


def sdt_paragraph(text: str, alias: str = "", tag: str = "",
                  style: str = "BodyText", dropdown=None) -> str:
    """A block-level rich-text content control (SDT) wrapping one paragraph.

    CT_SdtPr child order: rPr?, alias?, tag?, id?, then the control choice.
    No explicit choice element means Word treats it as plain rich text;
    `dropdown` (a list of {displayText, value}) turns it into a drop-down
    list control — that element sits in EG_SdtPrElement, right after `id`.
    """
    pr = ""
    if alias:
        pr += '<w:alias w:val="%s"/>' % esc(alias)
    if tag:
        pr += '<w:tag w:val="%s"/>' % esc(tag)
    pr += '<w:id w:val="%d"/>' % next_id()
    if dropdown:
        items = "".join('<w:listItem w:displayText="%s" w:value="%s"/>'
                        % (esc(str(i.get("displayText", ""))), esc(str(i.get("value", ""))))
                        for i in dropdown)
        pr += '<w:dropDownList>%s</w:dropDownList>' % items
    run = '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % esc(text) if text else ""
    return ('<w:sdt><w:sdtPr>%s</w:sdtPr>'
            '<w:sdtContent><w:p><w:pPr><w:pStyle w:val="%s"/></w:pPr>%s</w:p></w:sdtContent>'
            '</w:sdt>' % (pr, style, run))


BIB_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
SOURCES_CT = "application/vnd.openxmlformats-officedocument.customXmlSource+xml"
SOURCES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sources"

# CT_SourceType (shared-bibliography.xsd) is an UNORDERED unbounded choice of
# string fields + one Author slot — this list is the XSD's exact choice
# membership, used as a whitelist. Extra Word-UI fields that the ISO schema
# does NOT allow (Number -> Issue, State -> StateProvince, Country ->
# CountryRegion, Corporate-only roles like BookAuthor/Editor/Translator…
# which exist in Word's strict variant) are deliberately absent; keep to
# this list so both ISO-strict and Word-strict validation pass.
SOURCE_FIELD_WHITELIST = [
    "AbbreviatedCaseNumber", "AlbumTitle", "Author", "BookTitle",
    "Broadcaster", "BroadcastTitle", "CaseNumber", "ChapterNumber", "City",
    "Comments", "ConferenceName", "CountryRegion", "Court", "Day",
    "DayAccessed", "Department", "Distributor", "Edition", "Guid",
    "Institution", "InternetSiteTitle", "Issue", "JournalName", "LCID",
    "Medium", "Month", "MonthAccessed", "NumberVolumes", "Pages",
    "PatentNumber", "PeriodicalTitle", "ProductionCompany",
    "PublicationTitle", "Publisher", "RecordingNumber", "RefOrder",
    "Reporter", "ShortTitle", "StandardNumber", "StateProvince", "Station",
    "SourceType", "Tag", "Theater", "ThesisType", "Title", "Type", "URL",
    "Version", "Volume", "Year", "YearAccessed",
]
SOURCE_TYPES = {
    "ArticleInAPeriodical", "Book", "BookSection", "JournalArticle",
    "ConferenceProceedings", "Report", "SoundRecording", "Performance",
    "Art", "DocumentFromInternetSite", "InternetSite", "Film", "Interview",
    "Patent", "ElectronicSource", "Case", "Misc",
}


def is_corporate(name) -> bool:
    """True when `name` has to be written as `<b:Corporate>` rather than a
    `<b:Person>`.

    A dict says so explicitly (`{"corporate": "ISO/IEC JTC 1"}`). For a
    plain string, CJK personal names are at most a few characters, so
    anything longer than 4 characters is read as an organisation.
    """
    if isinstance(name, dict):
        return "corporate" in name
    return len(str(name)) > 4


def corporate_text(name) -> str:
    """The organisation name behind a corporate author entry."""
    if isinstance(name, dict):
        return str(name.get("corporate", ""))
    return str(name)


def _person(name):
    """The *children* of one `CT_PersonType` — Last / First / Middle.

    That is the whole content model of CT_PersonType: it has no Corporate
    child. A corporate name must go one level up, on CT_AuthorType, via
    `source_entry()` — writing `<b:Corporate>` inside `<b:Person>` produces
    a document that Word may repair and the XSD validator rejects.
    Accepts a dict ({last, first, middle}) or a string — a CJK-style
    string splits first char as surname, rest as given name.
    """
    if isinstance(name, dict):
        if "corporate" in name:
            raise ValueError(
                "_person() got a corporate name (%r). CT_PersonType accepts "
                "only Last/First/Middle — route organisations through "
                "is_corporate()/corporate_text() so they land on "
                "CT_AuthorType instead." % name["corporate"])
        parts = []
        for key, tag in (("last", "Last"), ("first", "First"), ("middle", "Middle")):
            if name.get(key):
                parts.append("<b:%s>%s</b:%s>" % (tag, esc(name[key]), tag))
        return "".join(parts)
    text = str(name)
    if is_corporate(text):
        raise ValueError(
            "_person() got a name longer than 4 characters (%r), which is "
            "read as a corporate author. CT_PersonType has no Corporate "
            "child — route it through is_corporate()/corporate_text()." % text)
    return "<b:Last>%s</b:Last><b:First>%s</b:First>" % (esc(text[:1]), esc(text[1:]))


def author_block(authors) -> str:
    """`<b:Author><b:Author>…</b:Author></b:Author>` for a list of people
    and/or organisations.

    People land in `<b:NameList>` as `<b:Person>` (Last/First/Middle);
    organisations land as `<b:Corporate>` — a direct child of
    CT_AuthorType, *not* of CT_PersonType. Both are valid children of the
    CT_AuthorType choice, so mixing them is fine.
    """
    persons, corporates = [], []
    for name in authors:
        if is_corporate(name):
            corporates.append("<b:Corporate>%s</b:Corporate>"
                              % esc(corporate_text(name)))
        else:
            persons.append("<b:Person>%s</b:Person>" % _person(name))
    inner = ""
    if persons:
        inner += "<b:NameList>%s</b:NameList>" % "".join(persons)
    inner += "".join(corporates)
    return ("<b:Author><b:Author>%s</b:Author></b:Author>" % inner)


def source_entry(tag: str, source_type: str, fields: dict,
                 authors=None) -> str:
    """One `<b:Source>` for word/sources.xml. `fields` keys are the b:
    element local names (Title / Year / Publisher / City / JournalName /
    Pages / URL / ...); anything outside SOURCE_FIELD_ORDER is rejected.
    Elements are emitted in CT_SourceType XSD order; authors land in their
    schema slot (after the string fields, before nothing — Author is late
    in the sequence)."""
    if source_type not in SOURCE_TYPES:
        raise ValueError("unknown b:SourceType %r; one of %s"
                         % (source_type, sorted(SOURCE_TYPES)))
    unknown = set(fields) - set(SOURCE_FIELD_WHITELIST)
    if unknown:
        raise ValueError("unknown source fields: %s (see SOURCE_FIELD_WHITELIST)"
                         % sorted(unknown))
    if "Tag" in fields or "SourceType" in fields or "Guid" in fields:
        raise ValueError("Tag / SourceType / Guid are managed by the op, "
                         "not by `fields`")

    def order_of(el_name):
        return SOURCE_FIELD_WHITELIST.index(el_name)

    parts = []
    parts.append((order_of("Tag"), "<b:Tag>%s</b:Tag>" % esc(tag)))
    parts.append((order_of("SourceType"),
                  "<b:SourceType>%s</b:SourceType>" % source_type))
    import uuid
    parts.append((order_of("Guid"),
                  "<b:Guid>{%s}</b:Guid>" % uuid.uuid4().hex.upper()))
    for k, v in fields.items():
        parts.append((order_of(k), "<b:%s>%s</b:%s>" % (k, esc(str(v)), k)))
    if authors:
        parts.append((order_of("Author"), author_block(authors)))
    parts.sort(key=lambda pair: pair[0])
    return "<b:Source>%s</b:Source>" % "".join(xml for _, xml in parts)


def sources_part(entries_xml: str) -> str:
    """Skeleton (or container) for word/sources.xml."""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<b:Sources xmlns:b="%s" xmlns="%s">%s</b:Sources>'
            % (BIB_NS, BIB_NS, entries_xml))


# --------------------------------------------------------------------------
# Drop cap frame, nested table, custom document properties
# --------------------------------------------------------------------------

PROPS_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
CUSTOM_PROPS_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_PROPS_REL = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                    "relationships/custom-properties")

_VT_TAG = {str: "lpwstr", bool: "bool", int: "i4", float: "r8"}


def custom_props_xml(properties: dict) -> str:
    """Skeleton for `docProps/custom.xml` (custom document properties).

    `properties` maps name -> value; str/bool/int/float map to the matching
    `vt:` type. pids start at 2 (1 is reserved by the format).
    """
    body = []
    pid = 2
    for name, value in properties.items():
        vt = _VT_TAG[type(value)] if type(value) in _VT_TAG else "lpwstr"
        if isinstance(value, bool):
            text = "true" if value else "false"
        else:
            text = esc(str(value))
        body.append('<property fmtid="%s" pid="%d" name="%s">'
                    "<vt:%s>%s</vt:%s></property>"
                    % (PROPS_FMTID, pid, esc(str(name)), vt, text, vt))
        pid += 1
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            'custom-properties" xmlns:vt="http://schemas.openxmlformats.org/'
            'officeDocument/2006/docPropsVTypes">%s</Properties>' % "".join(body))


def nested_table(rows, widths=None, border_color: str = "808080",
                 header_fill: str = "", header_bold: bool = False) -> str:
    """A simple bordered `w:tbl` for nesting inside a table cell.

    `rows` is a list of row lists (str values rendered as plain runs).
    `widths` is per-column twips; omit for equal auto widths. OOXML requires
    a `w:tc` to end with a paragraph — the op inserts the table before the
    cell's last paragraph, which keeps that invariant.
    """
    ncols = max(len(r) for r in rows)
    if widths is None:
        widths = [2600] * ncols
    border = ('<w:%s w:val="single" w:sz="4" w:space="0" w:color="%s"/>'
              % (side, border_color) for side in
              ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join('<w:gridCol w:w="%d"/>' % w_ for w_ in widths[:ncols])
    trs = []
    for i, row in enumerate(rows):
        fill = ('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % header_fill
                if i == 0 and header_fill else "")
        rpr = "<w:rPr><w:b/></w:rPr>" if i == 0 and header_bold else ""
        tcs = []
        for c in row:
            tcs.append('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>'
                       '<w:p><w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                       "</w:tc>" % (widths[min(len(tcs), ncols - 1, len(widths) - 1)], fill, rpr,
                                    esc(str(c))))
        trs.append("<w:tr>%s</w:tr>" % "".join(tcs))
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>%s'
            "</w:tblBorders></w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>"
            % ("".join(border), grid, "".join(trs)))
