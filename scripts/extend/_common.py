"""Shared plumbing for the docx-master OOXML extension channel.

The `apply` engine normalises structure and authors content. A handful of
Word features sit outside what it writes — section breaks, field-driven
lists (TOC / index), note parts, watermarks, page decoration, drawings,
embedded objects. Those are edited *after* `apply` by mutating the OOXML
package directly, and every one of them needs the same four primitives:

  1. unpack / repack the .docx without losing part order
  2. write a part and register it (content type + relationship)
  3. splice namespace declarations onto a root start tag
  4. insert an element at its schema position

This module is those primitives. `capabilities.py` sits on top and builds
the actual XML; `extend.py` drives both from a JSON instruction file.

Why textual rather than a DOM for (3): `xmlns:` declarations are not
round-trippable through most DOM writers, and Word is strict about a part
declaring every prefix it uses. Rewriting the root start tag after
serialization is the reliable way.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

# --------------------------------------------------------------------------
# Namespaces
# --------------------------------------------------------------------------

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

#: Declarations spliced onto every body / header / footer root. Kept in the
#: same set the engine writes so an extended document and an engine-written
#: one are indistinguishable.
ROOT_DECLARATIONS = [
    ("mc", NS["mc"]),
    ("v", NS["v"]),
    ("o", NS["o"]),
    ("w10", NS["w10"]),
    ("wps", NS["wps"]),
    ("wpg", NS["wpg"]),
    ("wpc", NS["wpc"]),
    ("w14", NS["w14"]),
    ("w15", NS["w15"]),
    ("wp14", NS["wp14"]),
]

#: `mc:Ignorable` lists *extensions* only. wps / wpg / wpc carry visible
#: rendering and must never appear here — the flag means "safe to drop".
IGNORABLE_PREFIXES = ["w14", "w15", "wp14"]
RENDERING_PREFIXES = {"wps", "wpg", "wpc", "v", "o", "w10"}

#: CT_SectPr child sequence (wml.xsd). A sectPr whose children are out of
#: order fails XSD and Word opens the file with a repair prompt.
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
    "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
    "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
]
SECTPR_ORDER = [
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
    "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
]

#: CT_PPr child sequence.
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]

CONTENT_TYPES_PART = "[Content_Types].xml"
RELS_NS = NS["rel"]


# --------------------------------------------------------------------------
# Paths / deletion (sandbox safe-delete shim puts rm into the Recycle Bin,
# which fails outright on some volumes — .NET deletion bypasses it)
# --------------------------------------------------------------------------


def rmtree_force(path) -> None:
    """Delete a directory tree without going through the shell shim."""
    path = str(path)
    if not os.path.exists(path):
        return
    if os.name == "nt":
        subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             '[System.IO.Directory]::Delete("%s", $true)' % path.replace('"', '""')],
            check=False, capture_output=True,
        )
        if os.path.exists(path):
            subprocess.run(["cmd", "/c", "rmdir", "/s", "/q", path], check=False,
                           capture_output=True)
    else:
        shutil.rmtree(path, ignore_errors=True)


# --------------------------------------------------------------------------
# Root-tag declaration splicing
# --------------------------------------------------------------------------

# The attribute group is lazy on purpose: `[^>"']` would otherwise also
# swallow the `/` of a self-closing tag, leaving the `/` inside `attrs` and
# the `(\/?)` group always empty. Lazy matching stops at the first `>` that
# the optional-slash group can satisfy.
_ROOT_TAG_RE = re.compile(r"<(?!\?|!)([\w.\-]+(?::[\w.\-]+)?)((?:\"[^\"]*\"|'[^']*'|[^>\"'])*?)(/?)>")


def ensure_root_declarations(xml: str) -> str:
    """Splice missing namespace declarations onto the root start tag and
    merge `mc:Ignorable`.

    Existing declarations win, so a source document's own bindings survive.
    Prefixes in RENDERING_PREFIXES are stripped from `mc:Ignorable` if some
    other producer put them there.
    """
    match = _ROOT_TAG_RE.search(xml)
    if not match:
        return xml
    name, attrs, self_close = match.group(1), match.group(2), match.group(3)

    for prefix, uri in ROOT_DECLARATIONS:
        if re.search(r"\sxmlns:%s\s*=" % re.escape(prefix), attrs):
            continue
        attrs += ' xmlns:%s="%s"' % (prefix, uri)

    ign = re.search(r"\s*mc:Ignorable\s*=\s*\"([^\"]*)\"", attrs)
    if ign:
        attrs = attrs.replace(ign.group(0), "")
    prefixes = {p for p in (ign.group(1).split() if ign else []) if p}
    prefixes -= RENDERING_PREFIXES
    prefixes |= set(IGNORABLE_PREFIXES)
    attrs += ' mc:Ignorable="%s"' % " ".join(sorted(prefixes))

    return xml[: match.start()] + "<" + name + attrs + self_close + ">" + xml[match.end():]


# --------------------------------------------------------------------------
# Ordered insertion
# --------------------------------------------------------------------------


def insert_ordered(parent, elem, order, tag=None):
    """Insert `elem` into `parent` at its position in `order`.

    Children whose tag isn't in `order` are ignored for positioning, so a
    document carrying elements the list doesn't know about still works.
    """
    name = tag or local(elem.tag)
    if name in order:
        idx = order.index(name)
        for child in list(parent):
            child_name = local(child.tag)
            if child_name in order and order.index(child_name) > idx:
                child.addprevious(elem)
                return elem
    parent.append(elem)
    return elem


def local(tag: str) -> str:
    """`{ns}name` -> `name`."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def W(tag: str) -> str:
    """Qualify a tag name with the wordprocessingml namespace."""
    return "{%s}%s" % (NS["w"], tag)


# --------------------------------------------------------------------------
# Package
# --------------------------------------------------------------------------


class DocxPackage:
    """An open .docx as an editable set of parts.

    Parts are held as text (or bytes for binaries) and re-serialised on
    `save`. Entry order from the source archive is preserved, and new parts
    append — Word is tolerant of order but byte-stable output is easier to
    diff.
    """

    def __init__(self, src):
        self.src = str(src)
        self._parts: "dict[str, bytes]" = {}
        self._order: "list[str]" = []
        self._dirty: "set[str]" = set()
        self._counter = 900

        with zipfile.ZipFile(self.src) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                self._parts[info.filename] = zf.read(info.filename)
                self._order.append(info.filename)

    # -- parts ------------------------------------------------------------

    def has(self, part: str) -> bool:
        return part in self._parts

    def read(self, part: str) -> "str | None":
        data = self._parts.get(part)
        return data.decode("utf-8") if data is not None else None

    def read_bytes(self, part: str) -> "bytes | None":
        return self._parts.get(part)

    def write(self, part: str, data) -> None:
        if isinstance(data, str):
            data = data.encode("utf-8")
        if part not in self._parts:
            self._order.append(part)
        self._parts[part] = data
        self._dirty.add(part)

    def upsert_xml(self, part: str, transform, default: str):
        """Read `part` (or `default` when absent), apply `transform(xml)`,
        and stage the result. Keeps callers free of the read/decode dance."""
        current = self.read(part) or default
        self.write(part, transform(current))
        return self

    def next_id(self, prefix: int = 1) -> int:
        self._counter += prefix
        return self._counter

    # -- registration -----------------------------------------------------

    def add_content_type_override(self, part_name: str, content_type: str) -> None:
        if not part_name.startswith("/"):
            part_name = "/" + part_name
        xml = self.read(CONTENT_TYPES_PART)
        if xml is None:
            raise RuntimeError("package has no [Content_Types].xml")
        if 'PartName="%s"' % part_name in xml:
            return
        xml = xml.replace(
            "</Types>",
            '<Override PartName="%s" ContentType="%s"/></Types>' % (part_name, content_type),
        )
        self.write(CONTENT_TYPES_PART, xml)

    def add_content_type_default(self, extension: str, content_type: str) -> None:
        xml = self.read(CONTENT_TYPES_PART)
        if xml is None:
            return
        if 'Extension="%s"' % extension in xml:
            return
        xml = xml.replace(
            "</Types>",
            '<Default Extension="%s" ContentType="%s"/></Types>' % (extension, content_type),
        )
        self.write(CONTENT_TYPES_PART, xml)

    def add_relationship(self, rels_part: str, rel_type: str, target: str, target_mode=None):
        """Append a relationship, skipping an identical (type, target) pair.

        Returns the relationship id, existing or new.
        """
        xml = self.read(rels_part)
        if xml is None:
            raise RuntimeError("missing relationships part: %s" % rels_part)
        existing = re.search(
            r'<Relationship[^>]*Type="%s"[^>]*Target="%s"[^>]*/>' % (re.escape(rel_type), re.escape(target)),
            xml,
        )
        if existing:
            rid = re.search(r'Id="([^"]+)"', existing.group(0))
            return rid.group(1) if rid else None
        used = [int(m) for m in re.findall(r'Id="rId(\d+)"', xml)]
        rid = "rId%d" % ((max(used) + 1) if used else 1)
        attrs = ' TargetMode="%s"' % target_mode if target_mode else ""
        xml = xml.replace(
            "</Relationships>",
            '<Relationship Id="%s" Type="%s" Target="%s"%s/></Relationships>'
            % (rid, rel_type, target, attrs),
        )
        self.write(rels_part, xml)
        return rid

    def declare_root(self, part: str) -> None:
        xml = self.read(part)
        if xml is None:
            return
        self.write(part, ensure_root_declarations(xml))

    def declare_all_body_parts(self) -> "list[str]":
        touched = []
        for part in list(self._parts):
            if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$", part):
                self.declare_root(part)
                touched.append(part)
        return touched

    # -- lifecycle --------------------------------------------------------

    def save(self, dst) -> str:
        dst = str(dst)
        tmp = dst + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for part in self._order:
                if part in self._parts:
                    zf.writestr(part, self._parts[part])
            # Any part added without going through write() (defensive).
            for part, data in self._parts.items():
                if part not in self._order:
                    zf.writestr(part, data)
        if os.path.exists(dst):
            os.remove(dst)
        os.replace(tmp, dst)
        return dst

    def close(self) -> None:
        """Nothing to release — parts live in memory."""

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False


# --------------------------------------------------------------------------
# Small XML helpers shared by the capability builders
# --------------------------------------------------------------------------


def esc(text: str) -> str:
    """XML-escape text content / attribute values."""
    return (str(text)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def para_text(p) -> str:
    """Concatenated `w:t` content of a paragraph element."""
    return "".join(t.text or "" for t in p.iter(W("t")))
