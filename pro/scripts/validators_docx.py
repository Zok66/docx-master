"""
Validator for Word document XML files against XSD schemas.
"""

import random
import re
import tempfile
import zipfile
from pathlib import Path

import defusedxml.minidom
import lxml.etree

from helpers import safe_extract

from validators_base import BaseSchemaValidator


class DOCXSchemaValidator(BaseSchemaValidator):

    WORD_2006_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"
    W16CID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid"

    ELEMENT_RELATIONSHIP_TYPES = {}

    def validate(self):
        if not self.validate_xml():
            return False

        all_valid = True
        if not self.validate_namespaces():
            all_valid = False

        if not self.validate_unique_ids():
            all_valid = False

        if not self.validate_file_references():
            all_valid = False

        if not self.validate_content_types():
            all_valid = False

        if not self.validate_against_xsd():
            all_valid = False

        if not self.validate_whitespace_preservation():
            all_valid = False

        if not self.validate_deletions():
            all_valid = False

        if not self.validate_insertions():
            all_valid = False

        if not self.validate_all_relationship_ids():
            all_valid = False

        if not self.validate_id_constraints():
            all_valid = False

        if not self.validate_comment_markers():
            all_valid = False

        if not self.validate_field_references():
            all_valid = False

        self.compare_paragraph_counts()

        return all_valid

    def validate_field_references(self):
        """Audit field codes whose *targets* must exist in the package.

        XSD validation cannot catch these — the XML is perfectly valid —
        but Word renders "Error! Undefined style" / "Error! Bookmark not
        defined" at display time, which users report as a broken document.

        Checked here:
          - STYLEREF → resolves against style *display names* (w:name),
            NOT styleIds. A styleId argument silently fails unless it
            happens to equal the display name.
          - REF / PAGEREF / HYPERLINK \\l → bookmark names.
          - DOCVARIABLE → w:docVar/@w:name in word/settings.xml.
        """
        errors = []
        bookmarks = set()
        style_names = set()
        doc_vars = set()
        field_parts = {}  # part name → list of (field code text, line hint)

        for xml_file in self.xml_files:
            rel = xml_file.relative_to(self.unpacked_dir).as_posix()
            try:
                root = lxml.etree.parse(str(xml_file)).getroot()
            except lxml.etree.XMLSyntaxError:
                continue

            W = self.WORD_2006_NAMESPACE
            if rel == "word/styles.xml":
                for name in root.iter(f"{{{W}}}name"):
                    v = name.get(f"{{{W}}}val")
                    if v:
                        style_names.add(v)
            if rel == "word/settings.xml":
                for dv in root.iter(f"{{{W}}}docVar"):
                    v = dv.get(f"{{{W}}}name")
                    if v:
                        doc_vars.add(v)
            for bm in root.iter(f"{{{W}}}bookmarkStart"):
                v = bm.get(f"{{{W}}}name")
                if v:
                    bookmarks.add(v)
            instrs = []
            for it in root.iter(f"{{{W}}}instrText"):
                if it.text:
                    instrs.append((it.text, getattr(it, "sourceline", 0)))
            if instrs:
                field_parts[rel] = instrs

        # Word matches STYLEREF case-insensitively
        style_names_ci = {n.lower() for n in style_names}

        for rel, instrs in sorted(field_parts.items()):
            for text, line in instrs:
                code = text.strip()

                m = re.match(r"STYLEREF\s+(.*)", code)
                if m:
                    arg = m.group(1).strip().strip('"').strip()
                    # numeric form STYLEREF 1..9 → built-in heading level,
                    # always resolvable; skip switch tail like " \n"
                    arg = re.split(r"\\+", arg)[0].strip().strip('"').strip()
                    if arg.isdigit():
                        continue
                    if arg and arg.lower() not in style_names_ci:
                        errors.append(
                            f"  {rel}: Line {line}: STYLEREF \"{arg}\" does not match any style display name (w:name) in word/styles.xml — Word will show \"Error! Undefined style\". STYLEREF resolves by display name, NOT by styleId."
                        )
                    continue

                m = re.match(r"(REF|PAGEREF)\s+(\S+)", code)
                if m:
                    target = m.group(2).strip('"')
                    if target and target not in bookmarks:
                        errors.append(
                            f"  {rel}: Line {line}: {m.group(1)} \"{target}\" has no matching bookmark — Word will show \"Error! Bookmark not defined\"."
                        )
                    continue

                m = re.search(r"HYPERLINK[^<]*?\\l\s*\"([^\"]+)\"", code)
                if m and m.group(1) not in bookmarks:
                    errors.append(
                        f"  {rel}: Line {line}: HYPERLINK \\l \"{m.group(1)}\" has no matching bookmark."
                    )
                    continue

                m = re.match(r"DOCVARIABLE\s+(\S+)", code)
                if m:
                    target = m.group(1).strip('"')
                    if target and target not in doc_vars:
                        errors.append(
                            f"  {rel}: Line {line}: DOCVARIABLE \"{target}\" has no matching w:docVar in word/settings.xml — the field shows an error until a value is set."
                        )

        if errors:
            print(f"FAILED - Found {len(errors)} field reference errors:")
            for error in errors:
                print(error)
            return False
        else:
            if self.verbose:
                print(
                    "PASSED - All field reference targets resolve (STYLEREF → style names, REF/PAGEREF/HYPERLINK → bookmarks, DOCVARIABLE → docVars)"
                )
            return True

    def validate_whitespace_preservation(self):
        errors = []

        for xml_file in self.xml_files:
            if not self.is_main_document(xml_file):
                continue

            try:
                root = lxml.etree.parse(str(xml_file)).getroot()

                for elem in root.iter(f"{{{self.WORD_2006_NAMESPACE}}}t"):
                    if elem.text:
                        text = elem.text
                        if re.search(r"^[ \t\n\r]", text) or re.search(
                            r"[ \t\n\r]$", text
                        ):
                            xml_space_attr = f"{{{self.XML_NAMESPACE}}}space"
                            if (
                                xml_space_attr not in elem.attrib
                                or elem.attrib[xml_space_attr] != "preserve"
                            ):
                                text_preview = (
                                    repr(text)[:50] + "..."
                                    if len(repr(text)) > 50
                                    else repr(text)
                                )
                                errors.append(
                                    f"  {xml_file.relative_to(self.unpacked_dir)}: "
                                    f"Line {elem.sourceline}: w:t element with whitespace missing xml:space='preserve': {text_preview}"
                                )

            except (lxml.etree.XMLSyntaxError, Exception) as e:
                errors.append(
                    f"  {xml_file.relative_to(self.unpacked_dir)}: Error: {e}"
                )

        if errors:
            print(f"FAILED - Found {len(errors)} whitespace preservation violations:")
            for error in errors:
                print(error)
            return False
        else:
            if self.verbose:
                print("PASSED - All whitespace is properly preserved")
            return True

    def validate_deletions(self):
        errors = []

        for xml_file in self.xml_files:
            if not self.is_main_document(xml_file):
                continue

            try:
                root = lxml.etree.parse(str(xml_file)).getroot()
                namespaces = {"w": self.WORD_2006_NAMESPACE}

                for t_elem in root.xpath(".//w:del//w:t", namespaces=namespaces):
                    if t_elem.text:
                        text_preview = (
                            repr(t_elem.text)[:50] + "..."
                            if len(repr(t_elem.text)) > 50
                            else repr(t_elem.text)
                        )
                        errors.append(
                            f"  {xml_file.relative_to(self.unpacked_dir)}: "
                            f"Line {t_elem.sourceline}: <w:t> found within <w:del>: {text_preview}"
                        )

                for instr_elem in root.xpath(
                    ".//w:del//w:instrText", namespaces=namespaces
                ):
                    text_preview = (
                        repr(instr_elem.text or "")[:50] + "..."
                        if len(repr(instr_elem.text or "")) > 50
                        else repr(instr_elem.text or "")
                    )
                    errors.append(
                        f"  {xml_file.relative_to(self.unpacked_dir)}: "
                        f"Line {instr_elem.sourceline}: <w:instrText> found within <w:del> (use <w:delInstrText>): {text_preview}"
                    )

            except (lxml.etree.XMLSyntaxError, Exception) as e:
                errors.append(
                    f"  {xml_file.relative_to(self.unpacked_dir)}: Error: {e}"
                )

        if errors:
            print(f"FAILED - Found {len(errors)} deletion validation violations:")
            for error in errors:
                print(error)
            return False
        else:
            if self.verbose:
                print("PASSED - No w:t elements found within w:del elements")
            return True

    def count_paragraphs_in_unpacked(self):
        count = 0

        for xml_file in self.xml_files:
            # Body part only — see is_main_document() for why a basename
            # match isn't enough (word/glossary/document.xml would win).
            if not self.is_main_document(xml_file):
                continue

            try:
                root = lxml.etree.parse(str(xml_file)).getroot()
                paragraphs = root.findall(f".//{{{self.WORD_2006_NAMESPACE}}}p")
                count = len(paragraphs)
            except Exception as e:
                print(f"Error counting paragraphs in unpacked document: {e}")

        return count

    def count_paragraphs_in_original(self):
        original = self.original_file
        if original is None:
            return 0

        count = 0

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                with zipfile.ZipFile(original, "r") as zip_ref:
                    safe_extract(zip_ref, Path(temp_dir))

                doc_xml_path = temp_dir + "/word/document.xml"
                root = lxml.etree.parse(doc_xml_path).getroot()

                paragraphs = root.findall(f".//{{{self.WORD_2006_NAMESPACE}}}p")
                count = len(paragraphs)

        except Exception as e:
            print(f"Error counting paragraphs in original document: {e}")

        return count

    def validate_insertions(self):
        errors = []

        for xml_file in self.xml_files:
            if not self.is_main_document(xml_file):
                continue

            try:
                root = lxml.etree.parse(str(xml_file)).getroot()
                namespaces = {"w": self.WORD_2006_NAMESPACE}

                invalid_elements = root.xpath(
                    ".//w:ins//w:delText[not(ancestor::w:del)]", namespaces=namespaces
                )

                for elem in invalid_elements:
                    text_preview = (
                        repr(elem.text or "")[:50] + "..."
                        if len(repr(elem.text or "")) > 50
                        else repr(elem.text or "")
                    )
                    errors.append(
                        f"  {xml_file.relative_to(self.unpacked_dir)}: "
                        f"Line {elem.sourceline}: <w:delText> within <w:ins>: {text_preview}"
                    )

            except (lxml.etree.XMLSyntaxError, Exception) as e:
                errors.append(
                    f"  {xml_file.relative_to(self.unpacked_dir)}: Error: {e}"
                )

        if errors:
            print(f"FAILED - Found {len(errors)} insertion validation violations:")
            for error in errors:
                print(error)
            return False
        else:
            if self.verbose:
                print("PASSED - No w:delText elements within w:ins elements")
            return True

    def compare_paragraph_counts(self):
        new_count = self.count_paragraphs_in_unpacked()
        if self.original_file is None:
            print(f"\nParagraphs: {new_count}")
            return

        original_count = self.count_paragraphs_in_original()
        diff = new_count - original_count
        diff_str = f"+{diff}" if diff > 0 else str(diff)
        print(f"\nParagraphs: {original_count} → {new_count} ({diff_str})")

    def _parse_id_value(self, val: str, base: int = 16) -> int:
        return int(val, base)

    def validate_id_constraints(self):
        errors = []
        para_id_attr = f"{{{self.W14_NAMESPACE}}}paraId"
        durable_id_attr = f"{{{self.W16CID_NAMESPACE}}}durableId"

        for xml_file in self.xml_files:
            try:
                for elem in lxml.etree.parse(str(xml_file)).iter():
                    if val := elem.get(para_id_attr):
                        try:
                            if self._parse_id_value(val, base=16) >= 0x80000000:
                                errors.append(
                                    f"  {xml_file.name}:{elem.sourceline}: paraId={val} >= 0x80000000"
                                )
                        except ValueError:
                            errors.append(
                                f"  {xml_file.name}:{elem.sourceline}: "
                                f"paraId={val} is not valid hex"
                            )

                    if val := elem.get(durable_id_attr):
                        if xml_file.name == "numbering.xml":
                            try:
                                if self._parse_id_value(val, base=10) >= 0x7FFFFFFF:
                                    errors.append(
                                        f"  {xml_file.name}:{elem.sourceline}: "
                                        f"durableId={val} >= 0x7FFFFFFF"
                                    )
                            except ValueError:
                                errors.append(
                                    f"  {xml_file.name}:{elem.sourceline}: "
                                    f"durableId={val} must be decimal in numbering.xml"
                                )
                        else:
                            try:
                                if self._parse_id_value(val, base=16) >= 0x7FFFFFFF:
                                    errors.append(
                                        f"  {xml_file.name}:{elem.sourceline}: "
                                        f"durableId={val} >= 0x7FFFFFFF"
                                    )
                            except ValueError:
                                errors.append(
                                    f"  {xml_file.name}:{elem.sourceline}: "
                                    f"durableId={val} is not valid hex"
                                )
            except lxml.etree.XMLSyntaxError:
                continue  

        if errors:
            print(f"FAILED - {len(errors)} ID constraint violations:")
            for e in errors:
                print(e)
        elif self.verbose:
            print("PASSED - All paraId/durableId values within constraints")
        return not errors

    def validate_comment_markers(self):
        errors = []

        document_xml = None
        comments_xml = None
        for xml_file in self.xml_files:
            if self.is_main_document(xml_file):
                document_xml = xml_file
            elif xml_file.name == "comments.xml":
                comments_xml = xml_file

        if not document_xml:
            if self.verbose:
                print("PASSED - No document.xml found (skipping comment validation)")
            return True

        try:
            doc_root = lxml.etree.parse(str(document_xml)).getroot()
            namespaces = {"w": self.WORD_2006_NAMESPACE}

            range_starts = {
                elem.get(f"{{{self.WORD_2006_NAMESPACE}}}id")
                for elem in doc_root.xpath(
                    ".//w:commentRangeStart", namespaces=namespaces
                )
            }
            range_ends = {
                elem.get(f"{{{self.WORD_2006_NAMESPACE}}}id")
                for elem in doc_root.xpath(
                    ".//w:commentRangeEnd", namespaces=namespaces
                )
            }
            references = {
                elem.get(f"{{{self.WORD_2006_NAMESPACE}}}id")
                for elem in doc_root.xpath(
                    ".//w:commentReference", namespaces=namespaces
                )
            }

            orphaned_ends = range_ends - range_starts
            for comment_id in sorted(
                orphaned_ends, key=lambda x: int(x) if x and x.isdigit() else 0
            ):
                errors.append(
                    f'  document.xml: commentRangeEnd id="{comment_id}" has no matching commentRangeStart'
                )

            orphaned_starts = range_starts - range_ends
            for comment_id in sorted(
                orphaned_starts, key=lambda x: int(x) if x and x.isdigit() else 0
            ):
                errors.append(
                    f'  document.xml: commentRangeStart id="{comment_id}" has no matching commentRangeEnd'
                )

            comment_ids = set()
            if comments_xml and comments_xml.exists():
                comments_root = lxml.etree.parse(str(comments_xml)).getroot()
                comment_ids = {
                    elem.get(f"{{{self.WORD_2006_NAMESPACE}}}id")
                    for elem in comments_root.xpath(
                        ".//w:comment", namespaces=namespaces
                    )
                }

                marker_ids = range_starts | range_ends | references
                invalid_refs = marker_ids - comment_ids
                for comment_id in sorted(
                    invalid_refs, key=lambda x: int(x) if x and x.isdigit() else 0
                ):
                    if comment_id:  
                        errors.append(
                            f'  document.xml: marker id="{comment_id}" references non-existent comment'
                        )

        except (lxml.etree.XMLSyntaxError, Exception) as e:
            errors.append(f"  Error parsing XML: {e}")

        if errors:
            print(f"FAILED - {len(errors)} comment marker violations:")
            for error in errors:
                print(error)
            return False
        else:
            if self.verbose:
                print("PASSED - All comment markers properly paired")
            return True

    def repair(self) -> int:
        repairs = super().repair()
        repairs += self.repair_durableId()
        return repairs

    def repair_durableId(self) -> int:
        DURABLE_ID_ATTRS = ("w16cid:durableId", "w16cex:durableId")
        repairs = 0
        renames: dict = {}  

        for xml_file in self.xml_files:
            try:
                content = xml_file.read_text(encoding="utf-8")
                dom = defusedxml.minidom.parseString(content)
                is_numbering = xml_file.name == "numbering.xml"
                base = 10 if is_numbering else 16
                pending = []  
                seen_in_file = set()
                modified = False

                for elem in dom.getElementsByTagName("*"):
                    for attr_name in DURABLE_ID_ATTRS:
                        if not elem.hasAttribute(attr_name):
                            continue

                        durable_id = elem.getAttribute(attr_name)
                        try:
                            key = self._parse_id_value(durable_id, base=base)
                            needs_repair = key >= 0x7FFFFFFF
                        except ValueError:
                            key = durable_id
                            needs_repair = True

                        if needs_repair:
                            if key in seen_in_file:
                                value = random.randint(1, 0x7FFFFFFE)
                            else:
                                seen_in_file.add(key)
                                if key not in renames:
                                    renames[key] = random.randint(1, 0x7FFFFFFE)
                                value = renames[key]
                            new_id = str(value) if is_numbering else f"{value:08X}"

                            elem.setAttribute(attr_name, new_id)
                            pending.append(
                                f"  Repaired: {xml_file.name}: durableId {durable_id} → {new_id}"
                            )
                            modified = True

                if modified:
                    xml_file.write_bytes(dom.toxml(encoding="UTF-8"))
                    for message in pending:
                        print(message)
                    repairs += len(pending)

            except Exception:
                pass

        return repairs


if __name__ == "__main__":
    raise RuntimeError("This module should not be run directly.")
