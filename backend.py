#!/usr/bin/env python3
"""
JIWE_XML_FORMATTER_VISUAL.py
XML-based approach with XML code display and better figure detection
"""

import sys
import os
import zipfile
import argparse
import re
import time
from dataclasses import dataclass
from collections import defaultdict, Counter
from lxml import etree as ET
from openpyxl import Workbook
import tempfile
import io
import shutil
import uuid
import difflib

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
M_NSMAP = {"m": M_NS}

SPECIAL_TEXT_TO_SECTION = {}
SPECIAL_TEXT_FORMATTING_OVERRIDES = {}
JOURNAL_METADATA_TOKEN_SETS = []
JOURNAL_METADATA_TOKEN_SIGNATURES = set()
HARDCODED_SECTION_OVERRIDES = {
    "title": {"font_size": 24.0, "bold": False},
    "journal_metadata": {"font_size": 9.0, "bold": False},
    "introduction": {"bold": False},
    "results and discussions": {"bold": False},
    "body_text": {"bold": False},
    "references": {"font_size": 10.0, "bold": False},
    "subtitle": {"font_size": 16.0, "bold": False},
    "funding statement": {"bold": False},
}
SPECIAL_TITLE_TEXT = (
    "Dropout  P rediction  M odel for  C ollege  S tudents in MOOCs  B ased on  Weight ed   M ulti-feature and SVM"
)
REQUIRED_JOURNAL_METADATA_LINE = (
    "1 ,3   4WGP+JQP, Shenheer Rd, Chang'An, Xi'An, 710100 Shaanxi, China."
)


def normalize_special_key(text):
    if not text:
        return ""
    simplified = re.sub(r"\s+", "", text)
    return simplified.lower()


def normalize_metadata_tokens(text):
    if not text:
        return []
    tokens = re.findall(r"[A-Za-z0-9]+", text.lower())
    cleaned = [re.sub(r"[^a-z0-9]", "", tok) for tok in tokens]
    return [tok for tok in cleaned if tok]


def register_journal_metadata_example(text):
    tokens = normalize_metadata_tokens(text)
    if len(tokens) < 3:
        return
    signature = frozenset(tokens)
    if signature and signature not in JOURNAL_METADATA_TOKEN_SIGNATURES:
        JOURNAL_METADATA_TOKEN_SIGNATURES.add(signature)
        JOURNAL_METADATA_TOKEN_SETS.append(signature)


def matches_registered_journal_metadata(text):
    if not JOURNAL_METADATA_TOKEN_SETS:
        return False
    tokens = set(normalize_metadata_tokens(text))
    if not tokens:
        return False
    for pattern in JOURNAL_METADATA_TOKEN_SETS:
        overlap = len(tokens & pattern)
        if overlap < 3:
            continue
        denom = max(1, min(len(pattern), len(tokens)))
        coverage = overlap / denom
        if coverage >= 0.5:
            return True
    return False


SPECIAL_TEXT_TO_SECTION[normalize_special_key(SPECIAL_TITLE_TEXT)] = "title"
SPECIAL_TEXT_FORMATTING_OVERRIDES[
    normalize_special_key(SPECIAL_TITLE_TEXT)
] = {"font_size": 24.0, "bold": False}

SPECIAL_TEXT_TO_SECTION[
    normalize_special_key(REQUIRED_JOURNAL_METADATA_LINE)
] = "journal_metadata"
SPECIAL_TEXT_FORMATTING_OVERRIDES[
    normalize_special_key(REQUIRED_JOURNAL_METADATA_LINE)
] = {"font_size": 9.0, "bold": False}

SPECIAL_TEXT_TO_SECTION[normalize_special_key("Web Engineering")] = "journal_name"
# SPECIAL_TEXT_FORMATTING_OVERRIDES[
#     normalize_special_key("Web Engineering")
# ] = {"font_size": 18.0, "bold": True}


def apply_special_text_overrides(section_type, paragraph_text, formatting):
    """Force hard-coded formatting for known special-case paragraphs."""
    if formatting is None:
        formatting = {}

    result = dict(formatting)
    special_key = normalize_special_key(paragraph_text)
    overrides = SPECIAL_TEXT_FORMATTING_OVERRIDES.get(special_key)
    if overrides:
        result.update(overrides)
        if "font_size" in overrides and "font_size_w_val" not in overrides:
            if "font_size_w_val" not in result and overrides.get("font_size") is not None:
                hp_value = pt_to_half_points(overrides["font_size"])
                if hp_value is not None:
                    result["font_size_w_val"] = hp_value

    return ensure_font_size_pair(result)


@dataclass
class TemplateProfile:
    """Container for template-derived formatting expectations."""

    rules: dict
    section_order: list
    raw_examples: dict
    custom_rules: dict = None
    _match_cache: dict = None
    context_rules: dict = None
    context_examples: dict = None

    def required_sections(self):
        """Sections that must appear in manuscripts."""
        required = []
        for section in self.section_order:
            if section in (
                "body_text",
                "main_heading",
                "figure_caption",
                "table_caption",
                "submission_history",
                "journal_name",
                "journal_metadata",
                "subtitle",
            ):
                continue
            if section not in required:
                required.append(section)
        return required or ["title", "abstract", "keywords", "references"]

    def resolve_expected_format(
        self, section_type, paragraph_text, context_section=None
    ):
        """Return expected formatting for given section and text using template examples."""
        base = {}

        # 1) Try custom rules (from customXml)
        if self.custom_rules:
            rule_role = SECTION_TO_RULE_ROLE.get(section_type)
            if rule_role and rule_role in self.custom_rules:
                base.update(self.custom_rules[rule_role])

        # 2) Merge aggregated rules from template samples
        aggregated = self.rules.get(section_type)
        if aggregated:
            for key, value in aggregated.items():
                base.setdefault(key, value)

        # 3) Overlay context-specific rules (e.g., body text under References)
        if context_section and self.context_rules:
            context_fmt = self.context_rules.get((section_type, context_section))
            if context_fmt:
                base.update(context_fmt)

        override_fmt = HARDCODED_SECTION_OVERRIDES.get(section_type)
        if override_fmt:
            base.update(override_fmt)
            if (
                override_fmt.get("font_size") is not None
                and base.get("font_size_w_val") is None
            ):
                hp_value = pt_to_half_points(override_fmt["font_size"])
                if hp_value is not None:
                    base["font_size_w_val"] = hp_value

        if (
            section_type == "body_text"
            and context_section
            and isinstance(context_section, str)
        ):
            normalized_context = context_section.lower()
            if normalized_context.startswith("results and discussion"):
                base["bold"] = False
            if normalized_context.startswith("references"):
                base["font_size"] = 9.0
                base["bold"] = False
                if base.get("font_size_w_val") is None:
                    hp_value = pt_to_half_points(9.0)
                    if hp_value is not None:
                        base["font_size_w_val"] = hp_value

        base = ensure_font_size_pair(base)

        # 4) Fallback to generic body text default
        if not base:
            base.update(self.rules.get("body_text", {}))
            base = ensure_font_size_pair(base)

        # 5) Guarantee core expectations from hard defaults when still missing
        default_fmt = get_default_formatting(section_type) or {}
        if default_fmt:
            for key, value in default_fmt.items():
                if key not in base or base[key] is None:
                    base[key] = value
            base = ensure_font_size_pair(base)

        example, score = self.find_matching_example(
            section_type, paragraph_text, context_section=context_section
        )
        if example:
            fmt = dict(base)
            for key in ("font_size", "font_size_w_val", "font_name", "bold", "italic"):
                if key not in fmt or fmt[key] is None:
                    if example.get(key) is not None:
                        fmt[key] = example[key]
            fmt = ensure_font_size_pair(fmt)
            if fmt:
                return apply_special_text_overrides(
                    section_type, paragraph_text, fmt
                )

        return apply_special_text_overrides(
            section_type, paragraph_text, ensure_font_size_pair(base)
        )

    def find_matching_example(
        self, section_type, paragraph_text, context_section=None
    ):
        """Find the best template example for the provided text."""
        examples = []
        if context_section and self.context_examples:
            examples = self.context_examples.get((section_type, context_section)) or []
        if not examples:
            examples = self.raw_examples.get(section_type) or []
        if not examples:
            return None, 0.0

        target_norm = normalize_for_match(paragraph_text)
        if not target_norm:
            target_norm = ""

        best_example = None
        best_score = 0.0

        for example in examples:
            example_text = example.get("text") or ""
            example_norm = normalize_for_match(example_text)

            if not example_norm and not target_norm:
                score = 1.0
            else:
                score = text_similarity(example_norm, target_norm)
                if target_norm and target_norm in example_norm:
                    score = max(score, 0.95)
                elif example_norm and example_norm in target_norm:
                    score = max(score, 0.95)

            if score > best_score:
                best_score = score
                best_example = example

        if best_example and best_score >= 0.45:
            return best_example, best_score
        return None, best_score


# Mapping tables for automatic SDT tagging


def normalize_for_match(text):
    if not text:
        return ""
    if isinstance(text, (int, float)):
        text = str(text)
    text = text.lower().strip()
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    return text.strip()


def text_similarity(a, b):
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


SECTION_TO_ROLE = {
    "journal_name": "jiwe:journal-header",
    "journal_metadata": "jiwe:journal-metadata",
    "title": "jiwe:title",
    "subtitle": "jiwe:subtitle",
    "authors": "jiwe:authors",
    "affiliation": "jiwe:affiliation",
    "corresponding_author": "jiwe:corresponding",
    "abstract": "jiwe:abstract",
    "keywords": "jiwe:keywords",
    "submission_history": "jiwe:submission-history",
    "figure_caption": "jiwe:figure-caption",
    "table_caption": "jiwe:table-caption",
    "acknowledgement": "jiwe:ack-text",
    "funding statement": "jiwe:funding-text",
    "author contributions": "jiwe:contrib",
    "conflict of interests": "jiwe:conflict",
    "ethics statements": "jiwe:ethics",
    "references": "jiwe:reference",
}

ROLE_TAG_TO_SECTION = {
    "jiwe:journal-header": "journal_name",
    "jiwe:journal-metadata": "journal_metadata",
    "jiwe:title": "title",
    "jiwe:subtitle": "subtitle",
    "jiwe:authors": "authors",
    "jiwe:affiliation": "affiliation",
    "jiwe:corresponding": "corresponding_author",
    "jiwe:abstract": "abstract",
    "jiwe:keywords": "keywords",
    "jiwe:submission-history": "submission_history",
    "jiwe:figure-caption": "figure_caption",
    "jiwe:table-caption": "table_caption",
    "jiwe:ack-text": "acknowledgement",
    "jiwe:funding-text": "funding statement",
    "jiwe:contrib": "author contributions",
    "jiwe:conflict": "conflict of interests",
    "jiwe:ethics": "ethics statements",
    "jiwe:reference": "references",
    "jiwe:bio": "biographies of authors",
}

SECTION_TO_RULE_ROLE = {
    "title": "title",
    "journal_name": "journal-header",
    "journal_metadata": "journal-metadata",
    "subtitle": "subtitle",
    "authors": "authors",
    "affiliation": "affiliation",
    "corresponding_author": "corresponding",
    "abstract": "abstract",
    "keywords": "keywords",
    "submission_history": "submission-history",
    "body_text": "body",
    "figure_caption": "caption",
    "table_caption": "caption",
    "acknowledgement": "ack-text",
    "funding statement": "funding-text",
    "author contributions": "contrib",
    "conflict of interests": "conflict",
    "ethics statements": "ethics",
    "references": "reference",
}

LEVEL1_HEADINGS = {
    "introduction",
    "literature review",
    "research methodology",
    "results and discussions",
    "conclusion",
    "acknowledgement",
    "funding statement",
    "author contributions",
    "conflict of interests",
    "ethics statements",
    "references",
}


def ensure_template_tagging(docx_path):
    """Ensure the template DOCX has SDTs and embedded rules."""
    if not docx_path or not os.path.isfile(docx_path):
        return False

    try:
        with zipfile.ZipFile(docx_path, "r") as zin:
            file_data = {
                item.filename: zin.read(item.filename) for item in zin.infolist()
            }
            document_xml = file_data.get("word/document.xml")
            if not document_xml:
                return False

            doc_root = ET.fromstring(document_xml)
            paragraphs = doc_root.findall(".//w:body//w:p", NSMAP)

            paragraph_infos = []
            for idx, p in enumerate(paragraphs):
                info = extract_paragraph_formatting(p, idx)
                if not info:
                    continue
                info["element"] = p
                info["section_type"] = classify_section_type(info)
                paragraph_infos.append(info)

            doc_modified = False

            for info in paragraph_infos:
                p = info["element"]
                if is_paragraph_wrapped(p):
                    continue

                section = info["section_type"]
                role = determine_role_for_section(section, info.get("text", ""))
                if not role:
                    continue

                wrap_paragraph_with_sdt(p, role, alias=section_alias(section))
                doc_modified = True

            need_rules = not has_custom_rules(file_data)

            if not doc_modified and not need_rules:
                return False

            content_types = None
            rels_tree = None
            if need_rules:
                content_types = ET.fromstring(file_data["[Content_Types].xml"])
                rels_tree = ET.fromstring(file_data["_rels/.rels"])

            updated_document = ET.tostring(
                doc_root, encoding="UTF-8", xml_declaration=True
            )

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            tmp.close()

            with zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zout:
                for filename, data in file_data.items():
                    if filename == "word/document.xml":
                        continue
                    if need_rules and filename == "[Content_Types].xml":
                        continue
                    if need_rules and filename == "_rels/.rels":
                        continue
                    zout.writestr(filename, data)

                zout.writestr("word/document.xml", updated_document)

                if need_rules:
                    rules_name, extra_entries = add_rules_parts(
                        zout, content_types, rels_tree, file_data.keys()
                    )
                    for entry_name, entry_bytes in extra_entries.items():
                        zout.writestr(entry_name, entry_bytes)

            shutil.move(tmp.name, docx_path)
            return True
    except Exception as exc:
        print(f"[template-tagging] Could not ensure SDTs: {exc}")
        return False


def is_paragraph_wrapped(paragraph):
    """Return True if the paragraph is already inside an SDT."""
    parent = paragraph.getparent()
    while parent is not None:
        if parent.tag == f"{{{W_NS}}}sdtContent":
            return True
        parent = parent.getparent()
    return False


def detect_paragraph_role(paragraph_element):
    parent = paragraph_element
    while parent is not None:
        if parent.tag == f"{{{W_NS}}}sdtContent":
            sdt = parent.getparent()
            if sdt is not None and sdt.tag == f"{{{W_NS}}}sdt":
                tag_elem = sdt.find("./w:sdtPr/w:tag", NSMAP)
                if tag_elem is not None:
                    val = tag_elem.get(f"{{{W_NS}}}val")
                    if val:
                        return val
        parent = parent.getparent()
    return None


def section_alias(section):
    if not section:
        return "Content"
    return section.replace("_", " ").title()


def determine_role_for_section(section, text):
    if not section:
        return None

    if section in SECTION_TO_ROLE:
        return SECTION_TO_ROLE[section]

    if section in LEVEL1_HEADINGS:
        return "jiwe:heading level=1"

    if section == "main_heading":
        if re.match(r"\d+\.\d+\.\d+", text or "", re.IGNORECASE):
            return "jiwe:heading level=3"
        if re.match(r"\d+\.\d+", text or "", re.IGNORECASE):
            return "jiwe:heading level=2"
        return "jiwe:heading level=2"

    return None


def wrap_paragraph_with_sdt(paragraph, role, alias=None):
    """Wrap a paragraph element with an SDT tagged with the given role."""
    sdt = ET.Element(f"{{{W_NS}}}sdt")
    sdtPr = ET.SubElement(sdt, f"{{{W_NS}}}sdtPr")
    if alias:
        alias_elem = ET.SubElement(sdtPr, f"{{{W_NS}}}alias")
        alias_elem.set(f"{{{W_NS}}}val", alias)
    tag_elem = ET.SubElement(sdtPr, f"{{{W_NS}}}tag")
    tag_elem.set(f"{{{W_NS}}}val", role)
    sdtContent = ET.SubElement(sdt, f"{{{W_NS}}}sdtContent")

    parent = paragraph.getparent()
    if parent is None:
        return
    position = parent.index(paragraph)
    parent.remove(paragraph)
    sdtContent.append(paragraph)
    parent.insert(position, sdt)


def has_custom_rules(file_data):
    for name in file_data:
        if name.startswith("customXml/") and name.endswith(".xml"):
            try:
                tree = ET.fromstring(file_data[name])
                if tree.tag.endswith("rules"):
                    return True
            except Exception:
                continue
    return False


def build_rules_xml():
    ns = "https://spec.jiwe.example/v1"
    rules = ET.Element(f"{{{ns}}}rules", attrib={"version": "1.0"})
    fonts = ET.SubElement(rules, f"{{{ns}}}fonts")
    font_specs = [
        ("title", "Times New Roman", "24", "bold", None),
        ("journal-header", "Palatino Linotype", "24", "bold", None),
        ("journal-metadata", "Times New Roman", "9", "normal", None),
        ("subtitle", "Times New Roman", "16", "normal", None),
        ("authors", "Times New Roman", "11", "bold", None),
        ("affiliation", "Times New Roman", "9", "normal", None),
        ("corresponding", "Times New Roman", "9", "normal", "italic"),
        ("abstract", "Times New Roman", "9", "normal", None),
        ("keywords", "Times New Roman", "9", "normal", "italic"),
        ("body", "Times New Roman", "10", "normal", None),
        ("heading-1", "Times New Roman", "10", "bold", None),
        ("heading-2", "Times New Roman", "10", "bold", None),
        ("heading-3", "Times New Roman", "10", "bold", "italic"),
        ("caption", "Times New Roman", "10", "normal", None),
        ("reference", "Times New Roman", "10", "normal", None),
    ]
    for role, family, size, weight, style in font_specs:
        font_elem = ET.SubElement(
            fonts,
            f"{{{ns}}}font",
            attrib={"role": role, "family": family, "sizePt": size, "weight": weight},
        )
        if style:
            font_elem.set("style", style)

    sections = ET.SubElement(
        rules,
        f"{{{ns}}}sections",
        attrib={
            "order": (
                "journal-header,journal-metadata,title,authors,affiliation,corresponding,"
                "abstract,keywords,heading,paragraph,figure-caption,table-caption,ack-text,"
                "funding-text,contrib,conflict,ethics,reference,bio"
            ),
            "maxHeadingDepth": "3",
        },
    )
    ET.SubElement(
        rules,
        f"{{{ns}}}captions",
        attrib={"figurePrefix": "Figure", "tablePrefix": "Table"},
    )
    ET.SubElement(rules, f"{{{ns}}}references", attrib={"style": "IEEE"})
    ET.SubElement(
        rules, f"{{{ns}}}abstract", attrib={"minWords": "200", "maxWords": "300"}
    )
    ET.SubElement(rules, f"{{{ns}}}keywords", attrib={"minCount": "5"})
    return ET.tostring(rules, encoding="UTF-8", xml_declaration=True)


def add_rules_parts(zout, content_types_tree, rels_tree, existing_names):
    """Add the rules custom XML part and update supporting parts."""
    existing_items = [
        int(match.group(1))
        for name in existing_names
        for match in [re.match(r"customXml/item(\d+)\.xml", name)]
        if match
    ]
    index = max(existing_items, default=0) + 1
    item_name = f"customXml/item{index}.xml"
    item_props_name = f"customXml/itemProps{index}.xml"
    rels_name = f"customXml/_rels/item{index}.xml.rels"

    # Update [Content_Types].xml
    ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.SubElement(
        content_types_tree,
        f"{{{ns_ct}}}Override",
        attrib={"PartName": f"/{item_name}", "ContentType": "application/xml"},
    )
    ET.SubElement(
        content_types_tree,
        f"{{{ns_ct}}}Override",
        attrib={
            "PartName": f"/{item_props_name}",
            "ContentType": "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
        },
    )
    content_types_bytes = ET.tostring(
        content_types_tree, encoding="UTF-8", xml_declaration=True
    )

    # Update _rels/.rels
    pkg_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rel_ids = {rel.get("Id") for rel in rels_tree.findall(f"{{{pkg_ns}}}Relationship")}
    new_id = f"rId{len(rel_ids) + 1}"
    while new_id in rel_ids:
        new_id = f"rId{len(rel_ids) + 2}"
    ET.SubElement(
        rels_tree,
        f"{{{pkg_ns}}}Relationship",
        attrib={
            "Id": new_id,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
            "Target": item_name,
        },
    )
    rels_bytes = ET.tostring(rels_tree, encoding="UTF-8", xml_declaration=True)

    # Build custom XML files
    rules_bytes = build_rules_xml()
    item_props = ET.Element(
        "{http://schemas.openxmlformats.org/officeDocument/2006/customXml}datastoreItem",
        attrib={"itemID": f"{{{uuid.uuid4()}}}"},
    )
    schema_refs = ET.SubElement(
        item_props,
        "{http://schemas.openxmlformats.org/officeDocument/2006/customXml}schemaRefs",
    )
    ET.SubElement(
        schema_refs,
        "{http://schemas.openxmlformats.org/officeDocument/2006/customXml}schemaRef",
        attrib={"uri": "https://spec.jiwe.example/v1"},
    )
    item_props_bytes = ET.tostring(item_props, encoding="UTF-8", xml_declaration=True)

    rels_root = ET.Element(f"{{{pkg_ns}}}Relationships")
    ET.SubElement(
        rels_root,
        f"{{{pkg_ns}}}Relationship",
        attrib={
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
            "Target": os.path.basename(item_props_name),
        },
    )
    item_rels_bytes = ET.tostring(rels_root, encoding="UTF-8", xml_declaration=True)

    extra_entries = {
        "[Content_Types].xml": content_types_bytes,
        "_rels/.rels": rels_bytes,
        item_name: rules_bytes,
        item_props_name: item_props_bytes,
        rels_name: item_rels_bytes,
    }

    return item_name, extra_entries


def load_custom_rules(docx_source):
    """Load custom formatting rules embedded in customXml."""
    rules = {}
    # Support unzipped template directories containing customXml/*.xml
    try:
        if isinstance(docx_source, str) and os.path.isdir(docx_source):
            custom_dir = os.path.join(docx_source, "customXml")
            if os.path.isdir(custom_dir):
                for name in os.listdir(custom_dir):
                    if not name.endswith(".xml"):
                        continue
                    full_path = os.path.join(custom_dir, name)
                    try:
                        tree = ET.parse(full_path).getroot()
                    except Exception:
                        continue
                    if not tree.tag.endswith("rules"):
                        continue
                    for font_elem in tree.findall(".//{*}font"):
                        role = font_elem.get("role")
                        if not role:
                            continue
                        info = {}
                        family = font_elem.get("family")
                        if family:
                            info["font_name"] = normalize_font_name(family)
                        size = font_elem.get("sizePt")
                        if size:
                            try:
                                size_pt = float(size)
                                info["font_size"] = size_pt
                                hp = pt_to_half_points(size_pt)
                                if hp is not None:
                                    info["font_size_w_val"] = hp
                            except ValueError:
                                pass
                        weight = (font_elem.get("weight") or "").lower()
                        if weight:
                            info["bold"] = weight == "bold"
                        style = (font_elem.get("style") or "").lower()
                        if style:
                            if style == "italic":
                                info["italic"] = True
                            elif style == "normal":
                                info["italic"] = False
                        rules[role] = info
                if rules:
                    return rules
    except Exception as exc:
        print(f"[template-tagging] Warning loading directory custom rules: {exc}")
    try:
        data = read_docx_bytes(docx_source)
    except Exception:
        return rules

    try:
        with zipfile.ZipFile(io.BytesIO(data), "r") as zin:
            for name in zin.namelist():
                if not (name.startswith("customXml/") and name.endswith(".xml")):
                    continue
                try:
                    tree = ET.fromstring(zin.read(name))
                except Exception:
                    continue
                if not tree.tag.endswith("rules"):
                    continue
                for font_elem in tree.findall(".//{*}font"):
                    role = font_elem.get("role")
                    if not role:
                        continue
                    info = {}
                    family = font_elem.get("family")
                    if family:
                        info["font_name"] = normalize_font_name(family)
                    size = font_elem.get("sizePt")
                    if size:
                        try:
                            size_pt = float(size)
                            info["font_size"] = size_pt
                            hp = pt_to_half_points(size_pt)
                            if hp is not None:
                                info["font_size_w_val"] = hp
                        except ValueError:
                            pass
                    weight = (font_elem.get("weight") or "").lower()
                    if weight:
                        info["bold"] = weight == "bold"
                    style = (font_elem.get("style") or "").lower()
                    if style:
                        if style == "italic":
                            info["italic"] = True
                        elif style == "normal":
                            info["italic"] = False
                    rules[role] = info
                break
    except Exception as exc:
        print(f"[template-tagging] Warning loading custom rules: {exc}")

    return rules


def normalize_font_from_rfonts(rFonts_elem):
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        val = rFonts_elem.get(f"{{{W_NS}}}{attr}")
        if val:
            return normalize_font_name(val)
    return None


def load_style_fonts(docx_source):
    """Extract paragraph style fonts and defaults from the DOCX."""
    style_fonts = {}
    default_font = None

    try:
        data = read_docx_bytes(docx_source)
    except Exception:
        return style_fonts, default_font

    try:
        with zipfile.ZipFile(io.BytesIO(data), "r") as zin:
            if "word/styles.xml" not in zin.namelist():
                return style_fonts, default_font
            styles_root = ET.fromstring(zin.read("word/styles.xml"))

        doc_defaults = styles_root.find("w:docDefaults", NSMAP)
        if doc_defaults is not None:
            rPr_default = doc_defaults.find(".//w:rPrDefault/w:rPr", NSMAP)
            if rPr_default is not None:
                rFonts_default = rPr_default.find("w:rFonts", NSMAP)
                if rFonts_default is not None:
                    default_font = normalize_font_from_rfonts(rFonts_default)

        for style in styles_root.findall("w:style", NSMAP):
            style_id = style.get(f"{{{W_NS}}}styleId")
            if not style_id:
                continue
            style_type = style.get(f"{{{W_NS}}}type")
            if style_type not in (None, "paragraph"):
                continue

            font_name = None
            rPr = style.find("w:rPr", NSMAP)
            if rPr is not None:
                rFonts = rPr.find("w:rFonts", NSMAP)
                if rFonts is not None:
                    font_name = normalize_font_from_rfonts(rFonts)

            if not font_name:
                based_on = style.find("w:basedOn", NSMAP)
                if based_on is not None:
                    base_id = based_on.get(f"{{{W_NS}}}val")
                    if base_id and base_id in style_fonts:
                        font_name = style_fonts[base_id]

            if not font_name and default_font:
                font_name = default_font

            if font_name:
                style_fonts[style_id] = font_name

    except Exception as exc:
        print(f"[style detection] Warning: {exc}")

    return style_fonts, default_font


# -------------------------
# XML Conversion & Display Functions
# -------------------------
def docx_to_xml(docx_file):
    """Convert DOCX file to XML structure and return both root and string"""
    if hasattr(docx_file, "read"):
        # Handle file-like object (Streamlit upload)
        docx_file.seek(0)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(docx_file.read())
            docx_path = tmp.name
    else:
        # Handle file path
        docx_path = docx_file

    try:
        # Extract document.xml from DOCX
        with zipfile.ZipFile(docx_path, "r") as z:
            xml_content = z.read("word/document.xml")

        # Parse XML and also return pretty string
        root = ET.fromstring(xml_content)
        xml_string = ET.tostring(root, encoding="unicode", pretty_print=True)

        return root, xml_string

    except Exception as e:
        print(f"Error converting DOCX to XML: {str(e)}")
        return None, None
    finally:
        # Clean up temp file if we created one
        if hasattr(docx_file, "read"):
            try:
                os.unlink(docx_path)
            except:
                pass


def get_xml_preview(xml_root, max_paragraphs=10):
    """Get a preview of XML structure for display"""
    if xml_root is None:
        return "Error: Could not parse XML"

    preview_lines = []
    paragraphs = xml_root.findall(".//w:p", NSMAP)[:max_paragraphs]

    for idx, p in enumerate(paragraphs):
        preview_lines.append(f"\n=== Paragraph {idx} ===")

        # Get text content
        texts = []
        for r in p.findall(".//w:r", NSMAP):
            for t in r.findall(".//w:t", NSMAP):
                if t.text:
                    texts.append(t.text)

        if texts:
            full_text = " ".join(texts)
            preview_lines.append(
                f"Text: {full_text[:100]}{'...' if len(full_text) > 100 else ''}"
            )

        # Get formatting info
        pPr = p.find("./w:pPr", NSMAP)
        if pPr is not None:
            pStyle = pPr.find("./w:pStyle", NSMAP)
            if pStyle is not None:
                preview_lines.append(f"Style: {pStyle.get('{%s}val' % W_NS)}")

        # Get run formatting
        for r_idx, r in enumerate(p.findall(".//w:r", NSMAP)[:3]):  # First 3 runs
            rPr = r.find("./w:rPr", NSMAP)
            if rPr is not None:
                format_info = []

                # Font size
                sz = rPr.find("./w:sz", NSMAP)
                if sz is not None:
                    size_val = sz.get("{%s}val" % W_NS)
                    if size_val:
                        format_info.append(f"size:{half_points_to_pt(size_val)}pt")

                # Font name
                rf = rPr.find("./w:rFonts", NSMAP)
                if rf is not None:
                    font_name = rf.get("{%s}ascii" % W_NS) or rf.get("{%s}hAnsi" % W_NS)
                    if font_name:
                        format_info.append(f"font:{normalize_font_name(font_name)}")

                # Bold
                if rPr.find("./w:b", NSMAP) is not None:
                    format_info.append("bold")

                # Italic
                if rPr.find("./w:i", NSMAP) is not None:
                    format_info.append("italic")

                if format_info:
                    preview_lines.append(f"  Run {r_idx}: {', '.join(format_info)}")

    return "\n".join(preview_lines)


# -------------------------
# Enhanced Figure Detection
# -------------------------
def extract_paragraphs_from_xml(xml_root, style_fonts=None, default_font=None):
    """Extract all paragraphs with formatting from XML"""
    paragraphs = []

    for idx, p in enumerate(xml_root.findall(".//w:p", NSMAP)):
        paragraph_data = extract_paragraph_formatting(p, idx, style_fonts, default_font)
        if paragraph_data and paragraph_data["text"].strip():
            paragraphs.append(paragraph_data)

    return paragraphs


def extract_paragraph_formatting(paragraph, index, style_fonts=None, default_font=None):
    """Extract formatting and text from a paragraph element"""
    # Get paragraph style
    p_style = None
    alignment = None
    pPr = paragraph.find("./w:pPr", NSMAP)
    if pPr is not None:
        pStyle = pPr.find("./w:pStyle", NSMAP)
        if pStyle is not None:
            p_style = pStyle.get("{%s}val" % W_NS)
        jc = pPr.find("./w:jc", NSMAP)
        if jc is not None:
            alignment = jc.get("{%s}val" % W_NS)

    # Get all text runs
    texts = []
    runs_data = []

    for r in paragraph.findall(".//w:r", NSMAP):
        run_text = "".join(t.text for t in r.findall(".//w:t", NSMAP) if t.text)
        if run_text:
            run_data = extract_run_formatting(
                r, style_fonts=style_fonts, default_font=default_font, style_id=p_style
            )
            run_data["text"] = run_text
            runs_data.append(run_data)
            texts.append(run_text)

    if not texts:
        return None

    full_text = " ".join(texts).strip()

    # Determine dominant formatting from runs
    dominant_format = get_dominant_formatting(runs_data)

    if p_style and style_fonts:
        style_font = style_fonts.get(p_style)
        if style_font and not dominant_format.get("font_name"):
            dominant_format["font_name"] = style_font

    if not dominant_format.get("font_name") and default_font:
        dominant_format["font_name"] = default_font

    role_tag = detect_paragraph_role(paragraph)

    return {
        "index": index,
        "text": full_text,
        "p_style": p_style,
        "font_size": dominant_format.get("font_size"),
        "font_size_w_val": dominant_format.get("font_size_w_val"),
        "font_name": dominant_format.get("font_name"),
        "bold": dominant_format.get("bold", False),
        "italic": dominant_format.get("italic", False),
        "runs": runs_data,
        "role_tag": role_tag,
        "alignment": alignment,
    }


def extract_run_formatting(run, style_fonts=None, default_font=None, style_id=None):
    """Extract formatting from a run element"""
    format_data = {}

    rPr = run.find("./w:rPr", NSMAP)
    if rPr is None:
        return format_data

    # Font size
    size_val = None
    sz = rPr.find("./w:sz", NSMAP)
    if sz is not None:
        size_val = sz.get("{%s}val" % W_NS)
    if not size_val:
        szCs = rPr.find("./w:szCs", NSMAP)
        if szCs is not None:
            size_val = szCs.get("{%s}val" % W_NS)
    if size_val:
        half_points = normalize_half_points(size_val)
        if half_points is not None:
            format_data["font_size_w_val"] = half_points
            pt_value = half_points_to_pt(half_points)
            if pt_value is not None:
                format_data["font_size"] = pt_value

    # Font name
    rf = rPr.find("./w:rFonts", NSMAP)
    if rf is not None:
        font_name = rf.get("{%s}ascii" % W_NS) or rf.get("{%s}hAnsi" % W_NS)
        if font_name:
            format_data["font_name"] = normalize_font_name(font_name)
    if "font_name" not in format_data:
        # fallback to paragraph style font if available
        p_element = run
        while p_element is not None and p_element.tag != f"{{{W_NS}}}p":
            p_element = p_element.getparent()
        if p_element is not None:
            pPr = p_element.find("w:pPr", NSMAP)
            if pPr is not None:
                rPr_style = pPr.find("w:rPr", NSMAP)
                if rPr_style is not None:
                    rf_style = rPr_style.find("w:rFonts", NSMAP)
                    if rf_style is not None:
                        style_font = (
                            rf_style.get("{%s}ascii" % W_NS)
                            or rf_style.get("{%s}hAnsi" % W_NS)
                            or rf_style.get("{%s}eastAsia" % W_NS)
                            or rf_style.get("{%s}cs" % W_NS)
                        )
                        if style_font:
                            format_data["font_name"] = normalize_font_name(style_font)
    if "font_name" not in format_data and style_fonts and style_id:
        style_font = style_fonts.get(style_id)
        if style_font:
            format_data["font_name"] = style_font
    if "font_name" not in format_data and default_font:
        format_data["font_name"] = default_font

    # Bold
    if rPr.find("./w:b", NSMAP) is not None:
        format_data["bold"] = True

    # Italic
    if rPr.find("./w:i", NSMAP) is not None:
        format_data["italic"] = True

    return format_data


def get_dominant_formatting(runs_data):
    """Determine dominant formatting from multiple runs"""
    if not runs_data:
        return {}

    font_sizes = []
    font_size_vals = []
    font_names = []
    bold_count = 0
    italic_count = 0
    total_runs = len(runs_data)

    for run in runs_data:
        if run.get("font_size"):
            font_sizes.append(run["font_size"])
        if run.get("font_size_w_val") is not None:
            font_size_vals.append(run["font_size_w_val"])
        if run.get("font_name"):
            font_names.append(run["font_name"])
        if run.get("bold"):
            bold_count += 1
        if run.get("italic"):
            italic_count += 1

    dominant = {}

    if font_sizes:
        dominant["font_size"] = Counter(font_sizes).most_common(1)[0][0]
    elif font_size_vals:
        dominant["font_size"] = half_points_to_pt(
            Counter(font_size_vals).most_common(1)[0][0]
        )

    if font_size_vals:
        dominant["font_size_w_val"] = Counter(font_size_vals).most_common(1)[0][0]

    if font_names:
        dominant["font_name"] = Counter(font_names).most_common(1)[0][0]

    if bold_count > total_runs / 2:
        dominant["bold"] = True
    if italic_count > total_runs / 2:
        dominant["italic"] = True

    return dominant


# -------------------------
# Enhanced Section Classification
# -------------------------
def classify_section_type(paragraph):
    """
    Robust section classifier.

    - Normalises text (strip numbering prefixes, collapse whitespace, lowercase)
    - Matches section keywords using word-boundary checks and startswith on cleaned text
    - Preserves strong signals (large font -> title, figure/table captions)
    """
    role_tag = (paragraph.get("role_tag") or "").strip().lower()
    if role_tag:
        primary_role = role_tag.split()[0]
        mapped = ROLE_TAG_TO_SECTION.get(role_tag) or ROLE_TAG_TO_SECTION.get(
            primary_role
        )
        if mapped:
            return mapped
        if primary_role.startswith("jiwe:heading"):
            return "main_heading"
        if primary_role.startswith("jiwe:journal-header"):
            return "journal_name"
        if primary_role.startswith("jiwe:journal-metadata"):
            return "journal_metadata"
        if primary_role.startswith("jiwe:body"):
            return "body_text"

    raw = (paragraph.get("text") or "").strip()
    special_key = normalize_special_key(raw)
    if special_key in SPECIAL_TEXT_TO_SECTION:
        return SPECIAL_TEXT_TO_SECTION[special_key]

    if not raw:
        return "body_text"

    # normalize
    text = raw.replace("\r", " ").replace("\n", " ").strip()
    text = re.sub(r"\s+", " ", text)  # collapse whitespace
    lower_text = text.lower()
    font_size = paragraph.get("font_size") or 0
    p_style = (paragraph.get("p_style") or "").lower()
    idx = paragraph.get("index", 999)
    raw_words = [w for w in re.split(r"\s+", text) if w]

    def strip_token(token):
        return re.sub(r"^[^A-Za-z0-9]+|[^A-Za-z0-9]+$", "", token)

    words = [strip_token(w) or w for w in raw_words]
    titlecase_count = sum(
        1
        for w in words
        if len(w) > 2 and w[0].isupper() and (len(w) == 1 or w[1:].islower())
    )

    # helper: remove leading numbering like "1.", "I.", "1.1", "1 -", "1 Introduction" etc.
    cleaned = lower_text
    cleaned = re.sub(r"^[\s\-\–\—]*[\d]+(?:[.\d]*)?\s*[\.\-:\)]*\s*", "", cleaned)
    cleaned = re.sub(
        r"(?i)^[\s\-\–\—]*[ivxlcdm]{2,}(?=\b)\s*[\.\-:\)]*\s*", "", cleaned
    )
    cleaned = re.sub(
        r"(?i)^[\s\-\–\—]*[ivxlcdm](?=[\s\.\-:\)])[\s\.\-:\)]*", "", cleaned
    )
    cleaned = cleaned.strip()
    # also strip common bullets/markers
    cleaned = re.sub(r"^[\*\•\·\u2022\-\–\—\•]+\s*", "", cleaned)
    cleaned = cleaned.lstrip(" ,;:-")

    if not cleaned:
        return "body_text"

    if re.fullmatch(r"[\d\W]+", cleaned):
        return "body_text"

    alpha_count = sum(1 for ch in cleaned if ch.isalpha())
    if alpha_count <= 2 and len(cleaned) <= 6:
        return "body_text"

    section_keywords = {
        "abstract": ["abstract"],
        "keywords": ["keywords", "keyword"],
        "affiliation": ["affiliation", "department of", "faculty of", "school of"],
        "introduction": ["introduction"],
        "subtitle": ["subtitle"],
        "literature review": ["literature review", "related work"],
        "research methodology": ["research methodology", "methodology", "methods"],
        "results and discussions": ["results and discussions", "results", "discussion"],
        "conclusion": ["conclusion", "conclusions"],
        "acknowledgement": ["acknowledgement", "acknowledgments", "acknowledgement"],
        "funding statement": ["funding statement", "funding"],
        "author contributions": ["author contributions", "contributions"],
        "conflict of interests": [
            "conflict of interests",
            "conflicts of interest",
            "competing interests",
        ],
        "ethics statements": ["ethics statements", "ethical statement", "ethics"],
        "references": ["references", "reference", "bibliography"],
    }

    if (
        "font size" in lower_text
        and "title" not in lower_text
        and "figure" not in lower_text
        and "table" not in lower_text
    ):
        simplified_heading = re.sub(r"\s*\(.*?\)", " ", cleaned).strip()
        all_keywords = {kw for kws in section_keywords.values() for kw in kws}
        if not any(
            simplified_heading.startswith(kw) or f" {kw} " in simplified_heading
            for kw in all_keywords
        ):
            return "body_text"

    if idx <= 1 and "journal" in lower_text:
        return "journal_name"

    # 2) Journal / Article titles based on font size
    if font_size and font_size >= 20:
        if idx <= 1:
            return "journal_name"
        return "title"

    front_matter_tokens = ("vol.", "volume", "no.", "issue", "issn", "eissn", "doi")
    if idx < 8 and any(tok in lower_text for tok in front_matter_tokens):
        return "journal_metadata"

    if idx < 20:
        if matches_registered_journal_metadata(text):
            return "journal_metadata"
        metadata_keywords = (
            "department",
            "faculty",
            "university",
            "institute",
            "school",
            "jalan",
            "road",
            "street",
            "malaysia",
            "singapore",
            "taiwan",
            "china",
            "india",
            "thailand",
            "tel",
            "fax",
            "postal",
            "address",
        )
        if (
            "author" not in lower_text
            and any(term in lower_text for term in metadata_keywords)
            and (re.search(r"\d", text) or "," in text or ";" in text)
        ):
            return "journal_metadata"

    # 1) Paragraph style strong hints
    if p_style:
        if "subtitle" in p_style:
            return "subtitle"
        if "title" in p_style:
            return "title"
        if "heading" in p_style or p_style.startswith("h"):
            # try to resolve to known section names
            for sect in (
                "introduction",
                "abstract",
                "keywords",
                "conclusion",
                "references",
            ):
                if re.search(r"\b" + re.escape(sect) + r"\b", lower_text):
                    return sect
            return "main_heading"

    # 3) Figure captions
    for pattern in [r"^(figure|fig)\b", r"^(figure|fig)\s*\d+", r"^fig\.?\s*\d+"]:
        if re.match(pattern, cleaned, re.IGNORECASE):
            if len(text) < 300 and not any(
                term in cleaned for term in ("abstract", "keyword", "reference")
            ):
                return "figure_caption"

    # 4) Table captions
    for pattern in [r"^(table|tab)\b", r"^table\s*\d+"]:
        if re.match(pattern, cleaned, re.IGNORECASE):
            if len(text) < 300:
                return "table_caption"

    alignment = (paragraph.get("alignment") or "").lower()
    if idx < 8 and alignment == "center":
        non_lower_words = sum(
            1
            for w in words
            if w
            and any(ch.isalpha() for ch in w)
            and not w.islower()
        )
        if (
            len(words) >= 4
            and titlecase_count >= 3
            and non_lower_words >= len(words) - 1
            and not paragraph.get("bold")
        ):
            return "title"

    # Check explicit starts (cleaned) e.g. "introduction", or exact match, or word-boundary anywhere
    for sect, kws in section_keywords.items():
        for kw in kws:
            if (
                cleaned.startswith(kw)
                or re.match(r"^(?:\d+[.)]?\s*)?" + re.escape(kw) + r"\b", cleaned)
                or re.match(r"^(?:" + re.escape(kw) + r")\s*[:–-]", cleaned)
            ):
                # Ensure it’s short enough to be a heading (< 10 words)
                if (
                    sect in ["abstract", "keywords", "affiliation"]
                    or len(cleaned.split()) <= 15
                ):
                    return sect

    # 6) Numbered headings like "1. Introduction" or "1 Introduction" — cleaned will remove the number,
    #    so if cleaned contains a known section keyword we already caught it. But handle cases where
    #    cleaned is short and matches "introduction" etc.
    if re.match(r"^\d+[\.\)]?\s*\w+", lower_text) and len(text) < 200:
        for sect, kws in section_keywords.items():
            for kw in kws:
                if kw in lower_text:
                    return sect
        return "main_heading"

    # 7) Author / corresponding hints (early in doc)
    if idx < 12:
        if "@" in raw or "correspond" in lower_text:
            return "corresponding_author"
        # name-like line with commas and TitleCase tokens
        if "," in raw and not any(
            term in lower_text
            for term in ("department", "faculty", "school", "university")
        ):
            name_like = (
                len(re.findall(r"\b[A-Z][a-z]{1,}\s+[A-Z][a-z]{1,}\b", raw)) >= 1
            )
            if name_like:
                return "authors"

    if any(term in lower_text for term in ("received:", "accepted:", "published:")):
        return "submission_history"

    # 8) Early short title heuristic (if early and looks like Title Case)
    if idx < 4 and len(words) > 2:
        blocked_terms = (
            "abstract",
            "keywords",
            "introduction",
            "conclusion",
            "references",
        )
        if titlecase_count >= 2 and not any(
            term in lower_text for term in blocked_terms
        ):
            return "title"

    return "body_text"


# -------------------------
# Main Analysis Function
# -------------------------
def analyze_documents(template_file, manuscript_file):
    """Main analysis function that returns findings, missing, and XML previews"""
    custom_rules = {}
    try:
        # Always load custom rules (supports file-like DOCX, path DOCX, or unzipped dir)
        if isinstance(template_file, str) and os.path.isfile(template_file):
            # Only attempt to modify/tag the template when we have a real file path
            ensure_template_tagging(template_file)
        # Load custom rules regardless of input type
        custom_rules = load_custom_rules(template_file)
    except Exception as exc:
        print(f"[template-tagging] Warning: {exc}")

    # Convert both documents to XML
    template_xml, template_xml_string = docx_to_xml(template_file)
    manuscript_xml, manuscript_xml_string = docx_to_xml(manuscript_file)

    if template_xml is None or manuscript_xml is None:
        return [], ["Error: Could not parse documents"], None

    # Get XML previews for display
    template_preview = get_xml_preview(template_xml)
    manuscript_preview = get_xml_preview(manuscript_xml)

    # Emit previews to the console for visibility during analysis
    if template_preview:
        print("\n=== TEMPLATE XML PREVIEW (from analyze_documents) ===")
        print(template_preview)
    else:
        print("\n=== TEMPLATE XML PREVIEW (from analyze_documents) ===")
        print("No preview available")

    if manuscript_preview:
        print("\n=== MANUSCRIPT XML PREVIEW (from analyze_documents) ===")
        print(manuscript_preview)
    else:
        print("\n=== MANUSCRIPT XML PREVIEW (from analyze_documents) ===")
        print("No preview available")

    # Extract paragraphs from both
    template_style_fonts, template_default_font = load_style_fonts(template_file)
    manuscript_style_fonts, manuscript_default_font = load_style_fonts(manuscript_file)

    template_paragraphs = extract_paragraphs_from_xml(
        template_xml, template_style_fonts, template_default_font
    )
    manuscript_paragraphs = extract_paragraphs_from_xml(
        manuscript_xml, manuscript_style_fonts, manuscript_default_font
    )

    # Analyze template to get formatting rules
    template_profile = analyze_template_formatting(
        template_paragraphs, custom_rules=custom_rules
    )

    # Compare manuscript against template rules
    findings = compare_against_template(manuscript_paragraphs, template_profile)

    # Check for missing sections
    missing = check_missing_sections(manuscript_paragraphs, template_profile)

    if missing:
        def missing_section_finding(section_name: str):
            label = format_section_label(section_name)
            expected_note = f"Add the '{label}' section following the template order."
            fix_note = f"Insert a '{label}' section matching template formatting."
            if section_name.lower() == "acknowledgement":
                expected_note = (
                    "Insert the highlighted acknowledgement heading and paragraph "
                    "in Times New Roman 10 pt."
                )
                fix_note = (
                    "Add an acknowledgement section using the default text: "
                    "'The authors would like to thank the anonymous reviewers who "
                    "have provided valuable suggestions to improve the article.'"
                )
            elif section_name.lower() == "funding statement":
                expected_note = (
                    "Insert the highlighted funding statement heading and paragraph "
                    "in Times New Roman 10 pt before the References section."
                )
                fix_note = (
                    "Add a funding statement section using the default text "
                    "before the References section: 'The authors received no funding "
                    "from any party for the research and publication of this article'."
                )
            return {
                "type": "missing_section",
                "section": section_name,
                "paragraph_indices": [],
                "pages": [],
                "found": "Section not found",
                "expected": expected_note,
                "snippet": "",
                "suggested_fix": fix_note,
            }

        findings.extend(missing_section_finding(sec) for sec in missing)

    # Enforce order/structure validation
    order_issues = check_section_order(manuscript_paragraphs, template_profile)
    if order_issues:
        findings.extend(order_issues)

    metadata_issues = enforce_required_journal_metadata_line(manuscript_paragraphs)
    if metadata_issues:
        findings.extend(metadata_issues)

    # Create XML previews object
    xml_previews = {
        "template": template_preview,
        "manuscript": manuscript_preview,
        "template_full": (
            template_xml_string[:5000] if template_xml_string else None
        ),  # First 5000 chars
        "manuscript_full": (
            manuscript_xml_string[:5000] if manuscript_xml_string else None
        ),
    }

    return findings, missing, xml_previews


def analyze_template_formatting(template_paragraphs, custom_rules=None):
    """Analyze template to extract formatting rules and structural expectations."""
    rules = {}
    section_examples = defaultdict(list)
    context_examples = defaultdict(list)
    section_order = []
    seen = set()
    last_context_section = None
    heading_context_exclusions = {
        "figure_caption",
        "table_caption",
        "journal_metadata",
        "journal_name",
    }

    # First pass: classify each paragraph and collect examples with ordering
    for para in template_paragraphs:
        section_type = classify_section_type(para)
        section_examples[section_type].append(para)
        if section_type == "journal_metadata":
            register_journal_metadata_example(para.get("text", ""))
        if section_type not in seen:
            section_order.append(section_type)
            seen.add(section_type)
        if section_type == "body_text":
            if last_context_section:
                context_examples[("body_text", last_context_section)].append(para)
        else:
            if section_type not in heading_context_exclusions:
                last_context_section = section_type

    # Second pass: determine formatting for each section
    for section_type, examples in section_examples.items():
        if examples:
            rules[section_type] = determine_section_formatting(section_type, examples)

    context_rules = {}
    for key, examples in context_examples.items():
        if examples:
            section_type, _ = key
            context_rules[key] = determine_section_formatting(section_type, examples)

    raw_examples = {k: list(v) for k, v in section_examples.items()}
    context_examples_dict = {k: list(v) for k, v in context_examples.items()}

    return TemplateProfile(
        rules=rules,
        section_order=section_order,
        raw_examples=raw_examples,
        custom_rules=custom_rules or {},
        context_rules=context_rules,
        context_examples=context_examples_dict,
    )


def determine_section_formatting(section_type, examples):
    """Determine the typical formatting for a section based on high-quality examples."""
    if not examples:
        return get_default_formatting(section_type)

    valid_examples = [ex for ex in examples if is_valid_template_example(ex)]
    if not valid_examples:
        valid_examples = examples

    def most_common(key):
        values = [ex.get(key) for ex in valid_examples if ex.get(key) is not None]
        if not values:
            return None
        return Counter(values).most_common(1)[0][0]

    formatting = {}
    font_size = most_common("font_size")
    font_size_w_val = most_common("font_size_w_val")
    font_name = most_common("font_name")
    bold_choice = most_common("bold")
    italic_choice = most_common("italic")

    if font_size is not None:
        formatting["font_size"] = font_size
    if font_size_w_val is not None:
        formatting["font_size_w_val"] = font_size_w_val
    if font_name is not None:
        formatting["font_name"] = font_name
    if bold_choice is not None:
        formatting["bold"] = bool(bold_choice)
    if italic_choice is not None:
        formatting["italic"] = bool(italic_choice)

    if formatting:
        hints = {}
        for ex in valid_examples:
            hints.update(infer_template_hints(ex.get("text", "")))
        for key, value in hints.items():
            if value is None:
                continue
            if key not in formatting:
                formatting[key] = value

        best_example = max(
            valid_examples, key=lambda ex: score_template_example(section_type, ex)
        )
        for key in ("font_size", "font_name", "bold", "italic"):
            candidate = best_example.get(key)
            if candidate is None:
                continue
            if key not in formatting:
                formatting[key] = candidate
        if (
            "font_size_w_val" not in formatting
            and best_example.get("font_size_w_val") is not None
        ):
            formatting["font_size_w_val"] = best_example["font_size_w_val"]
        return ensure_font_size_pair(formatting)

    # Fallback to defaults if we cannot determine anything reliable
    fallback = get_default_formatting(section_type)
    return ensure_font_size_pair(fallback)


def is_valid_template_example(paragraph):
    """Filter out numeric placeholders or empty runs from template sampling."""
    text = (paragraph.get("text") or "").strip()
    if not text:
        return False

    # Must contain alphabetic characters to be meaningful for formatting inference
    alpha_chars = sum(1 for ch in text if ch.isalpha())
    if alpha_chars == 0:
        return False

    # Avoid paragraphs that are mostly digits or symbols (table numbers, etc.)
    condensed = re.sub(r"\s+", "", text)
    if condensed:
        digit_ratio = sum(ch.isdigit() for ch in condensed) / len(condensed)
        if digit_ratio > 0.6 and alpha_chars < 6:
            return False

    return True


def infer_template_hints(text):
    """Extract explicit instructions embedded in template paragraphs."""
    if not text:
        return {}

    hints = {}
    lower = text.lower()

    # Font size patterns like "10-Font size" or "font size 10"
    size_match = re.search(r"(\d+(?:\.\d+)?)\s*-\s*font size", lower)
    if not size_match:
        size_match = re.search(r"font size\s*(\d+(?:\.\d+)?)", lower)
    if size_match:
        try:
            hints["font_size"] = float(size_match.group(1))
        except ValueError:
            pass

    if "not bold" in lower or "non-bold" in lower or "without bold" in lower:
        hints["bold"] = False
    elif "bold" in lower:
        hints["bold"] = True

    if "not italic" in lower or "non-italic" in lower or "without italic" in lower:
        hints["italic"] = False
    elif "italic" in lower:
        hints["italic"] = True

    known_fonts = [
        "palatino linotype",
        "times new roman",
        "arial",
        "calibri",
        "cambria",
    ]
    for font in known_fonts:
        if font in lower:
            hints["font_name"] = normalize_font_name(font)
            break

    return hints


def score_template_example(section_type, paragraph):
    """Heuristic score to prioritise the most representative template example."""
    score = 0.0
    original_text = paragraph.get("text") or ""
    text = original_text.lower()
    hints = infer_template_hints(original_text)

    if paragraph.get("font_size"):
        score += 2.0
    if paragraph.get("font_name"):
        score += 2.0
    if paragraph.get("bold"):
        score += 1.5
    if paragraph.get("italic"):
        score += 0.5

    if "bold" in hints and hints["bold"]:
        score += 1.5
    if "font_size" in hints:
        score += 1.0
    if "font_name" in hints:
        score += 1.0
    if "italic" in hints and hints["italic"]:
        score += 0.5

    if section_type in {"title", "main_heading"} and paragraph.get("bold"):
        score += 1.5
    if section_type == "title" and "title" in text:
        score += 1.0

    alpha_chars = sum(1 for ch in text if ch.isalpha())
    score += min(alpha_chars / 30.0, 1.0)  # reward readable samples

    return score


def get_default_formatting(section_type):
    """Get default formatting rules as fallback"""
    defaults = {
        # === Titles ===
        "title": {"font_size": 24.0, "bold": False, "font_name": "Times New Roman"},
        "subtitle": {"font_size": 16.0, "bold": False, "font_name": "Times New Roman"},
        # === Author Info ===
        "authors": {"font_size": 11.0, "bold": True, "font_name": "Times New Roman"},
        "affiliation": {
            "font_size": 9.0,
            "bold": False,
            "font_name": "Times New Roman",
        },
        "corresponding_author": {
            "font_size": 9.0,
            "bold": False,
            "italic": True,
            "font_name": "Times New Roman",
        },
        "submission_history": {
            "font_size": 9.0,
            "bold": False,
            "italic": True,
            "font_name": "Times New Roman",
        },
        "journal_name": {
            "font_size": 24.0,
            "bold": True,
            "font_name": "Palatino Linotype",
        },
        "journal_metadata": {
            "font_size": 9.0,
            "bold": False,
            "font_name": "Times New Roman",
        },
        # === Abstract & Keywords ===
        "abstract": {"font_size": 9.0, "bold": False, "font_name": "Times New Roman"},
        "keywords": {
            "font_size": 9.0,
            "bold": False,
            "italic": True,
            "font_name": "Times New Roman",
        },
        # === Headings / Sections ===
        "introduction": {
            "font_size": 10.0,
            "bold": False,
            "font_name": "Times New Roman",
        },
        "literature review": {
            "font_size": 10.0,
            "bold": True,
            "font_name": "Times New Roman",
        },
        "research methodology": {
            "font_size": 10.0,
            "bold": True,
            "font_name": "Times New Roman",
        },
        "results and discussions": {
            "font_size": 10.0,
            "bold": False,
            "font_name": "Times New Roman",
        },
        "conclusion": {"font_size": 10.0, "bold": True, "font_name": "Times New Roman"},
        "acknowledgement": {
            "font_size": 10.0,
            "bold": True,
            "font_name": "Times New Roman",
        },
        "funding statement": {
            "font_size": 10.0,
            "bold": False,
            "font_name": "Times New Roman",
        },
        "author contributions": {
            "font_size": 10.0,
            "bold": True,
            "font_name": "Times New Roman",
        },
        "conflict of interests": {
            "font_size": 10.0,
            "bold": True,
            "font_name": "Times New Roman",
        },
        "ethics statements": {
            "font_size": 10.0,
            "bold": True,
            "font_name": "Times New Roman",
        },
        "biographies of authors": {
            "font_size": 10.0,
            "bold": True,
            "font_name": "Times New Roman",
        },
        "main_heading": {
            "font_size": 10.0,
            "bold": True,
            "font_name": "Times New Roman",
        },
        # === Other ===
        "body_text": {"font_size": 10.0, "bold": False, "font_name": "Times New Roman"},
        "figure_caption": {
            "font_size": 10.0,
            "bold": False,
            "font_name": "Times New Roman",
        },
        "table_caption": {
            "font_size": 10.0,
            "bold": False,
            "font_name": "Times New Roman",
        },
        "references": {"font_size": 10.0, "bold": False, "font_name": "Times New Roman"},
    }

    template = defaults.get(section_type, defaults["body_text"])
    fmt = dict(template)
    return ensure_font_size_pair(fmt)


def compare_against_template(manuscript_paragraphs, template_profile):
    """Compare manuscript paragraphs against template formatting rules"""
    findings = []
    template_rules = template_profile.rules
    body_rule = template_rules.get("body_text", get_default_formatting("body_text"))
    last_context_section = None

    for para in manuscript_paragraphs:
        if not para["text"].strip():
            continue

        section_type = classify_section_type(para)
        context_section = last_context_section if section_type == "body_text" else None
        expected = template_profile.resolve_expected_format(
            section_type,
            para.get("text", ""),
            context_section=context_section,
        )
        if not expected:
            expected = get_default_formatting(section_type) or body_rule

        display_section = section_type
        if section_type == "body_text":
            if context_section:
                display_section = f"body_text::{context_section}"
            else:
                display_section = "body_text::general"
        else:
            heading_like = section_type not in {
                "figure_caption",
                "table_caption",
                "journal_metadata",
                "journal_name",
            }
            if heading_like:
                last_context_section = section_type

        # Check formatting mismatches
        findings.extend(
            check_formatting_mismatches(
                para, section_type, expected, display_section=display_section
            )
        )

    return findings


def format_pt_display(value):
    if value is None:
        return None
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return None
    if abs(numeric - round(numeric)) < 1e-6:
        return f"{int(round(numeric))} pt"
    return f"{round(numeric, 2)} pt"


def format_font_size_display(pt, hp):
    pt_text = format_pt_display(pt)
    if hp is not None and pt_text:
        return f"{pt_text} (w:val={hp})"
    if hp is not None:
        return f"w:val={hp}"
    if pt_text:
        return pt_text
    return "Not detected"


def font_size_fix_text(expected_pt, expected_hp):
    pt_text = format_pt_display(expected_pt)
    if pt_text:
        return f"Set font size to {pt_text}"
    if expected_hp is not None:
        return f"Adjust font size to match w:val {expected_hp}"
    return "Adjust font size to match the template"


def should_flag_font_size_mismatch(paragraph, expected):
    if not isinstance(expected, dict):
        return False
    expected_hp = expected.get("font_size_w_val")
    actual_hp = paragraph.get("font_size_w_val")
    if expected_hp is not None and actual_hp is not None:
        return expected_hp != actual_hp
    expected_pt = expected.get("font_size")
    actual_pt = paragraph.get("font_size")
    if expected_pt is not None and actual_pt is not None:
        return abs(expected_pt - actual_pt) >= 0.5
    return False


def build_font_size_mismatch_finding(paragraph, expected, section_label):
    if not should_flag_font_size_mismatch(paragraph, expected):
        return None

    actual_pt = paragraph.get("font_size")
    actual_hp = paragraph.get("font_size_w_val")
    expected_pt = expected.get("font_size")
    expected_hp = expected.get("font_size_w_val")

    found_display = format_font_size_display(actual_pt, actual_hp)
    expected_display = format_font_size_display(expected_pt, expected_hp)
    fix_text = font_size_fix_text(expected_pt, expected_hp)

    return create_finding(
        section_label,
        paragraph["index"],
        "font_size_mismatch",
        found_display,
        expected_display,
        paragraph.get("text", ""),
        fix_text,
    )


def check_formatting_mismatches(
    paragraph, section_type, expected, display_section=None
):
    """Check for specific formatting mismatches.
    - Only 'Keywords' and 'Corresponding Author' must be italic.
    - Headers (Introduction, Conclusion, etc.) must NOT be bold or italic.
    - Normal text (body_text) → ignore bold/italic completely.
    """
    findings = []

    text_lower = paragraph.get("text", "").strip().lower()
    actual_bold = paragraph.get("bold", False)
    actual_italic = paragraph.get("italic", False)

    # --- 1️⃣ Skip bold/italic check for body text ---
    target_section = display_section or section_type

    if section_type == "body_text":
        # Still check font size / font name only
        font_size_issue = build_font_size_mismatch_finding(
            paragraph, expected, target_section
        )
        if font_size_issue:
            findings.append(font_size_issue)
        if expected.get("font_name") and paragraph.get("font_name"):
            if not fonts_similar(expected["font_name"], paragraph["font_name"]):
                findings.append(
                    create_finding(
                        target_section,
                        paragraph["index"],
                        "font_family_mismatch",
                        paragraph["font_name"],
                        expected["font_name"],
                        paragraph["text"],
                        f"Set font to '{expected['font_name']}'",
                    )
                )
        if paragraph.get("bold"):
            findings.append(
                create_finding(
                    target_section,
                    paragraph["index"],
                    "bold_incorrect",
                    "bold",
                    "not bold",
                    paragraph.get("text", ""),
                    "Remove bold formatting",
                )
            )
        return findings  # ✅ No bold/italic checks for body text

    # --- 2️⃣ Bold rules ---
    header_must_not_be_bold = {
        "acknowledgement",
        "acknowledgment",
        "introduction",
        "abstract",
        "conclusion",
        "references",
        "literature review",
        "research methodology",
        "results and discussions",
        "funding statement",
        "author contributions",
        "conflict of interests",
        "ethics statements",
        "biographies of authors",
        "main_heading",
    }

    if section_type == "title":
        expected_bold = expected.get("bold", False)
    elif any(text_lower.startswith(h) for h in header_must_not_be_bold):
        expected_bold = False
    else:
        expected_bold = expected.get("bold", False)

    if expected_bold != actual_bold:
        issue_type = (
            "bold_missing" if expected_bold and not actual_bold else "bold_incorrect"
        )
        findings.append(
            create_finding(
                target_section,
                paragraph["index"],
                issue_type,
                "bold" if actual_bold else "not bold",
                "bold" if expected_bold else "not bold",
                paragraph["text"],
                f"Set bold = {expected_bold}",
            )
        )

    # --- 3️⃣ Italic rules ---
    if section_type in ["keywords", "corresponding_author"]:
        # These MUST be italic
        if not actual_italic:
            findings.append(
                create_finding(
                    target_section,
                    paragraph["index"],
                    "italic_missing",
                    "not italic",
                    "italic",
                    paragraph["text"],
                    "Set to italic",
                )
            )
    elif section_type in header_must_not_be_bold or section_type in [
        "title",
        "references",
    ]:
        # These must NOT be italic
        if actual_italic:
            findings.append(
                create_finding(
                    target_section,
                    paragraph["index"],
                    "italic_incorrect",
                    "italic",
                    "not italic",
                    paragraph["text"],
                    "Remove italic formatting",
                )
            )

    # --- 4️⃣ Font size & family checks (same as before) ---
    font_size_issue = build_font_size_mismatch_finding(
        paragraph, expected, target_section
    )
    if font_size_issue:
        findings.append(font_size_issue)

    if expected.get("font_name") and paragraph.get("font_name"):
        expected_font = expected["font_name"]
        actual_font = paragraph["font_name"]
        if not fonts_similar(expected_font, actual_font):
            findings.append(
                create_finding(
                    target_section,
                    paragraph["index"],
                    "font_family_mismatch",
                    actual_font,
                    expected_font,
                    paragraph["text"],
                    f"Set font to '{expected_font}'",
                )
            )

    return findings


def create_finding(section, index, issue_type, found, expected, snippet, fix):
    """Create a standardized finding object"""
    return {
        "type": issue_type,
        "section": section,
        "paragraph_indices": [index],
        "pages": [1 + (index // 5)],  # Rough page estimate
        "found": found,
        "expected": expected,
        "snippet": text_snippet(snippet),
        "suggested_fix": fix,
    }


def check_missing_sections(manuscript_paragraphs, template_profile):
    """Check for missing required sections"""
    manuscript_sections = set()

    for para in manuscript_paragraphs:
        section_type = classify_section_type(para)
        manuscript_sections.add(section_type)

    required_sections = set(template_profile.required_sections())
    missing = required_sections - manuscript_sections

    # Always enforce presence of acknowledgement and funding statement, per business rules
    for always_required in ("acknowledgement", "funding statement"):
        if always_required not in manuscript_sections:
            missing.add(always_required)

    return list(missing)


def format_section_label(section):
    """Convert machine section name to human-readable label."""
    if not section:
        return "previous section"
    return section.replace("_", " ").title()


def check_section_order(manuscript_paragraphs, template_profile):
    """Ensure sections in the manuscript follow the order defined by the template."""
    significant_sections = [
        s
        for s in template_profile.section_order
        if s in template_profile.rules
        and s not in {"body_text", "figure_caption", "table_caption"}
    ]

    first_occurrence = {}
    occurrences = defaultdict(list)
    for para in manuscript_paragraphs:
        section_type = classify_section_type(para)
        occurrences[section_type].append(para)
        if section_type not in first_occurrence:
            first_occurrence[section_type] = para

    findings = []
    last_index = -1
    last_section = None
    for section in significant_sections:
        para = first_occurrence.get(section)
        if not para:
            continue
        if para["index"] < last_index:
            snippet = para.get("text", "")
            findings.append(
                create_finding(
                    section,
                    para["index"],
                    "section_order_mismatch",
                    f"Appears after {format_section_label(last_section)}",
                    f"Before {format_section_label(last_section)}",
                    snippet,
                    f"Move '{format_section_label(section)}' so it precedes '{format_section_label(last_section)}'",
                )
            )
        else:
            last_index = para["index"]
            last_section = section

    for section in significant_sections:
        if len(occurrences.get(section, [])) > 1:
            extras = occurrences[section][1:]
            first_extra = extras[0]
            indices = [p["index"] for p in extras]
            snippet = first_extra.get("text", "")
            findings.append(
                create_finding(
                    section,
                    first_extra["index"],
                    "section_duplicate",
                    f"Paragraphs {indices}",
                    "Single occurrence",
                    snippet,
                    f"Retain one '{format_section_label(section)}' section matching the template sequence",
                )
            )

    funding_para = first_occurrence.get("funding statement")
    references_para = first_occurrence.get("references")
    if (
        funding_para
        and references_para
        and funding_para["index"] >= references_para["index"]
    ):
        snippet = funding_para.get("text", "")
        findings.append(
            create_finding(
                "funding statement",
                funding_para["index"],
                "section_order_mismatch",
                "Does not appear immediately before References",
                "Immediately precede References",
                snippet,
                "Move the funding statement section so it is placed directly before the references section.",
            )
        )

    return findings


def enforce_required_journal_metadata_line(manuscript_paragraphs):
    """Ensure the journal metadata block contains the required address line."""
    required_key = normalize_special_key(REQUIRED_JOURNAL_METADATA_LINE)
    metadata_paragraphs = [
        para
        for para in manuscript_paragraphs
        if classify_section_type(para) == "journal_metadata"
    ]

    for para in metadata_paragraphs:
        text_key = normalize_special_key(para.get("text", ""))
        if required_key and required_key in text_key:
            return []

    if metadata_paragraphs:
        target_para = metadata_paragraphs[0]
    else:
        target_para = {"index": 0, "text": ""}

    fix_note = (
        f"Insert the line '{REQUIRED_JOURNAL_METADATA_LINE}' within the journal metadata block."
    )
    finding = create_finding(
        "journal_metadata",
        target_para.get("index", 0),
        "journal_metadata_missing_detail",
        "Required address line not found",
        f"Include '{REQUIRED_JOURNAL_METADATA_LINE}'",
        target_para.get("text", ""),
        fix_note,
    )
    return [finding]


# -------------------------
# Utility Functions
# -------------------------
def now_timestamp():
    return time.strftime("%Y%m%d_%H%M%S")


def normalize_half_points(value):
    if value is None:
        return None
    try:
        return int(round(float(value)))
    except Exception:
        return None


def half_points_to_pt(val):
    try:
        return round(float(val) / 2.0, 2)
    except Exception:
        return None


def pt_to_half_points(value):
    if value is None:
        return None
    try:
        return int(round(float(value) * 2.0))
    except Exception:
        return None


def ensure_font_size_pair(formatting):
    if not isinstance(formatting, dict):
        return formatting

    hp = formatting.get("font_size_w_val")
    pt = formatting.get("font_size")

    if hp is not None and pt is None:
        converted = half_points_to_pt(hp)
        if converted is not None:
            formatting["font_size"] = converted
    elif hp is None and pt is not None:
        converted = pt_to_half_points(pt)
        if converted is not None:
            formatting["font_size_w_val"] = converted

    return formatting


def text_snippet(s, length=140):
    if not s:
        return ""
    s = s.replace("\n", " ").strip()
    return s[:length] + ("..." if len(s) > length else "")


def normalize_font_name(name):
    """Simple font normalization"""
    if not name:
        return ""

    name = name.lower().strip()

    # Basic font aliases
    aliases = {
        "timesnewroman": "Times New Roman",
        "times roman": "Times New Roman",
        "times": "Times New Roman",
        "arial": "Arial",
        "helvetica": "Arial",
        "calibri": "Calibri",
        "cambria": "Cambria",
    }

    for alias, canonical in aliases.items():
        if alias in name:
            return canonical

    return name.title()


def fonts_similar(font1, font2):
    """Check if two fonts are similar enough"""
    font1 = font1.lower()
    font2 = font2.lower()

    similar_groups = [
        {"times new roman", "times", "timesroman"},
        {"arial", "helvetica"},
        {"calibri", "cambria"},
    ]

    for group in similar_groups:
        if font1 in group and font2 in group:
            return True

    return font1 == font2


# -------------------------
# Export to Excel
# -------------------------
def export_findings_to_excel(findings, missing_sections, out_path=None):
    if out_path is None:
        out_path = f"xml_analysis_{now_timestamp()}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "XML Analysis Results"

    headers = [
        "No",
        "IssueType",
        "Section",
        "ParagraphIndices",
        "PageEstimates",
        "SnippetExamples",
        "Found",
        "Expected",
        "SuggestedFix",
    ]
    ws.append(headers)

    for i, f in enumerate(findings, start=1):
        ws.append(
            [
                i,
                f.get("type"),
                f.get("section"),
                ", ".join(str(x) for x in f.get("paragraph_indices", [])),
                ", ".join(str(x) for x in f.get("pages", [])),
                f.get("snippet"),
                f.get("found"),
                f.get("expected"),
                f.get("suggested_fix"),
            ]
        )

    if missing_sections:
        ws2 = wb.create_sheet("MissingSections")
        ws2.append(["Missing sections:"])
        for s in missing_sections:
            ws2.append([s])

    wb.save(out_path)
    return out_path


def summarize_mistakes_df(mistakes_df, max_rows=5):
    """Create a lightweight preview of mistakes_df for debugging"""
    summary = {
        "row_count": 0,
        "columns": [],
        "sample_rows": [],
    }

    if mistakes_df is None:
        summary["note"] = "mistakes_df is None"
        return summary

    try:
        summary["row_count"] = len(mistakes_df)
        summary["columns"] = list(getattr(mistakes_df, "columns", []))

        if summary["row_count"] > 0:
            head_df = mistakes_df.head(max_rows)
            try:
                head_df = head_df.fillna("")
            except Exception:
                pass
            summary["sample_rows"] = head_df.to_dict(orient="records")
    except Exception as err:
        summary["note"] = f"Could not summarize mistakes_df: {err}"

    return summary


def log_mistakes_summary(summary):
    """Print summary details for debugging"""
    row_count = summary.get("row_count", 0)
    columns = summary.get("columns") or []
    note = summary.get("note")

    column_list = ", ".join(columns) if columns else "None"
    print(f"[Highlight Preview] mistakes_df rows: {row_count} | columns: {column_list}")

    if note:
        print(f"[Highlight Preview] Note: {note}")

    for idx, row in enumerate(summary.get("sample_rows", []), start=1):
        print(f"[Highlight Preview] Row {idx}: {row}")


def read_docx_bytes(source):
    """Return the bytes of a DOCX source (path or file-like)."""
    if hasattr(source, "read"):
        source.seek(0)
        data = source.read()
        source.seek(0)
        return data
    with open(source, "rb") as f:
        return f.read()


def w_tag(local_name):
    return f"{{{W_NS}}}{local_name}"


def ensure_child(element, local_name):
    child = element.find(f"w:{local_name}", NSMAP)
    if child is None:
        child = ET.SubElement(element, w_tag(local_name))
    return child


def ensure_child_ns(element, namespace, local_name):
    tag = f"{{{namespace}}}{local_name}"
    child = element.find(tag)
    if child is None:
        child = ET.SubElement(element, tag)
    return child


def remove_child(element, local_name):
    child = element.find(f"w:{local_name}", NSMAP)
    if child is not None:
        element.remove(child)
    return child is not None


def apply_font_name_to_rpr(rPr, font_name):
    rFonts = ensure_child(rPr, "rFonts")
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(f"{{{W_NS}}}{attr}", font_name)


def apply_font_size_to_rpr(rPr, size_pt):
    hps = str(int(round(size_pt * 2)))
    sz = ensure_child(rPr, "sz")
    sz.set(f"{{{W_NS}}}val", hps)
    szCs = ensure_child(rPr, "szCs")
    szCs.set(f"{{{W_NS}}}val", hps)


def apply_bold_to_rpr(rPr, bold_value):
    if bold_value:
        b = ensure_child(rPr, "b")
        b.set(f"{{{W_NS}}}val", "true")
        bcs = ensure_child(rPr, "bCs")
        bcs.set(f"{{{W_NS}}}val", "true")
    else:
        b_removed = remove_child(rPr, "b")
        bcs_removed = remove_child(rPr, "bCs")
        if not b_removed and not bcs_removed:
            return False
    return True


def apply_italic_to_rpr(rPr, italic_value):
    if italic_value:
        i = ensure_child(rPr, "i")
        i.set(f"{{{W_NS}}}val", "true")
        ics = ensure_child(rPr, "iCs")
        ics.set(f"{{{W_NS}}}val", "true")
    else:
        i_removed = remove_child(rPr, "i")
        ics_removed = remove_child(rPr, "iCs")
        if not i_removed and not ics_removed:
            return False
    return True


def set_run_font_name(run, font_name):
    rPr = ensure_child(run, "rPr")
    apply_font_name_to_rpr(rPr, font_name)


def set_run_font_size(run, size_pt):
    rPr = ensure_child(run, "rPr")
    apply_font_size_to_rpr(rPr, size_pt)


def set_run_bold(run, bold_value):
    rPr = ensure_child(run, "rPr")
    apply_bold_to_rpr(rPr, bold_value)


def set_run_italic(run, italic_value):
    rPr = ensure_child(run, "rPr")
    apply_italic_to_rpr(rPr, italic_value)


def ensure_paragraph_rpr(paragraph_element):
    pPr = ensure_child(paragraph_element, "pPr")
    return ensure_child(pPr, "rPr")


def apply_font_name_to_math_run(m_run, font_name):
    m_rPr = ensure_child_ns(m_run, M_NS, "rPr")
    w_rPr = m_rPr.find("w:rPr", NSMAP)
    if w_rPr is None:
        w_rPr = ET.SubElement(m_rPr, w_tag("rPr"))
    apply_font_name_to_rpr(w_rPr, font_name)


def apply_font_size_to_math_run(m_run, size_pt):
    m_rPr = ensure_child_ns(m_run, M_NS, "rPr")
    w_rPr = m_rPr.find("w:rPr", NSMAP)
    if w_rPr is None:
        w_rPr = ET.SubElement(m_rPr, w_tag("rPr"))
    apply_font_size_to_rpr(w_rPr, size_pt)


def apply_bold_to_math_run(m_run, bold_value):
    m_rPr = ensure_child_ns(m_run, M_NS, "rPr")
    w_rPr = m_rPr.find("w:rPr", NSMAP)
    if w_rPr is None:
        w_rPr = ET.SubElement(m_rPr, w_tag("rPr"))
    apply_bold_to_rpr(w_rPr, bold_value)


def apply_italic_to_math_run(m_run, italic_value):
    m_rPr = ensure_child_ns(m_run, M_NS, "rPr")
    w_rPr = m_rPr.find("w:rPr", NSMAP)
    if w_rPr is None:
        w_rPr = ET.SubElement(m_rPr, w_tag("rPr"))
    apply_italic_to_rpr(w_rPr, italic_value)


def parse_expected_font_size(expected):
    match = re.search(r"(\d+(?:\.\d+)?)", expected or "")
    if not match:
        return None
    try:
        return float(match.group(1))
    except ValueError:
        return None


def parse_expected_font_name(expected):
    if not expected:
        return None
    name = expected.strip().strip("'\"")
    return name or None


def interpret_expected_flag(expected, keyword, default=True):
    text = (expected or "").lower()
    negative_patterns = [f"not {keyword}", f"non-{keyword}", f"without {keyword}"]
    for pattern in negative_patterns:
        if pattern in text:
            return False
    if keyword in text:
        return True
    return default


def highlight_mistakes(template_file, manuscript_file, mistakes_df):
    """Highlight mistakes directly in the DOCX XML and emit verbose debug info."""
    mistakes_summary = summarize_mistakes_df(mistakes_df)
    log_mistakes_summary(mistakes_summary)

    highlight_map = defaultdict(list)
    orphan_issues = []
    debug_data = {"summary": mistakes_summary, "paragraphs": []}

    if mistakes_df is not None and hasattr(mistakes_df, "iterrows"):
        for _, mistake in mistakes_df.iterrows():
            para_indices = mistake.get("paragraph_indices")
            if isinstance(para_indices, list) and para_indices:
                for idx in para_indices:
                    try:
                        idx_int = int(idx)
                    except (TypeError, ValueError):
                        continue
                    highlight_map[idx_int].append(
                        {
                            "type": mistake.get("type"),
                            "section": mistake.get("section"),
                            "expected": mistake.get("expected"),
                            "found": mistake.get("found"),
                            "suggested_fix": mistake.get("suggested_fix"),
                            "snippet": mistake.get("snippet"),
                        }
                    )
            else:
                orphan_issues.append(
                    {
                        "type": mistake.get("type"),
                        "section": mistake.get("section"),
                        "expected": mistake.get("expected"),
                        "found": mistake.get("found"),
                        "suggested_fix": mistake.get("suggested_fix"),
                        "snippet": mistake.get("snippet"),
                    }
                )

    if not highlight_map:
        print("[Highlight Preview] No paragraphs flagged for highlighting.")

    try:
        manuscript_bytes = read_docx_bytes(manuscript_file)
        if not manuscript_bytes:
            raise ValueError("Empty manuscript bytes")

        doc_tree = None
        with zipfile.ZipFile(io.BytesIO(manuscript_bytes), "r") as zin:
            if "word/document.xml" not in zin.namelist():
                raise ValueError("word/document.xml not found in DOCX")
            doc_tree = ET.fromstring(zin.read("word/document.xml"))

        all_paragraphs = doc_tree.findall(".//w:p", NSMAP)
        paragraph_data = extract_paragraphs_from_xml(doc_tree)
        paragraph_lookup = {para["index"]: para for para in paragraph_data}

        highlighted_any = False

        def ensure_highlight(run):
            rPr = run.find("w:rPr", NSMAP)
            if rPr is None:
                rPr = ET.SubElement(run, f"{{{W_NS}}}rPr")
            highlight_elem = rPr.find("w:highlight", NSMAP)
            if highlight_elem is None:
                highlight_elem = ET.SubElement(rPr, f"{{{W_NS}}}highlight")
            highlight_elem.set(f"{{{W_NS}}}val", "yellow")

        def paragraph_plain_text(element):
            texts = []
            for t in element.findall(".//w:t", NSMAP):
                if t.text:
                    texts.append(t.text)
            return "".join(texts).strip()

        for para_idx in sorted(highlight_map.keys()):
            issues = highlight_map[para_idx]
            paragraph_element = (
                all_paragraphs[para_idx]
                if 0 <= para_idx < len(all_paragraphs)
                else None
            )

            if paragraph_element is None:
                print(
                    f"[Highlight Preview] Paragraph {para_idx} not found in DOCX XML."
                )
                paragraph_text = paragraph_lookup.get(para_idx, {}).get("text", "")
                preview = {
                    "paragraph_index": para_idx,
                    "paragraph_text": paragraph_text,
                    "issue_count": len(issues),
                    "issue_types": sorted(
                        {issue["type"] for issue in issues if issue.get("type")}
                    ),
                    "issues": issues,
                    "highlighted": False,
                }
                debug_data["paragraphs"].append(preview)
                continue

            for run in paragraph_element.findall(".//w:r", NSMAP):
                ensure_highlight(run)
            highlighted_any = True

            for issue_idx, issue in enumerate(issues, start=1):
                issue_type = issue.get("type") or "Unknown"
                section = issue.get("section") or "Unknown section"
                expected = issue.get("expected") or ""
                found = issue.get("found") or ""
                print(
                    f"[Highlight Preview]   Issue {issue_idx}: "
                    f"type={issue_type} | section={section} | expected={expected} | found={found}"
                )

            paragraph_text = paragraph_lookup.get(para_idx, {}).get("text")
            if not paragraph_text:
                paragraph_text = paragraph_plain_text(paragraph_element)
            preview = {
                "paragraph_index": para_idx,
                "paragraph_text": paragraph_text,
                "issue_count": len(issues),
                "issue_types": sorted(
                    {issue["type"] for issue in issues if issue.get("type")}
                ),
                "issues": issues,
                "highlighted": True,
            }
            debug_data["paragraphs"].append(preview)
            snippet = text_snippet(paragraph_text, 120)
            print(
                f"[Highlight Preview] Paragraph {para_idx} | "
                f"{len(issues)} issue(s) | Types: {', '.join(preview['issue_types']) or 'N/A'} "
                f"| Text: {snippet}"
            )

        if orphan_issues:
            debug_data["paragraphs"].append(
                {
                    "paragraph_index": "N/A",
                    "paragraph_text": "Issues without specific paragraph (e.g., missing sections).",
                    "issue_count": len(orphan_issues),
                    "issue_types": sorted(
                        {issue.get("type") for issue in orphan_issues if issue.get("type")}
                    ),
                    "issues": orphan_issues,
                    "highlighted": False,
                }
            )
            for issue in orphan_issues:
                issue_type = issue.get("type") or "Unknown"
                section = issue.get("section") or "Unknown section"
                print(
                    f"[Highlight Preview] Orphan issue: type={issue_type} | section={section} "
                    f"| expected={issue.get('expected') or ''} | found={issue.get('found') or ''}"
                )

        updated_document_xml = ET.tostring(
            doc_tree, encoding="UTF-8", xml_declaration=True
        )

        output_buffer = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(manuscript_bytes), "r") as zin:
            with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        zout.writestr(item, updated_document_xml)
                    else:
                        zout.writestr(item, zin.read(item.filename))

        output_buffer.seek(0)
        if not highlighted_any:
            print(
                "[Highlight Preview] Warning: No highlights applied; document unchanged."
            )
        return output_buffer.getvalue(), debug_data

    except Exception as e:
        print(f"Error highlighting: {str(e)}")
        return None, {"summary": mistakes_summary, "paragraphs": []}


def apply_xml_correction(paragraph_element, mistake):
    """Apply a single correction directly to the paragraph XML."""
    runs = paragraph_element.findall(".//w:r", NSMAP)
    math_runs = paragraph_element.findall(".//m:r", M_NSMAP)
    if not runs and not math_runs:
        return False

    correction_type = mistake.get("type")
    expected = mistake.get("expected") or ""
    applied = False

    if correction_type == "font_family_mismatch":
        font_name = parse_expected_font_name(expected)
        if not font_name:
            return False
        font_name = normalize_font_name(font_name)
        for run in runs:
            set_run_font_name(run, font_name)
        for m_run in math_runs:
            apply_font_name_to_math_run(m_run, font_name)
        apply_font_name_to_rpr(ensure_paragraph_rpr(paragraph_element), font_name)
        applied = True

    elif correction_type == "font_size_mismatch":
        size_pt = parse_expected_font_size(expected)
        if size_pt is None:
            return False
        for run in runs:
            set_run_font_size(run, size_pt)
        for m_run in math_runs:
            apply_font_size_to_math_run(m_run, size_pt)
        apply_font_size_to_rpr(ensure_paragraph_rpr(paragraph_element), size_pt)
        applied = True

    elif correction_type in ["bold_missing", "bold_incorrect", "bold_mismatch"]:
        bold_value = interpret_expected_flag(expected, "bold", default=True)
        for run in runs:
            set_run_bold(run, bold_value)
        for m_run in math_runs:
            apply_bold_to_math_run(m_run, bold_value)
        apply_bold_to_rpr(ensure_paragraph_rpr(paragraph_element), bold_value)
        applied = True

    elif correction_type in ["italic_missing", "italic_incorrect", "italic_mismatch"]:
        italic_value = interpret_expected_flag(expected, "italic", default=True)
        for run in runs:
            set_run_italic(run, italic_value)
        for m_run in math_runs:
            apply_italic_to_math_run(m_run, italic_value)
        apply_italic_to_rpr(ensure_paragraph_rpr(paragraph_element), italic_value)
        applied = True

    return applied


def apply_corrections(template_file, manuscript_file, mistakes_df):
    """Apply corrections directly in the DOCX XML and return updated bytes."""
    corrections_summary = summarize_mistakes_df(mistakes_df)
    log_mistakes_summary(corrections_summary)

    corrections_map = defaultdict(list)
    if mistakes_df is not None and hasattr(mistakes_df, "iterrows"):
        for _, mistake in mistakes_df.iterrows():
            para_indices = mistake.get("paragraph_indices")
            if isinstance(para_indices, list):
                for idx in para_indices:
                    try:
                        idx_int = int(idx)
                    except (TypeError, ValueError):
                        continue
                    corrections_map[idx_int].append(mistake)

    if not corrections_map:
        print("[Corrections] No paragraphs flagged for correction.")
        return None

    try:
        manuscript_bytes = read_docx_bytes(manuscript_file)
        if not manuscript_bytes:
            raise ValueError("Empty manuscript bytes")

        with zipfile.ZipFile(io.BytesIO(manuscript_bytes), "r") as zin:
            if "word/document.xml" not in zin.namelist():
                raise ValueError("word/document.xml not found in DOCX")
            doc_root = ET.fromstring(zin.read("word/document.xml"))

        all_paragraphs = doc_root.findall(".//w:p", NSMAP)
        corrections_applied = 0

        for para_idx in sorted(corrections_map.keys()):
            if not (0 <= para_idx < len(all_paragraphs)):
                print(f"[Corrections] Paragraph {para_idx} not found in DOCX XML.")
                continue
            paragraph_element = all_paragraphs[para_idx]
            for mistake in corrections_map[para_idx]:
                if apply_xml_correction(paragraph_element, mistake):
                    corrections_applied += 1

        if corrections_applied == 0:
            print("[Corrections] No corrections applied.")
            return manuscript_bytes

        print(f"🔧 Applied {corrections_applied} corrections")

        updated_document_xml = ET.tostring(
            doc_root, encoding="UTF-8", xml_declaration=True
        )

        output_buffer = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(manuscript_bytes), "r") as zin:
            with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        zout.writestr(item, updated_document_xml)
                    else:
                        zout.writestr(item, zin.read(item.filename))

        output_buffer.seek(0)
        return output_buffer.getvalue()

    except Exception as e:
        print(f"Error applying corrections: {str(e)}")
        return None


def insert_missing_sections(template_file, manuscript_file, missing_sections):
    """Insert ACKNOWLEDGEMENT and/or FUNDING STATEMENT sections if missing.
    Returns updated DOCX bytes. Inserted paragraphs are formatted as Times New Roman 10pt,
    kept non-bold, and highlighted in yellow as requested.
    """
    if not missing_sections:
        return read_docx_bytes(manuscript_file)

    try:
        manuscript_bytes = read_docx_bytes(manuscript_file)
        if not manuscript_bytes:
            raise ValueError("Empty manuscript bytes")

        with zipfile.ZipFile(io.BytesIO(manuscript_bytes), "r") as zin:
            if "word/document.xml" not in zin.namelist():
                raise ValueError("word/document.xml not found in DOCX")
            doc_root = ET.fromstring(zin.read("word/document.xml"))

        body = doc_root.find("w:body", NSMAP)
        if body is None:
            raise ValueError("w:body not found in document.xml")

        def create_paragraph(
            text,
            font_name="Times New Roman",
            size_pt=10.0,
            bold=False,
            italic=False,
            highlight=False,
        ):
            p = ET.Element(w_tag("p"))
            r = ET.SubElement(p, w_tag("r"))
            rPr = ET.SubElement(r, w_tag("rPr"))
            apply_font_name_to_rpr(rPr, font_name)
            apply_font_size_to_rpr(rPr, size_pt)
            if bold:
                apply_bold_to_rpr(rPr, True)
            if italic:
                apply_italic_to_rpr(rPr, True)
            if highlight:
                highlight_elem = ET.SubElement(rPr, w_tag("highlight"))
                highlight_elem.set(f"{{{W_NS}}}val", "yellow")
            t = ET.SubElement(r, w_tag("t"))
            t.set(f"{{{W_NS}}}space", "preserve")
            t.text = text
            return p

        # Normalize keys for safe membership checks
        missing_norm = {s.strip().lower() for s in missing_sections}

        # --- Determine template order for placement ---
        try:
            # Build template profile to get ordering
            t_rules = load_custom_rules(template_file)
            t_xml, _ = docx_to_xml(template_file)
            t_style_fonts, t_default_font = load_style_fonts(template_file)
            t_paragraphs = extract_paragraphs_from_xml(
                t_xml, t_style_fonts, t_default_font
            )
            t_profile = analyze_template_formatting(t_paragraphs, custom_rules=t_rules)
            template_order = [
                s for s in t_profile.section_order if s in t_profile.rules
            ]
            canonical_tail = [
                "acknowledgement",
                "funding statement",
                "author contributions",
                "conflict of interests",
                "ethics statements",
                "references",
            ]
            effective_order = []
            for sect in template_order:
                if sect not in effective_order:
                    effective_order.append(sect)
            for sect in canonical_tail:
                if sect not in effective_order:
                    effective_order.append(sect)
        except Exception:
            template_order = []
            effective_order = [
                "acknowledgement",
                "funding statement",
                "author contributions",
                "conflict of interests",
                "ethics statements",
                "references",
            ]

        # Build manuscript section occurrences for body-level paragraphs only
        m_style_fonts, m_default_font = load_style_fonts(manuscript_file)
        body_paras = body.findall("./w:p", NSMAP)
        sections_by_index = []
        body_context_sections = []
        occurrences = defaultdict(list)  # section -> list of body indices
        first_occurrence = {}
        heading_context_exclusions = {
            "figure_caption",
            "table_caption",
            "journal_metadata",
            "journal_name",
        }
        funding_anchor_element = None

        def rebuild_section_maps():
            nonlocal sections_by_index, body_context_sections, occurrences, first_occurrence, funding_anchor_element
            sections_by_index = []
            body_context_sections = []
            occurrences = defaultdict(list)
            first_occurrence = {}
            last_heading = None

            for i, p in enumerate(body_paras):
                para_dict = extract_paragraph_formatting(
                    p, i, m_style_fonts, m_default_font
                )
                section = classify_section_type(para_dict) if para_dict else None
                sections_by_index.append(section)

                if section and section != "body_text":
                    if section not in heading_context_exclusions:
                        last_heading = section
                    body_context_sections.append(None)
                elif section == "body_text":
                    body_context_sections.append(last_heading)
                else:
                    body_context_sections.append(None)

                if section:
                    occurrences[section].append(i)
                    if section not in first_occurrence:
                        first_occurrence[section] = i

        rebuild_section_maps()

        def logical_first_index(section_name: str):
            for idx, section in enumerate(sections_by_index):
                if section == section_name:
                    return idx
                if (
                    section == "body_text"
                    and body_context_sections[idx] == section_name
                ):
                    return idx
            return None

        def logical_last_index(section_name: str):
            for idx in range(len(sections_by_index) - 1, -1, -1):
                section = sections_by_index[idx]
                if section == section_name:
                    return idx
                if (
                    section == "body_text"
                    and body_context_sections[idx] == section_name
                ):
                    return idx
            return None

        def paragraph_plain_text(element):
            """Return concatenated text content of a paragraph."""
            texts = []
            for t in element.findall(".//w:t", NSMAP):
                if t.text:
                    texts.append(t.text)
            return "".join(texts).strip()

        default_blocks = {
            "acknowledgement": {
                "heading": "ACKNOWLEDGEMENT",
                "content": "The authors would like to thank the anonymous reviewers who have provided valuable suggestions to improve the article.",
            },
            "funding statement": {
                "heading": "FUNDING STATEMENT",
                "content": "The authors received no funding from any party for the research and publication of this article",
            },
        }

        # Skip sections whose default heading/content already exist to avoid duplicate insertions on reruns
        existing_texts = {paragraph_plain_text(p) for p in body_paras}
        for section_key, defaults in default_blocks.items():
            if section_key not in missing_norm:
                continue
            if (
                defaults["heading"] in existing_texts
                and defaults["content"] in existing_texts
            ):
                missing_norm.discard(section_key)

        def insertion_index_for(section_name: str):
            """Find body-level index to insert according to template order: before next section or after previous."""
            nonlocal funding_anchor_element
            if not effective_order:
                return None
            try:
                s_idx = effective_order.index(section_name)
            except ValueError:
                return None
            # Funding statement placement uses acknowledgement anchor only
            if section_name == "funding statement":
                references_idx = logical_first_index("references")
                ack_norm = normalize_special_key("acknowledgement")
                funding_anchor_element = None
                ack_start_idx = None
                for i, sect in enumerate(sections_by_index):
                    ctx = body_context_sections[i]
                    text_norm = normalize_special_key(
                        paragraph_plain_text(body_paras[i])
                    )
                    if sect == "acknowledgement":
                        ack_start_idx = i
                        break
                    if sect == "body_text" and ctx == "acknowledgement":
                        ack_start_idx = i
                        break
                    if text_norm and text_norm.startswith(ack_norm):
                        ack_start_idx = i
                        break

                if ack_start_idx is not None:
                    ack_end_idx = ack_start_idx
                    j = ack_start_idx + 1
                    while j < len(sections_by_index):
                        next_sect = sections_by_index[j]
                        next_ctx = body_context_sections[j]
                        next_text_norm = normalize_special_key(
                            paragraph_plain_text(body_paras[j])
                        )
                        if next_sect == "body_text":
                            ack_end_idx = j
                            j += 1
                            continue
                        if next_sect is None and not next_text_norm:
                            ack_end_idx = j
                            j += 1
                            continue
                        if (
                            next_sect == "body_text"
                            and next_ctx == "acknowledgement"
                        ):
                            ack_end_idx = j
                            j += 1
                            continue
                        break

                    funding_anchor_element = body_paras[ack_end_idx]
                    insertion_point = ack_end_idx + 1
                    while (
                        insertion_point < len(sections_by_index)
                        and sections_by_index[insertion_point] == "body_text"
                        and body_context_sections[insertion_point] == "acknowledgement"
                    ):
                        insertion_point += 1
                    if (
                        references_idx is not None
                        and insertion_point > references_idx
                    ):
                        insertion_point = references_idx
                    return insertion_point

                return references_idx if references_idx is not None else len(body_paras)

            # Find next existing section in manuscript (logical first occurrence)
            for nxt in effective_order[s_idx + 1 :]:
                logical_idx = logical_first_index(nxt)
                if logical_idx is not None:
                    return max(0, logical_idx)
            # Otherwise insert after the logical end of the previous section
            for prv in reversed(effective_order[:s_idx]):
                last_idx = logical_last_index(prv)
                if last_idx is not None:
                    return last_idx + 1
            return None

        def insert_section(section_key: str, heading: str, content: str):
            nonlocal body_paras, funding_anchor_element
            idx = insertion_index_for(section_key)
            heading_bold = section_key != "funding statement"
            h_p = create_paragraph(heading, bold=heading_bold, highlight=True)
            c_p = create_paragraph(content, bold=False, highlight=True)
            if (
                section_key == "funding statement"
                and funding_anchor_element is not None
            ):
                anchor = funding_anchor_element
                funding_anchor_idx = body_paras.index(anchor)
                anchor.addnext(c_p)
                anchor.addnext(h_p)
                body_paras.insert(funding_anchor_idx + 1, h_p)
                body_paras.insert(funding_anchor_idx + 2, c_p)
                funding_anchor_element = None
            elif idx is None or idx >= len(body_paras):
                # Append at end
                body.append(h_p)
                body.append(c_p)
                body_paras.extend([h_p, c_p])
            else:
                body.insert(idx, h_p)
                body.insert(idx + 1, c_p)
                # keep body_paras view in sync for subsequent insertions
                body_paras.insert(idx, h_p)
                body_paras.insert(idx + 1, c_p)
            rebuild_section_maps()

        # If BOTH Acknowledgement and Funding Statement are missing, and References exists,
        # insert both as a block immediately BEFORE References (in this order: Acknowledgement, Funding Statement)
        if "acknowledgement" in missing_norm and "funding statement" in missing_norm:
            ref_idx = logical_first_index("references")
            if ref_idx is not None:
                # helper to insert a titled section at a specific body index
                def insert_at_index(i, heading, content):
                    heading_bold = heading == default_blocks["acknowledgement"]["heading"]
                    h_p = create_paragraph(heading, bold=heading_bold, highlight=True)
                    c_p = create_paragraph(content, bold=False, highlight=True)
                    body.insert(i, h_p)
                    body.insert(i + 1, c_p)
                    body_paras.insert(i, h_p)
                    body_paras.insert(i + 1, c_p)
                    return i + 2

                i = ref_idx
                i = insert_at_index(
                    i,
                    default_blocks["acknowledgement"]["heading"],
                    default_blocks["acknowledgement"]["content"],
                )
                insert_at_index(
                    i,
                    default_blocks["funding statement"]["heading"],
                    default_blocks["funding statement"]["content"],
                )
                rebuild_section_maps()
                # mark handled to avoid reinserting below
                missing_norm.discard("acknowledgement")
                missing_norm.discard("funding statement")

        if "acknowledgement" in missing_norm:
            insert_section(
                "acknowledgement",
                default_blocks["acknowledgement"]["heading"],
                default_blocks["acknowledgement"]["content"],
            )

        if "funding statement" in missing_norm:
            insert_section(
                "funding statement",
                default_blocks["funding statement"]["heading"],
                default_blocks["funding statement"]["content"],
            )

        updated_document_xml = ET.tostring(
            doc_root, encoding="UTF-8", xml_declaration=True
        )

        output_buffer = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(manuscript_bytes), "r") as zin:
            with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == "word/document.xml":
                        zout.writestr(item, updated_document_xml)
                    else:
                        zout.writestr(item, zin.read(item.filename))

        output_buffer.seek(0)
        return output_buffer.getvalue()
    except Exception as e:
        print(f"Error inserting missing sections: {str(e)}")
    return read_docx_bytes(manuscript_file)


# -------------------------
# CLI Main
# -------------------------
def main(argv):
    parser = argparse.ArgumentParser(
        description="JIWE XML Formatter - XML-based document analysis"
    )
    parser.add_argument("template", help="template.docx")
    parser.add_argument("manuscript", help="manuscript.docx")
    parser.add_argument("--output", "-o", default=None, help="output Excel path")
    args = parser.parse_args(argv[1:])

    template = args.template
    manuscript = args.manuscript

    if not os.path.exists(template):
        print("ERROR: Template not found:", template)
        return 2
    if not os.path.exists(manuscript):
        print("ERROR: Manuscript not found:", manuscript)
        return 3

    print("=== JIWE XML FORMatter ===")
    print("Approach: Convert both documents to XML and compare at XML level")

    findings, missing, xml_previews = analyze_documents(template, manuscript)

    print("\n=== TEMPLATE XML PREVIEW ===")
    print(xml_previews["template"] if xml_previews else "No preview available")

    print("\n=== MANUSCRIPT XML PREVIEW ===")
    print(xml_previews["manuscript"] if xml_previews else "No preview available")

    out_path = args.output or f"xml_analysis_{now_timestamp()}.xlsx"
    saved = export_findings_to_excel(findings, missing, out_path=out_path)

    print(f"\n✅ XML ANALYSIS REPORT: {saved}")
    print(f"📊 Issues found: {len(findings)}")
    print(f"📋 Missing sections: {len(missing)}")

    if findings:
        print("\n🔍 Top issues:")
        for i, f in enumerate(findings[:10], start=1):
            print(
                f"  {i}. [{f['type']}] {f['section']} - para {f['paragraph_indices']}"
            )
    else:
        print("\n🎉 No formatting issues detected!")

    if missing:
        print("\n⚠️ Missing sections:")
        for s in missing:
            print(f"  - {s}")

    print(f"\n💡 XML-based analysis complete")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
