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
from collections import defaultdict, Counter
from lxml import etree as ET
from openpyxl import Workbook
import tempfile
import io
import shutil

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {'w': W_NS}

# -------------------------
# XML Conversion & Display Functions
# -------------------------
def docx_to_xml(docx_file):
    """Convert DOCX file to XML structure and return both root and string"""
    if hasattr(docx_file, 'read'):
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
        with zipfile.ZipFile(docx_path, 'r') as z:
            xml_content = z.read('word/document.xml')
        
        # Parse XML and also return pretty string
        root = ET.fromstring(xml_content)
        xml_string = ET.tostring(root, encoding='unicode', pretty_print=True)
        
        return root, xml_string
        
    except Exception as e:
        print(f"Error converting DOCX to XML: {str(e)}")
        return None, None
    finally:
        # Clean up temp file if we created one
        if hasattr(docx_file, 'read'):
            try:
                os.unlink(docx_path)
            except:
                pass

def get_xml_preview(xml_root, max_paragraphs=10):
    """Get a preview of XML structure for display"""
    if not xml_root:
        return "Error: Could not parse XML"
    
    preview_lines = []
    paragraphs = xml_root.findall('.//w:p', NSMAP)[:max_paragraphs]
    
    for idx, p in enumerate(paragraphs):
        preview_lines.append(f"\n=== Paragraph {idx} ===")
        
        # Get text content
        texts = []
        for r in p.findall('.//w:r', NSMAP):
            for t in r.findall('.//w:t', NSMAP):
                if t.text:
                    texts.append(t.text)
        
        if texts:
            full_text = ' '.join(texts)
            preview_lines.append(f"Text: {full_text[:100]}{'...' if len(full_text) > 100 else ''}")
        
        # Get formatting info
        pPr = p.find('./w:pPr', NSMAP)
        if pPr is not None:
            pStyle = pPr.find('./w:pStyle', NSMAP)
            if pStyle is not None:
                preview_lines.append(f"Style: {pStyle.get('{%s}val' % W_NS)}")
        
        # Get run formatting
        for r_idx, r in enumerate(p.findall('.//w:r', NSMAP)[:3]):  # First 3 runs
            rPr = r.find('./w:rPr', NSMAP)
            if rPr is not None:
                format_info = []
                
                # Font size
                sz = rPr.find('./w:sz', NSMAP)
                if sz is not None:
                    size_val = sz.get('{%s}val' % W_NS)
                    if size_val:
                        format_info.append(f"size:{half_points_to_pt(size_val)}pt")
                
                # Font name
                rf = rPr.find('./w:rFonts', NSMAP)
                if rf is not None:
                    font_name = rf.get('{%s}ascii' % W_NS) or rf.get('{%s}hAnsi' % W_NS)
                    if font_name:
                        format_info.append(f"font:{normalize_font_name(font_name)}")
                
                # Bold
                if rPr.find('./w:b', NSMAP) is not None:
                    format_info.append("bold")
                
                # Italic
                if rPr.find('./w:i', NSMAP) is not None:
                    format_info.append("italic")
                
                if format_info:
                    preview_lines.append(f"  Run {r_idx}: {', '.join(format_info)}")
    
    return "\n".join(preview_lines)

# -------------------------
# Enhanced Figure Detection
# -------------------------
def extract_paragraphs_from_xml(xml_root):
    """Extract all paragraphs with formatting from XML"""
    paragraphs = []
    
    for idx, p in enumerate(xml_root.findall('.//w:p', NSMAP)):
        paragraph_data = extract_paragraph_formatting(p, idx)
        if paragraph_data and paragraph_data['text'].strip():
            paragraphs.append(paragraph_data)
    
    return paragraphs

def extract_paragraph_formatting(paragraph, index):
    """Extract formatting and text from a paragraph element"""
    # Get all text runs
    texts = []
    runs_data = []
    
    for r in paragraph.findall('.//w:r', NSMAP):
        run_text = ''.join(t.text for t in r.findall('.//w:t', NSMAP) if t.text)
        if run_text:
            run_data = extract_run_formatting(r)
            run_data['text'] = run_text
            runs_data.append(run_data)
            texts.append(run_text)
    
    if not texts:
        return None
    
    full_text = ' '.join(texts).strip()
    
    # Get paragraph style
    p_style = None
    pPr = paragraph.find('./w:pPr', NSMAP)
    if pPr is not None:
        pStyle = pPr.find('./w:pStyle', NSMAP)
        if pStyle is not None:
            p_style = pStyle.get('{%s}val' % W_NS)
    
    # Determine dominant formatting from runs
    dominant_format = get_dominant_formatting(runs_data)
    
    return {
        'index': index,
        'text': full_text,
        'p_style': p_style,
        'font_size': dominant_format.get('font_size'),
        'font_name': dominant_format.get('font_name'),
        'bold': dominant_format.get('bold', False),
        'italic': dominant_format.get('italic', False),
        'runs': runs_data
    }

def extract_run_formatting(run):
    """Extract formatting from a run element"""
    format_data = {}
    
    rPr = run.find('./w:rPr', NSMAP)
    if rPr is None:
        return format_data
    
    # Font size
    sz = rPr.find('./w:sz', NSMAP)
    if sz is not None and sz.get('{%s}val' % W_NS):
        format_data['font_size'] = half_points_to_pt(sz.get('{%s}val' % W_NS))
    
    # Font name
    rf = rPr.find('./w:rFonts', NSMAP)
    if rf is not None:
        font_name = rf.get('{%s}ascii' % W_NS) or rf.get('{%s}hAnsi' % W_NS)
        if font_name:
            format_data['font_name'] = normalize_font_name(font_name)
    
    # Bold
    if rPr.find('./w:b', NSMAP) is not None:
        format_data['bold'] = True
    
    # Italic
    if rPr.find('./w:i', NSMAP) is not None:
        format_data['italic'] = True
    
    return format_data

def get_dominant_formatting(runs_data):
    """Determine dominant formatting from multiple runs"""
    if not runs_data:
        return {}
    
    font_sizes = []
    font_names = []
    bold_count = 0
    italic_count = 0
    total_runs = len(runs_data)
    
    for run in runs_data:
        if run.get('font_size'):
            font_sizes.append(run['font_size'])
        if run.get('font_name'):
            font_names.append(run['font_name'])
        if run.get('bold'):
            bold_count += 1
        if run.get('italic'):
            italic_count += 1
    
    dominant = {}
    
    if font_sizes:
        dominant['font_size'] = Counter(font_sizes).most_common(1)[0][0]
    
    if font_names:
        dominant['font_name'] = Counter(font_names).most_common(1)[0][0]
    
    if bold_count > total_runs / 2:
        dominant['bold'] = True
    if italic_count > total_runs / 2:
        dominant['italic'] = True
    
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
    raw = (paragraph.get('text') or '').strip()
    if not raw:
        return 'body_text'

    # normalize
    text = raw.replace('\r', ' ').replace('\n', ' ').strip()
    text = re.sub(r'\s+', ' ', text)              # collapse whitespace
    lower_text = text.lower()
    font_size = paragraph.get('font_size') or 0
    p_style = (paragraph.get('p_style') or '').lower()
    idx = paragraph.get('index', 999)

    # helper: remove leading numbering like "1.", "I.", "1.1", "1 -", "1 Introduction" etc.
    cleaned = re.sub(r'^[\s\-\–\—]*(?:[ivxlcdmIVXLCDM]+|[\d]+(?:[.\d]*)?)[\s\.\-\:\)]*', '', lower_text).strip()
    # also strip common bullets/markers
    cleaned = re.sub(r'^[\*\•\·\u2022\-\–\—\•]+\s*', '', cleaned)

    # 1) Paragraph style strong hints
    if p_style:
        if 'title' in p_style:
            return 'title'
        if 'heading' in p_style or p_style.startswith('h'):
            # try to resolve to known section names
            for sect in ('introduction', 'abstract', 'keywords', 'conclusion', 'references'):
                if re.search(r'\b' + re.escape(sect) + r'\b', lower_text):
                    return sect
            return 'main_heading'

    # 2) Large-font title
    if font_size and font_size >= 20:
        return 'title'

    # 3) Figure captions
    for pattern in [r'^(figure|fig)\b', r'^(figure|fig)\s*\d+', r'^fig\.?\s*\d+']:
        if re.match(pattern, cleaned, re.IGNORECASE):
            if len(text) < 300 and not any(term in cleaned for term in ('abstract','keyword','reference')):
                return 'figure_caption'

    # 4) Table captions
    for pattern in [r'^(table|tab)\b', r'^table\s*\d+']:
        if re.match(pattern, cleaned, re.IGNORECASE):
            if len(text) < 300:
                return 'table_caption'

    # 5) Section keywords (use cleaned text and word-boundary checks)
    section_keywords = {
        'abstract': ['abstract'],
        'keywords': ['keywords', 'keyword'],
        'introduction': ['introduction'],
        'literature review': ['literature review', 'related work'],
        'research methodology': ['research methodology', 'methodology', 'methods'],
        'results and discussions': ['results and discussions', 'results', 'discussion'],
        'conclusion': ['conclusion', 'conclusions'],
        'acknowledgement': ['acknowledgement', 'acknowledgments', 'acknowledgement'],
        'funding statement': ['funding statement', 'funding'],
        'author contributions': ['author contributions', 'contributions'],
        'conflict of interests': ['conflict of interests', 'conflicts of interest', 'competing interests'],
        'ethics statements': ['ethics statements', 'ethical statement', 'ethics'],
        'references': ['references', 'reference', 'bibliography'],
    }

    # Check explicit starts (cleaned) e.g. "introduction", or exact match, or word-boundary anywhere
    for sect, kws in section_keywords.items():
        for kw in kws:
            if (
                cleaned.startswith(kw)
                or re.match(r'^(?:\d+[.)]?\s*)?'+re.escape(kw)+r'\b', cleaned)
                or re.match(r'^(?:' + re.escape(kw) + r')\s*[:–-]', cleaned)
            ):
                # Ensure it’s short enough to be a heading (< 10 words)
                if sect in ['abstract', 'keywords'] or len(cleaned.split()) <= 15:
                    return sect

    # 6) Numbered headings like "1. Introduction" or "1 Introduction" — cleaned will remove the number,
    #    so if cleaned contains a known section keyword we already caught it. But handle cases where
    #    cleaned is short and matches "introduction" etc.
    if re.match(r'^\d+[\.\)]?\s*\w+', lower_text) and len(text) < 200:
        for sect, kws in section_keywords.items():
            for kw in kws:
                if kw in lower_text:
                    return sect
        return 'main_heading'

    # 7) Author / corresponding hints (early in doc)
    if idx < 6:
        if '@' in raw or 'correspond' in lower_text:
            return 'corresponding_author'
        # name-like line with commas and TitleCase tokens
        if ',' in raw:
            name_like = len(re.findall(r'\b[A-Z][a-z]{1,}\s+[A-Z][a-z]{1,}\b', raw)) >= 1
            if name_like:
                return 'authors'

    # 8) Early short title heuristic (if early and looks like Title Case)
    if idx < 4 and len(text.split()) > 2 and font_size >= 14:
        titlecase_count = sum(1 for w in re.split(r'\s+', text) if w and len(w) > 2 and w[0].isupper() and w[1:].islower())
        if titlecase_count >= 2:
            return 'title'

    return 'body_text'

# -------------------------
# Main Analysis Function
# -------------------------
def analyze_documents(template_file, manuscript_file):
    """Main analysis function that returns findings, missing, and XML previews"""
    # Convert both documents to XML
    template_xml, template_xml_string = docx_to_xml(template_file)
    manuscript_xml, manuscript_xml_string = docx_to_xml(manuscript_file)
    
    if not template_xml or not manuscript_xml:
        return [], ["Error: Could not parse documents"], None
    
    # Get XML previews for display
    template_preview = get_xml_preview(template_xml)
    manuscript_preview = get_xml_preview(manuscript_xml)
    
    # Extract paragraphs from both
    template_paragraphs = extract_paragraphs_from_xml(template_xml)
    manuscript_paragraphs = extract_paragraphs_from_xml(manuscript_xml)
    
    # Analyze template to get formatting rules
    template_rules = analyze_template_formatting(template_paragraphs)
    
    # Compare manuscript against template rules
    findings = compare_against_template(manuscript_paragraphs, template_rules)
    
    # Check for missing sections
    missing = check_missing_sections(manuscript_paragraphs, template_rules)
    
    # Create XML previews object
    xml_previews = {
        'template': template_preview,
        'manuscript': manuscript_preview,
        'template_full': template_xml_string[:5000] if template_xml_string else None,  # First 5000 chars
        'manuscript_full': manuscript_xml_string[:5000] if manuscript_xml_string else None
    }
    
    return findings, missing, xml_previews

def analyze_template_formatting(template_paragraphs):
    """Analyze template to extract formatting rules for each section type"""
    rules = {}
    section_examples = defaultdict(list)
    
    # First pass: classify each paragraph and collect examples
    for para in template_paragraphs:
        section_type = classify_section_type(para)
        section_examples[section_type].append(para)
    
    # Second pass: determine dominant formatting for each section
    for section_type, examples in section_examples.items():
        if examples:
            rules[section_type] = determine_section_formatting(examples)
    
    return rules

def determine_section_formatting(examples):
    """Determine the typical formatting for a section based on examples"""
    if not examples:
        return get_default_formatting('body_text')
    
    # Collect all formatting attributes
    font_sizes = []
    font_names = []
    bold_flags = []
    italic_flags = []
    
    for example in examples:
        if example.get('font_size'):
            font_sizes.append(example['font_size'])
        if example.get('font_name'):
            font_names.append(example['font_name'])
        bold_flags.append(example.get('bold', False))
        italic_flags.append(example.get('italic', False))
    
    # Use most common values
    formatting = {}
    
    if font_sizes:
        formatting['font_size'] = Counter(font_sizes).most_common(1)[0][0]
    
    if font_names:
        formatting['font_name'] = Counter(font_names).most_common(1)[0][0]
    
    # Use majority for boolean flags
    if bold_flags:
        formatting['bold'] = sum(bold_flags) > len(bold_flags) / 2
    
    if italic_flags:
        formatting['italic'] = sum(italic_flags) > len(italic_flags) / 2
    
    return formatting

def get_default_formatting(section_type):
    """Get default formatting rules as fallback"""
    defaults = {
        # === Titles ===
        'title': {'font_size': 24.0, 'bold': True, 'font_name': 'Times New Roman'},

        # === Author Info ===
        'authors': {'font_size': 11.0, 'bold': True, 'font_name': 'Times New Roman'},
        'affiliation': {'font_size': 9.0, 'bold': False, 'font_name': 'Times New Roman'},
        'corresponding_author': {'font_size': 9.0, 'bold': False, 'italic': True, 'font_name': 'Times New Roman'},

        # === Abstract & Keywords ===
        'abstract': {'font_size': 9.0, 'bold': False, 'font_name': 'Times New Roman'},
        'keywords': {'font_size': 9.0, 'bold': False, 'italic': True, 'font_name': 'Times New Roman'},

        # === Headings / Sections ===
        'introduction': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'literature review': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'research methodology': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'results and discussions': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'conclusion': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'acknowledgement': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'funding statement': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'author contributions': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'conflict of interests': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'ethics statements': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'biographies of authors': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},
        'main_heading': {'font_size': 10.0, 'bold': True, 'font_name': 'Times New Roman'},

        # === Other ===
        'body_text': {'font_size': 10.0, 'bold': False, 'font_name': 'Times New Roman'},
        'figure_caption': {'font_size': 10.0, 'bold': False, 'font_name': 'Times New Roman'},
        'table_caption': {'font_size': 10.0, 'bold': False, 'font_name': 'Times New Roman'},
        'references': {'font_size': 9.0, 'bold': False, 'font_name': 'Times New Roman'}
    }

    return defaults.get(section_type, defaults['body_text'])


def compare_against_template(manuscript_paragraphs, template_rules):
    """Compare manuscript paragraphs against template formatting rules"""
    findings = []
    
    for para in manuscript_paragraphs:
        if not para['text'].strip():
            continue
            
        section_type = classify_section_type(para)
        expected = template_rules.get(section_type, template_rules.get('body_text'))
        
        # Skip body text for most checks to reduce false positives
        if section_type == 'body_text':
            continue
        
        # Check formatting mismatches
        findings.extend(check_formatting_mismatches(para, section_type, expected))
    
    return findings

def check_formatting_mismatches(paragraph, section_type, expected):
    """Check for specific formatting mismatches.
       - Only 'Keywords' and 'Corresponding Author' must be italic.
       - Headers (Introduction, Conclusion, etc.) must NOT be bold or italic.
       - Normal text (body_text) → ignore bold/italic completely.
    """
    findings = []

    text_lower = paragraph.get('text', '').strip().lower()
    actual_bold = paragraph.get('bold', False)
    actual_italic = paragraph.get('italic', False)

    # --- 1️⃣ Skip bold/italic check for body text ---
    if section_type == 'body_text':
        # Still check font size / font name only
        if expected.get('font_size') and paragraph.get('font_size'):
            expected_size = expected['font_size']
            actual_size = paragraph['font_size']
            if abs(expected_size - actual_size) > 1.0:
                findings.append(create_finding(
                    section_type, paragraph['index'],
                    'font_size_mismatch',
                    f"{actual_size} pt", f"{expected_size} pt",
                    paragraph['text'], f"Set font size to {expected_size} pt"
                ))
        if expected.get('font_name') and paragraph.get('font_name'):
            if not fonts_similar(expected['font_name'], paragraph['font_name']):
                findings.append(create_finding(
                    section_type, paragraph['index'],
                    'font_family_mismatch',
                    paragraph['font_name'], expected['font_name'],
                    paragraph['text'], f"Set font to '{expected['font_name']}'"
                ))
        return findings  # ✅ No bold/italic checks for body text

    # --- 2️⃣ Bold rules ---
    header_must_not_be_bold = {
        'acknowledgement', 'acknowledgment', 'introduction', 'abstract',
        'conclusion', 'references', 'literature review',
        'research methodology', 'results and discussions',
        'funding statement', 'author contributions', 'conflict of interests',
        'ethics statements', 'biographies of authors', 'main_heading'
    }

    if section_type == 'title':
        expected_bold = True
    elif any(text_lower.startswith(h) for h in header_must_not_be_bold):
        expected_bold = False
    else:
        expected_bold = expected.get('bold', False)

    if expected_bold != actual_bold:
        issue_type = 'bold_missing' if expected_bold and not actual_bold else 'bold_incorrect'
        findings.append(create_finding(
            section_type, paragraph['index'], issue_type,
            'bold' if actual_bold else 'not bold',
            'bold' if expected_bold else 'not bold',
            paragraph['text'], f"Set bold = {expected_bold}"
        ))

    # --- 3️⃣ Italic rules ---
    if section_type in ['keywords', 'corresponding_author']:
        # These MUST be italic
        if not actual_italic:
            findings.append(create_finding(
                section_type, paragraph['index'], 'italic_missing',
                'not italic', 'italic',
                paragraph['text'], "Set to italic"
            ))
    elif section_type in header_must_not_be_bold or section_type in ['title', 'references']:
        # These must NOT be italic
        if actual_italic:
            findings.append(create_finding(
                section_type, paragraph['index'], 'italic_incorrect',
                'italic', 'not italic',
                paragraph['text'], "Remove italic formatting"
            ))

    # --- 4️⃣ Font size & family checks (same as before) ---
    if expected.get('font_size') and paragraph.get('font_size'):
        expected_size = expected['font_size']
        actual_size = paragraph['font_size']
        if abs(expected_size - actual_size) > 1.0:
            findings.append(create_finding(
                section_type, paragraph['index'], 'font_size_mismatch',
                f"{actual_size} pt", f"{expected_size} pt",
                paragraph['text'], f"Set font size to {expected_size} pt"
            ))

    if expected.get('font_name') and paragraph.get('font_name'):
        expected_font = expected['font_name']
        actual_font = paragraph['font_name']
        if not fonts_similar(expected_font, actual_font):
            findings.append(create_finding(
                section_type, paragraph['index'], 'font_family_mismatch',
                actual_font, expected_font,
                paragraph['text'], f"Set font to '{expected_font}'"
            ))

    return findings

def create_finding(section, index, issue_type, found, expected, snippet, fix):
    """Create a standardized finding object"""
    return {
        'type': issue_type,
        'section': section,
        'paragraph_indices': [index],
        'pages': [1 + (index // 5)],  # Rough page estimate
        'found': found,
        'expected': expected,
        'snippet': text_snippet(snippet),
        'suggested_fix': fix
    }

def check_missing_sections(manuscript_paragraphs, template_rules):
    """Check for missing required sections"""
    manuscript_sections = set()
    
    for para in manuscript_paragraphs:
        section_type = classify_section_type(para)
        manuscript_sections.add(section_type)
    
    required_sections = {'title', 'abstract', 'keywords', 'references'}
    missing = required_sections - manuscript_sections
    
    return list(missing)

# -------------------------
# Utility Functions
# -------------------------
def now_timestamp():
    return time.strftime("%Y%m%d_%H%M%S")

def half_points_to_pt(val):
    try:
        return round(float(val) / 2.0, 2)
    except Exception:
        return None

def text_snippet(s, length=140):
    if not s:
        return ""
    s = s.replace("\n", " ").strip()
    return (s[:length] + ("..." if len(s) > length else ""))

def normalize_font_name(name):
    """Simple font normalization"""
    if not name:
        return ""
    
    name = name.lower().strip()
    
    # Basic font aliases
    aliases = {
        'timesnewroman': 'Times New Roman',
        'times roman': 'Times New Roman',
        'times': 'Times New Roman',
        'arial': 'Arial',
        'helvetica': 'Arial',
        'calibri': 'Calibri',
        'cambria': 'Cambria'
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
        {'times new roman', 'times', 'timesroman'},
        {'arial', 'helvetica'},
        {'calibri', 'cambria'}
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
    
    headers = ["No", "IssueType", "Section", "ParagraphIndices", "PageEstimates", "SnippetExamples", "Found", "Expected", "SuggestedFix"]
    ws.append(headers)
    
    for i, f in enumerate(findings, start=1):
        ws.append([
            i,
            f.get('type'),
            f.get('section'),
            ", ".join(str(x) for x in f.get('paragraph_indices', [])),
            ", ".join(str(x) for x in f.get('pages', [])),
            f.get('snippet'),
            f.get('found'),
            f.get('expected'),
            f.get('suggested_fix')
        ])
    
    if missing_sections:
        ws2 = wb.create_sheet("MissingSections")
        ws2.append(["Missing sections:"])
        for s in missing_sections:
            ws2.append([s])
    
    wb.save(out_path)
    return out_path

# -------------------------
# Auto-Correction Functions
# -------------------------
def apply_corrections(template_file, manuscript_file, mistakes_df):
    """Apply corrections using python-docx"""
    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError:
        raise ImportError("python-docx is required for auto-correction")
    
    try:
        # Create temporary file for manuscript
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_manuscript:
            manuscript_file.seek(0)
            tmp_manuscript.write(manuscript_file.read())
            manuscript_path = tmp_manuscript.name

        # Load the document
        doc = Document(manuscript_path)
        
        # Group mistakes by paragraph index
        mistakes_by_paragraph = defaultdict(list)
        for _, mistake in mistakes_df.iterrows():
            para_indices = mistake['paragraph_indices']
            if isinstance(para_indices, list):
                for idx in para_indices:
                    if idx < len(doc.paragraphs):
                        mistakes_by_paragraph[idx].append(mistake)
        
        # Apply corrections to each paragraph
        corrections_applied = 0
        for para_idx, mistakes in mistakes_by_paragraph.items():
            if para_idx < len(doc.paragraphs):
                paragraph = doc.paragraphs[para_idx]
                
                for mistake in mistakes:
                    if apply_single_correction(paragraph, mistake):
                        corrections_applied += 1
        
        # Save corrected document
        corrected_bytes = io.BytesIO()
        doc.save(corrected_bytes)
        corrected_bytes.seek(0)
        
        os.unlink(manuscript_path)
        
        print(f"🔧 Applied {corrections_applied} corrections")
        return corrected_bytes.getvalue()
        
    except Exception as e:
        print(f"Error applying corrections: {str(e)}")
        try:
            os.unlink(manuscript_path)
        except:
            pass
        return None

def apply_single_correction(paragraph, mistake):
    """Apply a single formatting correction"""
    try:
        from docx.shared import Pt
        
        correction_type = mistake['type']
        expected = mistake['expected']
        
        if correction_type == 'font_size_mismatch':
            size_match = re.search(r'(\d+\.?\d*)', expected)
            if size_match:
                font_size = float(size_match.group(1))
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
                return True
                
        elif correction_type in ['bold_missing', 'bold_incorrect', 'bold_mismatch']:
            bold_value = 'bold' in expected.lower()
            for run in paragraph.runs:
                run.font.bold = bold_value
            return True
            
        elif correction_type == 'font_family_mismatch':
            font_name = expected
            for run in paragraph.runs:
                run.font.name = font_name
            return True
            
    except Exception as e:
        print(f"Could not apply correction: {mistake['type']} - {str(e)}")
        return False
    
    return False

def highlight_mistakes(template_file, manuscript_file, mistakes_df):
    """Highlight mistakes in the manuscript"""
    try:
        from docx import Document
    except ImportError:
        raise ImportError("python-docx required for highlighting")
    
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_manuscript:
            manuscript_file.seek(0)
            tmp_manuscript.write(manuscript_file.read())
            manuscript_path = tmp_manuscript.name

        doc = Document(manuscript_path)
        
        # Collect paragraphs to highlight
        paragraphs_to_highlight = set()
        for _, mistake in mistakes_df.iterrows():
            para_indices = mistake['paragraph_indices']
            if isinstance(para_indices, list):
                for idx in para_indices:
                    if idx < len(doc.paragraphs):
                        paragraphs_to_highlight.add(idx)
        
        # Apply highlighting
        for para_idx in paragraphs_to_highlight:
            paragraph = doc.paragraphs[para_idx]
            for run in paragraph.runs:
                run.font.highlight_color = 7  # Yellow
        
        # Save highlighted document
        highlighted_bytes = io.BytesIO()
        doc.save(highlighted_bytes)
        highlighted_bytes.seek(0)
        
        os.unlink(manuscript_path)
        return highlighted_bytes.getvalue()
        
    except Exception as e:
        print(f"Error highlighting: {str(e)}")
        return None

# -------------------------
# CLI Main
# -------------------------
def main(argv):
    parser = argparse.ArgumentParser(description="JIWE XML Formatter - XML-based document analysis")
    parser.add_argument('template', help='template.docx')
    parser.add_argument('manuscript', help='manuscript.docx')
    parser.add_argument('--output', '-o', default=None, help='output Excel path')
    args = parser.parse_args(argv[1:])

    template = args.template
    manuscript = args.manuscript

    if not os.path.exists(template):
        print("ERROR: Template not found:", template); return 2
    if not os.path.exists(manuscript):
        print("ERROR: Manuscript not found:", manuscript); return 3

    print("=== JIWE XML FORMatter ===")
    print("Approach: Convert both documents to XML and compare at XML level")
    
    findings, missing, xml_previews = analyze_documents(template, manuscript)
    
    print("\n=== TEMPLATE XML PREVIEW ===")
    print(xml_previews['template'] if xml_previews else "No preview available")
    
    print("\n=== MANUSCRIPT XML PREVIEW ===")
    print(xml_previews['manuscript'] if xml_previews else "No preview available")
    
    out_path = args.output or f"xml_analysis_{now_timestamp()}.xlsx"
    saved = export_findings_to_excel(findings, missing, out_path=out_path)
    
    print(f"\n✅ XML ANALYSIS REPORT: {saved}")
    print(f"📊 Issues found: {len(findings)}")
    print(f"📋 Missing sections: {len(missing)}")
    
    if findings:
        print("\n🔍 Top issues:")
        for i, f in enumerate(findings[:10], start=1):
            print(f"  {i}. [{f['type']}] {f['section']} - para {f['paragraph_indices']}")
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