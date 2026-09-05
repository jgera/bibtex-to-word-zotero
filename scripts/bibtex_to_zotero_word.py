#!/usr/bin/env python3
"""
BibTeX to Zotero Word (.docx) Reference Inserter
Reads a .docx document containing {citekey} placeholders and a BibTeX (.txt / .bib) file,
and outputs a Word document with embedded native Zotero citation fields.
"""

import os
import re
import json
import random
import string
import zipfile
import argparse
import xml.etree.ElementTree as ET

# Namespaces for Word OpenXML
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace('w', WORD_NS)
ET.register_namespace('w14', "http://schemas.openxmlformats.org/wordprocessingml/2010/main")
ET.register_namespace('r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

def generate_random_id(length=8):
    """Generate random alphanumeric ID for Zotero citation keys."""
    chars = string.ascii_letters + string.digits
    return ''.join(random.choices(chars, k=length))

def clean_latex_accents(text):
    """Convert LaTeX accent escape sequences to clean unicode characters."""
    if not text:
        return ""
    # Map common LaTeX accent patterns
    replacements = [
        (r'{\L}', 'Ł'), (r'{\l}', 'ł'), (r'\L', 'Ł'), (r'\l', 'ł'),
        (r"{\'e}", 'é'), (r"{\`e}", 'è'), (r"{\^e}", 'ê'), (r'{\"e}', 'ë'),
        (r"{\'a}", 'á'), (r"{\`a}", 'à'), (r"{\^a}", 'â'), (r'{\"a}', 'ä'),
        (r"{\'o}", 'ó'), (r"{\`o}", 'ò'), (r"{\^o}", 'ô'), (r'{\"o}', 'ö'),
        (r"{\'u}", 'ú'), (r"{\`u}", 'ù'), (r"{\^u}", 'û'), (r'{\"u}', 'ü'),
        (r"{\'i}", 'í'), (r"{\`i}", 'ì'), (r"{\^i}", 'î'), (r'{\"i}', 'ï'),
        (r"\'e", 'é'), (r"\`e", 'è'), (r"\^e", 'ê'), (r'\"e', 'ë'),
        (r"\'a", 'á'), (r"\`a", 'à'), (r"\^a", 'â'), (r'\"a', 'ä'),
        (r"\'o", 'ó'), (r"\`o", 'ò'), (r"\^o", 'ô'), (r'\"o', 'ö'),
        (r"\'u", 'ú'), (r"\`u", 'ù'), (r"\^u", 'û'), (r'\"u', 'ü'),
        (r"\'i", 'í'), (r"\`i", 'ì'), (r"\^i", 'î'), (r'\"i', 'ï'),
        (r'\c{c}', 'ç'), (r'{\c c}', 'ç'), (r'\~n', 'ñ'), (r'{\~n}', 'ñ'),
        (r'\u{g}', 'ğ'), (r'{\u g}', 'ğ'), (r'\v{s}', 'š'), (r'{\v s}', 'š'),
        (r'\"', '"'), (r'\{', ''), (r'\}', ''), (r'\\', '')
    ]
    for pattern, repl in replacements:
        text = text.replace(pattern, repl)
    # Remove any remaining standalone backslashes to avoid Word field switch conflicts
    text = text.replace('\\', '')
    return text

def parse_authors(author_str):
    """Parse BibTeX author string into CSL author list."""
    authors = []
    if not author_str:
        return authors
    
    # Split multiple authors
    raw_authors = author_str.split(' and ')
    for a in raw_authors:
        a = clean_latex_accents(a.strip())
        if not a:
            continue
        if ',' in a:
            parts = a.split(',', 1)
            family = parts[0].strip()
            given = parts[1].strip()
        else:
            parts = a.rsplit(' ', 1)
            if len(parts) == 2:
                given = parts[0].strip()
                family = parts[1].strip()
            else:
                family = a
                given = ""
                
        # Filter out 'others' or 'et al.' entries
        if family.lower() in ('others', 'et al.', 'et al'):
            continue
            
        authors.append({'family': family, 'given': given})
    return authors

def clean_bib_value(val):
    """Clean LaTeX and BibTeX formatting from string value."""
    if not val:
        return ""
    val = val.strip()
    # Strip outer braces or quotes
    if (val.startswith('{') and val.endswith('}')) or (val.startswith('"') and val.endswith('"')):
        val = val[1:-1].strip()
    val = clean_latex_accents(val)
    val = re.sub(r'\s+', ' ', val)
    # Common LaTeX cleanups
    val = val.replace('--', '–')
    val = val.replace('``', '"').replace("''", '"')
    return val

def parse_bibtex_fields(body):
    """Parse key=value pairs inside a BibTeX entry body, handling arbitrary nested braces."""
    fields = {}
    i = 0
    n = len(body)
    while i < n:
        m = re.search(r'([a-zA-Z0-9_\-]+)\s*=\s*', body[i:])
        if not m:
            break
        key = m.group(1).lower()
        val_start = i + m.end()
        if val_start >= n:
            break
        
        char = body[val_start]
        val = ""
        if char == '{':
            depth = 1
            j = val_start + 1
            while j < n and depth > 0:
                if body[j] == '{':
                    depth += 1
                elif body[j] == '}':
                    depth -= 1
                j += 1
            val = body[val_start+1:j-1]
            i = j
        elif char == '"':
            j = val_start + 1
            while j < n and body[j] != '"':
                if body[j] == '\\' and j + 1 < n:
                    j += 2
                else:
                    j += 1
            val = body[val_start+1:j]
            i = j + 1
        else:
            m_end = re.search(r'[,}\n]', body[val_start:])
            if m_end:
                val = body[val_start:val_start+m_end.start()]
                i = val_start + m_end.start()
            else:
                val = body[val_start:]
                i = n
                
        fields[key] = clean_bib_value(val)
    return fields

def parse_bibtex(bib_content):
    """Parse BibTeX file content into dictionary of entries."""
    entries = {}
    # Pattern to match entry type, key, and body
    entry_pattern = re.compile(r'@(\w+)\s*\{\s*([^,\s]+)\s*,\s*(.*?)(?=\n@|\s*$)', re.DOTALL)
    
    for match in entry_pattern.finditer(bib_content):
        entry_type = match.group(1).lower()
        cite_key = match.group(2).strip()
        body = match.group(3)
        
        fields = parse_bibtex_fields(body)
            
        # Map entry type to CSL type
        csl_type_map = {
            'article': 'article-journal',
            'inproceedings': 'paper-conference',
            'conference': 'paper-conference',
            'book': 'book',
            'incollection': 'chapter',
            'phdthesis': 'thesis',
            'mastersthesis': 'thesis',
            'techreport': 'report',
            'misc': 'article'
        }
        csl_type = csl_type_map.get(entry_type, 'article-journal')
        
        entries[cite_key] = {
            'type': csl_type,
            'raw_type': entry_type,
            'fields': fields
        }
        
    return entries

def bib_entry_to_csl_item_data(item_id, entry):
    """Convert parsed BibTeX entry to CSL itemData object."""
    fields = entry['fields']
    csl_type = entry['type']
    
    item_data = {
        'id': item_id,
        'type': csl_type,
        'title': fields.get('title', ''),
    }
    
    # Authors
    if 'author' in fields:
        item_data['author'] = parse_authors(fields['author'])
        
    # Container title (journal / booktitle / organization)
    if 'journal' in fields:
        item_data['container-title'] = fields['journal']
    elif 'booktitle' in fields:
        item_data['container-title'] = fields['booktitle']
        
    # Issued date (year)
    if 'year' in fields:
        year_str = fields['year']
        year_match = re.search(r'\d{4}', year_str)
        if year_match:
            item_data['issued'] = {'date-parts': [[year_match.group(0)]]}
            
    # Page numbers
    if 'pages' in fields:
        item_data['page'] = fields['pages']
        
    # Publisher / Organization
    if 'publisher' in fields:
        item_data['publisher'] = fields['publisher']
    elif 'organization' in fields and csl_type == 'paper-conference':
        item_data['publisher'] = fields['organization']
        
    # Volume / Issue
    if 'volume' in fields:
        item_data['volume'] = fields['volume']
    if 'number' in fields or 'issue' in fields:
        item_data['issue'] = fields.get('number') or fields.get('issue')
        
    # DOI / URL
    if 'doi' in fields:
        item_data['DOI'] = fields['doi']
    if 'url' in fields:
        item_data['URL'] = fields['url']
        
    return item_data

def generate_zotero_csl_citation_json(cite_key, bib_entry, citation_index):
    """Build the JSON string for ADDIN ZOTERO_ITEM CSL_CITATION."""
    citation_id = generate_random_id(8)
    zotero_item_key = generate_random_id(8)
    item_id = 10000 + citation_index
    
    formatted_label = f"[{citation_index}]"
    
    item_data = bib_entry_to_csl_item_data(item_id, bib_entry)
    
    csl_obj = {
        "citationID": citation_id,
        "properties": {
            "unsorted": False,
            "formattedCitation": formatted_label,
            "plainCitation": formatted_label,
            "noteIndex": 0
        },
        "citationItems": [
            {
                "id": item_id,
                "uris": [
                    f"http://zotero.org/users/local/items/{zotero_item_key}"
                ],
                "itemData": item_data
            }
        ],
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
    }
    
    return json.dumps(csl_obj, separators=(',', ':')), formatted_label

def register_all_namespaces(xml_bytes):
    """Automatically register all XML namespace prefixes found in the document header."""
    header = xml_bytes[:4000].decode('utf-8', errors='ignore')
    # Match xmlns:prefix="uri"
    matches = re.findall(r'xmlns:([a-zA-Z0-9_\-]+)=["\']([^"\']+)["\']', header)
    for prefix, uri in matches:
        ET.register_namespace(prefix, uri)
    # Also register default namespace
    default_match = re.search(r'xmlns=["\']([^"\']+)["\']', header)
    if default_match:
        ET.register_namespace('', default_match.group(1))

def process_document_xml(xml_bytes, bib_entries):
    """Process word/document.xml and replace {citekey} placeholders with Zotero fields."""
    register_all_namespaces(xml_bytes)
    
    # Store exact original opening <w:document ...> tag (all text up to <w:body>) to preserve mc:Ignorable & all xmlns declarations
    xml_str_orig = xml_bytes.decode('utf-8')
    body_pos = xml_str_orig.find('<w:body>')
    orig_header = xml_str_orig[:body_pos] if body_pos != -1 else ""
    
    root = ET.fromstring(xml_bytes)
    
    w_p = f"{{{WORD_NS}}}p"
    w_r = f"{{{WORD_NS}}}r"
    w_t = f"{{{WORD_NS}}}t"
    w_fldChar = f"{{{WORD_NS}}}fldChar"
    w_instrText = f"{{{WORD_NS}}}instrText"
    w_rPr = f"{{{WORD_NS}}}rPr"
    w_rFonts = f"{{{WORD_NS}}}rFonts"
    
    citation_counter = 1
    citations_replaced = 0
    
    # Iterate through all paragraph (<w:p>) elements
    for p in root.iter(w_p):
        # Extract full text of paragraph to check for placeholders
        text_runs = []
        rPr_sample = None
        
        for elem in p:
            if elem.tag == w_r:
                if rPr_sample is None:
                    rPr_elem = elem.find(w_rPr)
                    if rPr_elem is not None:
                        rPr_sample = rPr_elem
                t_elem = elem.find(w_t)
                if t_elem is not None and t_elem.text:
                    text_runs.append(t_elem.text)
                    
        full_text = "".join(text_runs)
        
        # Check if any {citekey} placeholder is in full_text
        placeholders_in_p = []
        for key in bib_entries.keys():
            placeholder = f"{{{key}}}"
            if placeholder in full_text:
                placeholders_in_p.append((placeholder, key))
                
        if not placeholders_in_p:
            continue
            
        # Rebuild paragraph contents to replace placeholders with field elements
        pPr = p.find(f"{{{WORD_NS}}}pPr")
        p.clear()
        if pPr is not None:
            p.append(pPr)
            
        # Reconstruct elements run by run
        pattern = "|".join(re.escape(ph[0]) for ph in placeholders_in_p)
        splits = re.split(f"({pattern})", full_text)
        
        for segment in splits:
            if not segment:
                continue
                
            # Check if segment matches a placeholder
            matched_key = None
            for ph, key in placeholders_in_p:
                if segment == ph:
                    matched_key = key
                    break
                    
            if matched_key:
                bib_entry = bib_entries[matched_key]
                csl_json_str, formatted_label = generate_zotero_csl_citation_json(
                    matched_key, bib_entry, citation_counter
                )
                citation_counter += 1
                citations_replaced += 1
                
                # Helper to append run with optional rPr
                def make_run():
                    r_el = ET.Element(w_r)
                    if rPr_sample is not None:
                        r_el.append(ET.fromstring(ET.tostring(rPr_sample)))
                    return r_el
                
                # 1. fldChar begin
                r_begin = make_run()
                fld_begin = ET.SubElement(r_begin, w_fldChar)
                fld_begin.set(f"{{{WORD_NS}}}fldCharType", "begin")
                p.append(r_begin)
                
                # 2. instrText
                r_instr = make_run()
                instr_text = ET.SubElement(r_instr, w_instrText)
                instr_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                instr_text.text = f" ADDIN ZOTERO_ITEM CSL_CITATION {csl_json_str} "
                p.append(r_instr)
                
                # 3. fldChar separate
                r_sep = make_run()
                fld_sep = ET.SubElement(r_sep, w_fldChar)
                fld_sep.set(f"{{{WORD_NS}}}fldCharType", "separate")
                p.append(r_sep)
                
                # 4. formatted label run
                r_label = ET.Element(w_r)
                rPr_label = ET.SubElement(r_label, w_rPr)
                rFonts = ET.SubElement(rPr_label, w_rFonts)
                rFonts.set(f"{{{WORD_NS}}}cs", "Times New Roman")
                t_label = ET.SubElement(r_label, w_t)
                t_label.text = formatted_label
                p.append(r_label)
                
                # 5. fldChar end
                r_end = make_run()
                fld_end = ET.SubElement(r_end, w_fldChar)
                fld_end.set(f"{{{WORD_NS}}}fldCharType", "end")
                p.append(r_end)
            else:
                # Normal text segment run
                r_text = ET.Element(w_r)
                if rPr_sample is not None:
                    r_text.append(ET.fromstring(ET.tostring(rPr_sample)))
                t_text = ET.SubElement(r_text, w_t)
                t_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t_text.text = segment
                p.append(r_text)
                
    modified_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True).decode('utf-8')
    
    # Restore original <w:document ...> header tag up to <w:body> so no mc:Ignorable namespace declarations are lost
    first_body = modified_xml.find('<w:body>')
    if orig_header and first_body != -1:
        modified_xml = orig_header + modified_xml[first_body:]
        
    return modified_xml.encode('utf-8'), citations_replaced

def build_custom_xml(style_id="http://www.zotero.org/styles/ieee"):
    """Build docProps/custom.xml containing Zotero document settings, chunked into <=255 char properties."""
    zotero_pref_content = (
        f'<data data-version="3" zotero-version="7.0.0">'
        f'<session id="{generate_random_id(8)}"/>'
        f'<style id="{style_id}" locale="en-US" hasBibliography="1" bibliographyStyleHasBeenSet="0"/>'
        f'<prefs><pref name="fieldType" value="Field"/><pref name="automaticJournalAbbreviations" value="true"/></prefs>'
        f'</data>'
    )
    
    # Split into chunks of max 255 chars for Word custom properties limit
    chunks = [zotero_pref_content[i:i+255] for i in range(0, len(zotero_pref_content), 255)]
    
    props_xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    props_xml.append('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">')
    
    for idx, chunk in enumerate(chunks):
        pid = idx + 2
        name = f"ZOTERO_PREF_{idx + 1}"
        escaped_chunk = chunk.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        props_xml.append(f'  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="{pid}" name="{name}">')
        props_xml.append(f'    <vt:lpwstr>{escaped_chunk}</vt:lpwstr>')
        props_xml.append('  </property>')
        
    props_xml.append('</Properties>')
    return "\n".join(props_xml).encode('utf-8')

def update_content_types(xml_bytes):
    """Ensure docProps/custom.xml override is present in [Content_Types].xml cleanly without namespace changes."""
    xml_str = xml_bytes.decode('utf-8')
    if 'PartName="/docProps/custom.xml"' not in xml_str:
        override = '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
        xml_str = xml_str.replace('</Types>', f'{override}</Types>')
    return xml_str.encode('utf-8')

def update_rels(xml_bytes):
    """Ensure custom-properties relationship is present in _rels/.rels cleanly with a unique rId."""
    xml_str = xml_bytes.decode('utf-8')
    if 'Target="docProps/custom.xml"' not in xml_str:
        # Extract all existing rId numeric suffixes to find next available ID
        r_ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', xml_str)]
        next_id = max(r_ids) + 1 if r_ids else 99
        new_r_id = f"rId{next_id}"
        rel = f'<Relationship Id="{new_r_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
        xml_str = xml_str.replace('</Relationships>', f'{rel}</Relationships>')
    return xml_str.encode('utf-8')

def convert_docx(input_docx, bib_file, output_docx):
    """Main function to insert Zotero references into Word .docx file."""
    if not os.path.exists(input_docx):
        raise FileNotFoundError(f"Input Word document not found: {input_docx}")
    if not os.path.exists(bib_file):
        raise FileNotFoundError(f"BibTeX file not found: {bib_file}")
        
    with open(bib_file, 'r', encoding='utf-8', errors='ignore') as f:
        bib_content = f.read()
        
    bib_entries = parse_bibtex(bib_content)
    print(f"Loaded {len(bib_entries)} entries from BibTeX file: {list(bib_entries.keys())}")
    
    with zipfile.ZipFile(input_docx, 'r') as z_in:
        file_map = {name: z_in.read(name) for name in z_in.namelist()}
        
    if 'word/document.xml' not in file_map:
        raise ValueError("Invalid .docx archive: word/document.xml missing")
        
    doc_xml_bytes = file_map['word/document.xml']
    modified_doc_xml, count = process_document_xml(doc_xml_bytes, bib_entries)
    file_map['word/document.xml'] = modified_doc_xml
    
    # Add docProps/custom.xml and update relationships & content types
    file_map['docProps/custom.xml'] = build_custom_xml()
    if '[Content_Types].xml' in file_map:
        file_map['[Content_Types].xml'] = update_content_types(file_map['[Content_Types].xml'])
    if '_rels/.rels' in file_map:
        file_map['_rels/.rels'] = update_rels(file_map['_rels/.rels'])
        
    # Write to output docx
    try:
        with zipfile.ZipFile(output_docx, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
            for fname, data in file_map.items():
                z_out.writestr(fname, data)
    except PermissionError:
        print(f"\n[ERROR] Permission denied when writing to '{output_docx}'.")
        print(f"[HINT] If the file is currently open in Microsoft Word, please close it and re-run.")
        return
            
    print(f"Successfully generated '{output_docx}' with {count} Zotero citations inserted.")

def main():
    parser = argparse.ArgumentParser(description="Insert native Zotero citations from BibTeX into Word docx")
    parser.add_argument('--input', '-i', default='MiniLM.docx', help='Input Word .docx document containing {citekey}')
    parser.add_argument('--bib', '-b', default='MiniLM.txt', help='Input BibTeX file (.txt or .bib)')
    parser.add_argument('--output', '-o', default='MiniLM_Generated_Zotero.docx', help='Output Word .docx document path')
    
    args = parser.parse_args()
    convert_docx(args.input, args.bib, args.output)

if __name__ == '__main__':
    main()
