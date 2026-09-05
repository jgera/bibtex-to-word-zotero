# BibTeX to Zotero Word (.docx) Reference Converter & Antigravity Skill

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python Version](https://img.shields.io/badge/Python-3.8%2B-green.svg)](pyproject.toml)
[![Antigravity Skill](https://img.shields.io/badge/Antigravity-Skill-purple.svg)](SKILL.md)

A zero-dependency Python tool and **Antigravity AI Agent Skill** that automatically parses BibTeX (`.bib` / `.txt`) files and embeds **native Zotero field codes** (`ADDIN ZOTERO_ITEM CSL_CITATION`) directly into Microsoft Word (`.docx`) documents.

Generated `.docx` documents open seamlessly in Microsoft Word **without any recovery prompts**, allowing researchers to instantly click **Refresh** or **Add/Edit Bibliography** using Zotero's desktop plugin.

---

## Key Features

- **Native Zotero OpenXML Integration**: Injects full 5-run OpenXML field sequences (`w:fldChar` begin -> `w:instrText` CSL-JSON -> `w:fldChar` separate -> `w:t` label -> `w:fldChar` end).
- **Zero External Dependencies**: Operates strictly using Python standard libraries (`zipfile`, `xml.etree.ElementTree`, `re`, `json`).
- **Depth-Balanced BibTeX Parser**: Correctly extracts fields containing arbitrary nested curly braces `{...}` without dropping authors or metadata.
- **LaTeX Accent Sanitization**: Converts LaTeX accent codes (`{\L}`, `{\'e}`, `{\`e}`, `\"o`) to UTF-8 Unicode (`Ł`, `é`, `è`, `ö`) and removes backslashes to avoid Microsoft Word field switch conflicts (`\f`, `\h`).
- **OpenXML 255-Character Limit Resolution**: Automatically chunks Zotero document settings across `ZOTERO_PREF_1` (`pid=2`), `ZOTERO_PREF_2` (`pid=3`) in `docProps/custom.xml`.
- **Dynamic Relationship ID Generation**: Computes unique `rIdN` suffixes in `_rels/.rels` to eliminate relationship ID collisions.
- **Namespace & Header Preservation**: Retains all original `<w:document>` root element attributes (`xmlns:w15`, `xmlns:mc`, `mc:Ignorable`).

---

## Installation

### Method 1: Use as an Antigravity AI Skill
To give any Antigravity AI coding agent instant ability to convert BibTeX files to Zotero-linked Word documents:

- **Workspace Skill (Project-specific)**:
  Copy this folder to `.agents/skills/bibtex-zotero` in your project root.
- **Global Skill (System-wide)**:
  Copy this folder to `~/.gemini/config/skills/bibtex-zotero/`.

### Method 2: Standalone CLI Tool
Clone the repository and run directly:
```bash
git clone https://github.com/your-username/bibtex-zotero.git
cd bibtex-zotero
```

---

## Quick Start / Usage

### Basic CLI Command
```bash
python scripts/bibtex_to_zotero_word.py --input input_paper.docx --bib references.bib --output output_zotero.docx
```

### Windows Batch Runner
Run all test samples at once:
```cmd
scripts\run_samples.cmd
```
or run a specific sample:
```cmd
scripts\run_samples.cmd 1
```

---

## Document & Field Code Structure

### 1. In-Text Citation OpenXML Structure
```xml
<!-- 1. Begin Field Code -->
<w:r><w:fldChar w:fldCharType="begin"/></w:r>

<!-- 2. CSL-JSON Instruction Code -->
<w:r>
  <w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"3ntzNmjI","properties":{"formattedCitation":"[1]","plainCitation":"[1]"},"citationItems":[{"id":10001,"uris":["http://zotero.org/users/local/items/ZMrbCCKx"],"itemData":{"type":"article-journal","title":"...","author":[{"family":"Wang","given":"Jianfeng"}]}}]} </w:instrText>
</w:r>

<!-- 3. Separate Instruction from Result -->
<w:r><w:fldChar w:fldCharType="separate"/></w:r>

<!-- 4. Formatted Display Label -->
<w:r><w:t>[1]</w:t></w:r>

<!-- 5. End Field Code -->
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

### 2. Custom Properties (`docProps/custom.xml`)
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ZOTERO_PREF_1">
    <vt:lpwstr>&lt;data data-version="3" zotero-version="7.0.0"&gt;&lt;session id="slBw3MX2"/&gt;&lt;style id="http://www.zotero.org/styles/ieee" locale="en-US" hasBibliography="1" bibliographyStyleHasBeenSet="0"/&gt;&lt;prefs&gt;&lt;pref name="fieldType" value="Field"/&gt;&lt;pref name="automaticJourn</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="ZOTERO_PREF_2">
    <vt:lpwstr>alAbbreviations" value="true"/&gt;&lt;/prefs&gt;&lt;/data&gt;</vt:lpwstr>
  </property>
</Properties>
```

---

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests to improve CSL item mapping, bibtex field support, or citation styles.

## License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.
