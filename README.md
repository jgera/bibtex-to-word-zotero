# BibTeX to Word Zotero (`bibtex-to-word-zotero`)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python Version](https://img.shields.io/badge/Python-3.8%2B-green.svg)](pyproject.toml)
[![Antigravity AI Skill](https://img.shields.io/badge/AI%20Skill-bibtex--to--word--zotero-purple.svg)](SKILL.md)
[![OpenXML Compliant](https://img.shields.io/badge/OpenXML-100%25%20Valid-brightgreen.svg)](references/OOXML_Zotero_Spec.md)

Automate inserting native Zotero reference fields into Microsoft Word (`.docx`) documents directly from BibTeX (`.bib` / `.txt`) files. Zero dependencies, OpenXML compliant, and AI Agent Skill ready.

---

## How It Works

```mermaid
flowchart LR
    INPUT["Word Document ({citekey})<br/>+ BibTeX File"] --> ENGINE["bibtex-to-word-zotero<br/>Engine"]
    ENGINE --> OUTPUT["Zotero-Linked<br/>Word Document (.docx)"]
    OUTPUT --> WORD["Open in Word<br/>(Zotero Add-in Recognizes Citations)"]
```

`bibtex-to-word-zotero` takes a Microsoft Word (`.docx`) file containing citation placeholders (such as `{wang2020minivlm}` or `{vaswani2017attention}`) and a BibTeX file (`.bib` or `.txt`), and automatically injects native Zotero field codes (`ADDIN ZOTERO_ITEM CSL_CITATION`).

The resulting `.docx` document opens in Microsoft Word without any recovery warnings, allowing researchers to click **Refresh** or **Add/Edit Bibliography** instantly using Zotero's desktop plugin.

---

## Features & Capabilities

- **Native Zotero OpenXML Integration**: Generates native 5-run OpenXML field sequences (`w:fldChar` begin → `w:instrText` CSL-JSON → `w:fldChar` separate → `w:t` label → `w:fldChar` end).
- **Multiple Citation Styles (`--style`)**: Supports `ieee`, `apa`, `chicago-author-date`, `mla`, `nature`, `vancouver`, `harvard`. Automatically formats in-text labels (`[1]` vs `(Vaswani & Shazeer, 2017)`).
- **Automated Zotero Bibliography Field (`--add-bibliography`)**: Embeds `ADDIN ZOTERO_BIBL CSL_BIBLIOGRAPHY` fields at `{bibliography}` placeholders or document end.
- **Zero External Dependencies**: Standard Python library implementation (`zipfile`, `xml.etree.ElementTree`, `re`, `json`). No Office or Zotero installation needed for conversion.
- **Robust LaTeX Sanitization**: Converts LaTeX accent codes (`{\L}`, `{\'e}`, `{\`e}`, `\"o`) into clean UTF-8 Unicode (`Ł`, `é`, `è`, `ö`) and removes backslashes to prevent Microsoft Word field switch conflicts (`\f`, `\h`).
- **Depth-Balanced Brace Parsing**: Safely extracts BibTeX values containing nested curly braces `{...}` without dropping author arrays.
- **OpenXML 255-Character Limit Resolution**: Automatically chunks `ZOTERO_PREF_` custom properties to comply with OpenXML string length limits.
- **Dynamic Relationship ID Allocation**: Assigns unique `rIdN` suffixes in `_rels/.rels` to prevent relationship ID collisions.
- **AI Agent Skill Ready (`--json`)**: Includes `SKILL.md` and `--json` structured output contracts for AI Coding Assistants (Gemini, Claude, GPT-4, LLaMA).

---

## Quick Start

### 1. Installation

#### As an AI Agent Skill (Antigravity / AGY)
- **Workspace Skill**: Copy repository to `.agents/skills/bibtex-to-word-zotero` in your project root.
- **Global Skill**: Copy repository to `~/.gemini/config/skills/bibtex-to-word-zotero/`.

#### As a Standalone CLI Tool
```bash
git clone https://github.com/jgera/bibtex-to-word-zotero.git
cd bibtex-to-word-zotero
```

---

### 2. Command Line Usage

#### IEEE Style Conversion (Numeric)
```bash
python scripts/bibtex_to_zotero_word.py --input document.docx --bib references.bib --output document_zotero.docx
```

#### APA Style Conversion with Automated Bibliography
```bash
python scripts/bibtex_to_zotero_word.py --input paper.docx --bib references.bib --output paper_zotero.docx --style apa --add-bibliography
```

#### AI Agent JSON Output Mode
```bash
python scripts/bibtex_to_zotero_word.py --input paper.docx --bib refs.bib --output result.docx --json
```

#### Diagnostic Health Check Mode
```bash
python scripts/bibtex_to_zotero_word.py --input document_zotero.docx --check --json
```

---

## Testing

Run the automated zero-dependency unit test suite:

```bash
python run_tests.py
```

Expected Output:
```text
test_clean_latex_accents ... ok
test_custom_xml_property_chunking ... ok
test_format_citation_label ... ok
test_parse_authors ... ok
test_parse_bibtex_depth_balanced ... ok
test_update_rels_unique_id ... ok

----------------------------------------------------------------------
Ran 6 tests in 0.011s

OK
[SUCCESS] All unit tests passed cleanly!
```

---

## Internal Engine Pipeline & Architecture

```mermaid
flowchart TD
    subgraph Inputs["1. Input Files"]
        DOCX["Word Document (.docx)<br/>with {citekey} placeholders"]
        BIB["BibTeX File (.bib / .txt)<br/>with reference metadata"]
    end

    subgraph Engine["2. bibtex-to-word-zotero Engine"]
        CORE["Zero-Dependency OpenXML Converter"]
        PARSE["1. Depth-Balanced BibTeX Parser"]
        SAN["2. LaTeX Accent & Special Character Sanitizer"]
        CHUNK["3. ZOTERO_PREF 255-Char Property Chunker"]
        RID["4. Unique Relationship ID Allocator (rIdN)"]

        CORE --> PARSE --> SAN --> CHUNK --> RID
    end

    subgraph Output["3. Output Result"]
        OUT_DOCX["Converted Word Document (.docx)<br/>with Native Zotero Field Codes"]
    end

    subgraph Desktop["4. Microsoft Word & Zotero"]
        WORD_UI["Microsoft Word"]
        ZOTERO_UI["Zotero Add-in Instant Recognition & Bibliography Rendering"]

        WORD_UI --> ZOTERO_UI
    end

    DOCX --> CORE
    BIB --> CORE
    RID --> OUT_DOCX
    OUT_DOCX --> WORD_UI
```

For in-depth reverse-engineering details, OpenXML field schemas, custom property chunking rules, and relationship ID handling, see [`references/OOXML_Zotero_Spec.md`](./references/OOXML_Zotero_Spec.md).

---

## License

Distributed under the **MIT License**. See [`LICENSE`](LICENSE) for details.
