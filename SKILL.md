---
name: bibtex-zotero
description: >-
  Automate inserting native Zotero reference field codes into Microsoft Word (.docx) documents
  from BibTeX (.txt or .bib) files. Use when the user wants to convert citation placeholders
  (like {citekey}) into Zotero-linked citations and generate a Word document ready for Zotero's desktop plugin.
---

# BibTeX to Zotero Word Citation Automation Skill

Automate inserting native Microsoft Word Zotero field codes (`ADDIN ZOTERO_ITEM CSL_CITATION`) directly from BibTeX (`.bib` / `.txt`) files without external binary dependencies.

## Capabilities & Features

- **CSL-JSON Conversion**: Converts BibTeX entry types (`@article`, `@inproceedings`, `@book`, `@phdthesis`, etc.) to Citation Style Language (CSL) objects.
- **LaTeX Accent Sanitization**: Converts LaTeX accent codes (`{\L}`, `{\'e}`, `{\`e}`, `\"o`) into clean UTF-8 Unicode characters (`Ł`, `é`, `è`, `ö`) and removes backslashes to avoid Word field switch conflicts (`\f`, `\h`).
- **Depth-Balanced Brace Parsing**: Robustly parses BibTeX field values containing arbitrary nested curly braces `{...}` without dropping author arrays or metadata.
- **Custom Property 255-Char Chunking**: Splits Zotero preference payloads across `ZOTERO_PREF_1` (`pid=2`), `ZOTERO_PREF_2` (`pid=3`) to comply with OpenXML's 255-character string limit.
- **Unique Relationship IDs**: Dynamically computes next available `rIdN` in `_rels/.rels` to prevent ID collisions (`rId5`, `rId6`).
- **Header & Namespace Preservation**: Retains all original `<w:document>` root element attributes (`xmlns:w15`, `xmlns:mc`, `mc:Ignorable`).

## Quick Usage

### CLI Execution
```bash
python scripts/bibtex_to_zotero_word.py --input <input_doc.docx> --bib <references.bib> --output <output_doc.docx>
```

### Installation as Antigravity Skill
1. **Workspace Skill**: Copy this repository to `.agents/skills/bibtex-zotero/` in your project root.
2. **Global Skill**: Copy this repository to `~/.gemini/config/skills/bibtex-zotero/`.

## References & Documentation
- **Python CLI Tool**: [`scripts/bibtex_to_zotero_word.py`](./scripts/bibtex_to_zotero_word.py)
- **Technical Specification**: Refer to [`references/OOXML_Zotero_Spec.md`](./references/OOXML_Zotero_Spec.md) for OpenXML reverse-engineering details.
