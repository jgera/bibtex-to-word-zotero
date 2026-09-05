---
name: bibtex-to-word-zotero
description: >-
  Automate inserting native Zotero reference field codes into Microsoft Word (.docx) documents
  from BibTeX (.txt or .bib) files. Use when the user wants to convert citation placeholders
  (like {citekey}) into Zotero-linked citations, format in-text citations (IEEE, APA, Chicago, MLA),
  or generate Word documents ready for Zotero's desktop plugin.
---

# BibTeX to Word Zotero Citation Automation Skill (`bibtex-to-word-zotero`)

Use this skill to convert citation placeholders (e.g. `{wang2020minivlm}`) in Microsoft Word (`.docx`) documents into **native Zotero OpenXML field codes** (`ADDIN ZOTERO_ITEM CSL_CITATION`) generated directly from BibTeX (`.bib` / `.txt`) files.

---

## When to Activate This Skill

Activate this skill when the user asks to:
- Insert or convert BibTeX references into a Word document (`.docx`).
- Convert inline citation placeholders (e.g., `{citekey}`) into native Zotero fields.
- Format document citations using IEEE, APA, Chicago, MLA, Nature, or Harvard styles.
- Add an automated Zotero Bibliography section (`{bibliography}`).
- Verify or check an existing Word document for Zotero field health.

---

## Command Execution & AI Agent Contract

Run the helper engine [`scripts/bibtex_to_zotero_word.py`](./scripts/bibtex_to_zotero_word.py) using the following parameters:

```bash
python scripts/bibtex_to_zotero_word.py --input <input.docx> --bib <references.bib> --output <output.docx> [OPTIONS]
```

### Supported CLI Parameters

| Flag | Long Flag | Description | Default |
| :--- | :--- | :--- | :--- |
| `-i` | `--input` | Input Word document path containing `{citekey}` placeholders | *Required* |
| `-b` | `--bib` | Input BibTeX file (`.bib` or `.txt`) | *Required* |
| `-o` | `--output` | Target output `.docx` file path | *Required* |
| `-s` | `--style` | Citation style: `ieee`, `apa`, `chicago`, `mla`, `nature`, `vancouver`, `harvard` | `ieee` |
| | `--add-bibliography` | Automatically embed `ADDIN ZOTERO_BIBL` field at `{bibliography}` or document end | `false` |
| | `--check` | Non-mutating diagnostic mode to inspect `.docx` Zotero field health | `false` |
| | `--json` | Return machine-readable JSON execution summary for AI agents | `false` |
| `-q` | `--quiet` | Suppress non-JSON console output | `false` |

---

## AI Agent Step-by-Step Workflow

1. **Locate Input Files**: Verify path to input `.docx` file and `.bib` / `.txt` file.
2. **Execute Conversion with `--json`**:
   ```bash
   python scripts/bibtex_to_zotero_word.py --input doc.docx --bib refs.bib --output output.docx --style ieee --json
   ```
3. **Parse Agent Feedback JSON**:
   ```json
   {
     "status": "success",
     "output_file": "output.docx",
     "citations_inserted": 21,
     "citation_keys": ["wang2020minivlm", "vaswani2017attention"],
     "style": "ieee"
   }
   ```
4. **Verify Document Health**:
   ```bash
   python scripts/bibtex_to_zotero_word.py --input output.docx --check --json
   ```

---

## Troubleshooting & Self-Correction Guidelines

- **File Lock / PermissionError**: If the script returns `"Permission denied when writing to..."`, the file is open in Microsoft Word. Inform the user to close Word or choose a different `--output` file path.
- **Missing Citation Keys**: If a placeholder `{key}` is not replaced, inspect the `.bib` file to ensure `@type{key, ...}` exists.
- **Unreadable Content Warning**: The conversion engine automatically chunks `ZOTERO_PREF_` custom properties to <=255 characters, assigns unique `rIdN` relationship IDs in `_rels/.rels`, and sanitizes LaTeX backslashes to ensure zero recovery warnings in Word.

---

## Technical Reference
For deep architectural details on CSL-JSON payloads, 5-run OpenXML field sequences, and `docProps/custom.xml` properties, see [`references/OOXML_Zotero_Spec.md`](./references/OOXML_Zotero_Spec.md).
