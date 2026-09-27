---
name: word-document-skill
description: Apply the bundled 公文排版Word模板.dotx to supplied Word content, automatically repair formatting from validator reports, and validate the final DOCX. This skill handles Word formatting only and does not create or rewrite substantive content.
---

# Word Document Skill

This skill has one closed loop. Do not manually remember or replay formatting rules.

## Required workflow

1. Preserve the supplied substantive content.
2. Run the formatter:
   ```bash
   python scripts/format_docx.py output.docx --in-place
   ```
3. Run the validator and save its report:
   ```bash
   python scripts/validate_docx.py output.docx --report validation.json
   ```
4. If validation fails and the report contains auto-fixable issues, run:
   ```bash
   python scripts/format_docx.py output.docx --fix validation.json --in-place
   python scripts/validate_docx.py output.docx --report validation.json
   ```
5. Allow at most two automatic repair cycles. If it still fails, stop and report the remaining validator codes. Do not invent another formatting procedure.
6. Use `--strict` only for final structural QA or troubleshooting.
7. DOCX→PDF rendering is separate. Render only when the user explicitly requests a format/PDF check or the task has explicitly entered final visual QA.

## Source of truth

- `assets/公文排版Word模板.dotx`: template-defined formatting.
- `config/rules.json`: machine formatting rules shared by formatter and validator.
- `scripts/format_docx.py`: the only normal entry point for applying and repairing Word formatting.
- `scripts/validate_docx.py`: read-only validation; it never modifies the document.
- `references/template-spec.md`: diagnostic reference only; do not read it during normal execution.

## Scope boundary

This skill may apply styles, indentation, pagination, table formatting, page setup, footers/page numbers, and metadata cleanup. It must not draft, rewrite, summarize, expand, research, or fact-check substantive content.

## Safeguards

- Do not edit the bundled DOTX in place.
- Do not rebuild an existing DOCX from extracted plain text when a structure-preserving edit is possible.
- Do not fully reserialize `document.xml` or `styles.xml` with `xml.etree.ElementTree`.
- Do not put generator/tool/account metadata into the final Word file.
- Do not use PDF rendering as part of normal validation.
