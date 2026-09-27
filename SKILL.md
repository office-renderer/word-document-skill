---
name: word-document-skill
description: Create, edit, normalize, and validate Word documents that use the bundled 公文排版Word模板.dotx. Use for Word tasks that explicitly require this template or the established 公文排版 style.
---

# Word Document Skill

Keep the execution path short. Do not load every reference file up front.

## Required workflow

1. Use `assets/公文排版Word模板.dotx` as the formatting source of truth. A current explicit user instruction overrides only the corresponding template rule.
2. If the task includes drafting, rewriting, summarizing, or materially changing text, read `references/content-rules.md`. For format-only work, do not read it.
3. Create or edit the Word document while preserving content and document objects outside the requested scope.
4. Run:
   ```bash
   python scripts/normalize_docx.py output.docx --in-place
   python scripts/validate_docx.py output.docx
   ```
5. If validation fails, fix only the reported rule IDs and rerun normalization + validation. Do not try to remember and manually replay every formatting rule.
6. Use `python scripts/validate_docx.py output.docx --strict` only for final structural QA or troubleshooting.
7. DOCX→PDF rendering is separate from validation. Render only when the user explicitly asks for a format/PDF check or when the task has explicitly entered final visual-layout QA.

## Source-of-truth map

- `assets/公文排版Word模板.dotx`: template-defined formatting.
- `config/rules.json`: machine-enforced overrides and validation parameters. Both scripts read this same file.
- `references/content-rules.md`: text-generation rules; read only for content work.
- `references/template-spec.md`: human-readable template reference; read only when diagnosing layout/template details or modifying this skill.
- `scripts/normalize_docx.py`: mechanically fixes enforceable formatting and removes document-property metadata.
- `scripts/validate_docx.py`: mechanically checks the result; quick by default, strict only on demand.

## Non-negotiable safeguards

- Do not edit the bundled DOTX in place.
- Do not rebuild an existing DOCX from extracted plain text when a structure-preserving edit is possible.
- Do not fully reserialize `document.xml` or `styles.xml` with `xml.etree.ElementTree`.
- Do not put generator/tool/account metadata into the final Word file.
- Do not run PDF rendering as part of normal validation.
