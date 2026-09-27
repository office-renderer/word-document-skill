---
name: word-document-skill
description: Apply the bundled 公文排版Word模板.dotx to existing or user-supplied content, normalize Word formatting, and validate the resulting DOCX. Use for Word tasks that require this template or the established 公文排版 style. This skill does not draft, rewrite, summarize, expand, research, or fact-check document content.
---

# Word Document Skill

This skill is responsible for Word formatting only. Treat document content as input, not as something this skill should create or rewrite.

## Required workflow

1. Use `assets/公文排版Word模板.dotx` as the formatting source of truth. A current explicit user instruction overrides only the corresponding formatting rule.
2. Take the existing document or user-supplied text/content as-is. Do not draft, rewrite, summarize, expand, fact-check, or otherwise alter substantive content unless that work has already been completed outside this skill.
3. Apply the template structure and styles while preserving document content and objects outside the requested formatting scope.
4. Run:
   ```bash
   python scripts/normalize_docx.py output.docx --in-place
   python scripts/validate_docx.py output.docx
   ```
5. If validation fails, fix only the reported rule IDs and rerun normalization + validation. Do not manually replay all formatting rules from memory.
6. Use `python scripts/validate_docx.py output.docx --strict` only for final structural QA or troubleshooting.
7. DOCX→PDF rendering is separate from validation. Render only when the user explicitly asks for a format/PDF check or when the task has explicitly entered final visual-layout QA.

## Source-of-truth map

- `assets/公文排版Word模板.dotx`: template-defined formatting.
- `config/rules.json`: machine-enforced formatting overrides and validation parameters. Both scripts read this same file.
- `references/template-spec.md`: human-readable template reference; read only when diagnosing layout/template details or modifying this skill.
- `scripts/normalize_docx.py`: mechanically fixes enforceable formatting and removes document-property metadata.
- `scripts/validate_docx.py`: mechanically checks the result; quick by default, strict only on demand.

## Scope boundary

This skill may:

- place supplied content into a Word document;
- map existing paragraphs to template styles;
- normalize body indentation, pagination options, tables, page setup, footers, and metadata;
- validate Word package and formatting structure.

This skill must not:

- generate new substantive text;
- rewrite or polish existing prose;
- summarize or expand content;
- add facts, figures, dates, projects, policies, citations, conclusions, or analysis;
- use web research or general knowledge to change document content.

If content work is needed, complete that work separately first, then use this skill only for Word formatting.

## Non-negotiable safeguards

- Do not edit the bundled DOTX in place.
- Do not rebuild an existing DOCX from extracted plain text when a structure-preserving edit is possible.
- Do not fully reserialize `document.xml` or `styles.xml` with `xml.etree.ElementTree`.
- Do not put generator/tool/account metadata into the final Word file.
- Do not run PDF rendering as part of normal validation.
