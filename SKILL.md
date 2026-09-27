---
name: word-document-skill
description: Create, edit, and validate Microsoft Word documents that must follow the bundled AI生成文档使用模板.dotx. Use when a .docx/.dotx task explicitly asks to use this template, says “按模板/按AI生成文档使用模板/按公文样式”, or belongs to a workflow that has designated this template as the Word formatting standard. Preserve the template's page setup, heading hierarchy, paragraph styles, TOC behavior, fonts, spacing, and odd/even page numbering unless the user explicitly overrides a specific item. Do not use for unrelated Word documents with a different requested style.
---

# Word Document Skill

Treat `assets/AI生成文档使用模板.dotx` as the authoritative formatting source. Treat `references/template-spec.md` as an explanatory and validation reference, not as a replacement for the template file.

## Core rules

1. Resolve paths relative to this `SKILL.md`; do not assume a fixed installation directory.
2. Preserve the bundled `.dotx` unchanged during ordinary document work. Copy or instantiate it into the working document rather than editing the skill asset in place.
3. Prefer inheriting the template's existing OOXML styles, section settings, footer definitions, and field codes over recreating them manually.
4. Preserve the user's existing document content, tables, images, equations, hyperlinks, fields, bookmarks, section breaks, and other objects unless the task explicitly requires changing them.
5. Apply only requested content/format changes. Do not “beautify” the document by changing template-defined margins, typefaces, heading sizes, paragraph spacing, line spacing, or page-number layout without an explicit instruction.
6. Let an explicit user instruction override the corresponding template rule. Keep all unaffected template rules intact.
7. Do not bundle, copy, or redistribute font files. Refer to fonts by name only.

## Style hierarchy

Use the template's semantic hierarchy rather than inventing a new one:

- Word `Heading 1` / `标题 1`: document main title (`公文标题`), not the first body-level heading.
- Word `Heading 2` / `标题 2`: body level 1 (`一级标题`).
- Word `Heading 3` / `标题 3`: body level 2 (`二级标题`).
- Word `Heading 4` / `标题 4`: body level 3 (`三级标题`).
- Word `Heading 5` / `标题 5`: body level 4 (`四级标题`).

Use the heading text conventions already demonstrated by the template (`一、`, `（一）`, `1.`, `（1）`) unless the user supplies another numbering scheme. Do not assume Word automatic multilevel numbering is the source of those visible labels: the template's numbering definition intentionally leaves several lower-level number texts blank, and example headings contain their visible numbering in the text itself.

Read `references/template-spec.md` when exact font, size, spacing, margin, footer, TOC, or OOXML details are needed.

## New document workflow

1. Start from `assets/AI生成文档使用模板.dotx` so the package keeps the original styles, settings, and odd/even footers.
2. Replace the template's demonstration text with the requested content; do not leave sample paragraphs in the final deliverable.
3. If the chosen library cannot directly instantiate `.dotx`, create a working OOXML copy and convert the package's main content type from Word template to Word document before editing. Do not rebuild the document styles from scratch merely to work around `.dotx` handling.
4. Map paragraphs to the template styles instead of applying equivalent direct formatting where a matching style already exists.
5. Preserve the `PAGE` field in the odd/even footer pair.

## Existing document workflow

1. Inspect the current document before modifying it.
2. Make minimal, structure-preserving edits. Avoid rebuilding the whole DOCX from extracted text.
3. When the user requests this template's formatting, migrate paragraph semantics onto the template's styles and page settings while preserving non-text objects and intentional section-specific exceptions.
4. Keep mixed-orientation or special sections only when they already exist for a reason or the user requests them. The base/default section should still follow the template unless overridden.

## Validation

Run the bundled structural validator after creating or substantially reformatting a Word document:

```bash
python scripts/validate_docx.py /path/to/output.docx
```

The validator compares the target against the bundled `.dotx` for the formatting invariants maintained by this skill. Use `--template` only when validating against a deliberately updated template copy.

Treat validator failures as formatting defects unless they correspond to an explicit user-requested exception. The validator intentionally does not reject arbitrary run-level fonts inside equations, symbols, imported graphics, or other content where direct formatting can be legitimate.

## PDF rendering and visual checks

Do not convert every DOCX to PDF automatically. Render DOCX to PDF only when the user explicitly asks for a PDF/format check, or when the current task has entered a final visual-layout verification stage where rendering is necessary to verify pagination, clipping, page numbers, tables, figures, or other page-level behavior.

When a PDF visual check is performed, verify at minimum:

- title and heading hierarchy;
- paragraph indentation and 28.8 pt fixed line spacing;
- A4 margins and text block placement;
- odd-page page number on the right and even-page page number on the left;
- page-number format `— n —`;
- table/figure overflow, orphaned headings, and unexpected blank pages.

Do not treat visual rendering as a substitute for the structural validator; use both when a final layout check is required.
