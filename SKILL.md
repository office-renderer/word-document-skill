---
name: word-document-skill
description: Format and validate supplied DOCX files with the bundled 公文排版Word模板.dotx. Formatting only; do not create or rewrite substantive content.
---

# Word Document Skill

Use the bundled template and the repository scripts. Do not manually replay Word formatting rules.

## Normal workflow

Run one command:

```bash
python scripts/run_skill.py document.docx
```

It performs:

```text
format → validate → targeted repair → validate
```

Automatic repair is bounded by `config/rules.json`. If the limit is reached, stop and report the remaining validator codes; do not invent another repair procedure.

Use `--strict` only for final structural QA or troubleshooting. DOCX→PDF rendering is separate and should run only when explicitly requested or when the task has entered final visual-layout QA.

## Responsibilities

- `assets/公文排版Word模板.dotx`: template source of truth.
- `config/rules.json`: machine rules and repair limit.
- `scripts/format_docx.py`: formatting and targeted repair.
- `scripts/validate_docx.py`: read-only validation.
- `scripts/run_skill.py`: bounded format/validate/repair loop.
- `references/template-spec.md`: diagnostic reference only; do not read during normal execution.

## Required safeguards

- Preserve substantive content.
- Keep template Chinese fonts such as `仿宋_GB2312`; do not substitute `FangSong`.
- Plain mixed Chinese/ASCII text must be split into separate Word runs with explicit font slots; preserve any existing direct font choice.
- Keep `w:themeFontLang/@w:eastAsia = zh-CN`.
- Do not edit the bundled DOTX in place.
- Do not rebuild an existing DOCX from extracted plain text when structure-preserving editing is possible.
- Do not fully reserialize `document.xml` or `styles.xml` with `xml.etree.ElementTree`.
- Do not modify fields, drawings, tabs, breaks, hyperlinks, or other non-plain-text OOXML merely to normalize fonts.
- Remove generator/tool/account metadata from final outputs.
