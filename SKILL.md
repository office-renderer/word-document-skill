---
name: word-document-skill
description: Validate, format, and repair supplied DOCX files with the bundled 公文排版Word模板.dotx. Formatting only; do not create or rewrite substantive content.
---

# Word Document Skill

Use the bundled template and repository scripts. Do not manually replay Word formatting rules.

## Normal workflow

Run:

```bash
python scripts/run_skill.py document.docx
```

The runner is validate-first:

```text
validate
  ├─ PASS → finish without rewriting the DOCX
  └─ FAIL → targeted repair → validate again
```

Automatic repair is bounded by `config/rules.json`. If the limit is reached, report the remaining validator codes instead of inventing another repair procedure.

Use `--strict` only for final structural QA or troubleshooting. PDF rendering is separate and runs only when explicitly requested or when the task has entered final visual-layout QA.

## Responsibilities

- `assets/公文排版Word模板.dotx`: template source of truth.
- `config/rules.json`: machine rules and repair limit.
- `scripts/validate_docx.py`: read-only validation.
- `scripts/format_docx.py`: targeted formatting repair.
- `scripts/run_skill.py`: validate-first bounded repair loop.
- `references/template-spec.md`: diagnostic reference only.

## Required safeguards

- Preserve substantive content.
- The first main-title paragraph must be followed by exactly one real empty paragraph; repeated runs must not add more.
- Preserve existing portrait/landscape section orientation; apply template page geometry within each orientation.
- Size tables against the usable width of the section that actually contains them.
- Split only true mixed Chinese/East-Asian + ASCII-alphanumeric plain-text runs; spaces and punctuation alone must not cause extra runs.
- Keep template Chinese fonts such as `仿宋_GB2312`; do not substitute `FangSong`.
- Preserve any existing direct font choice when a mixed run is split.
- Keep `w:themeFontLang/@w:eastAsia = zh-CN`.
- Do not edit the bundled DOTX in place.
- Do not rebuild an existing DOCX from extracted plain text when structure-preserving editing is possible.
- Do not fully reserialize `document.xml` or `styles.xml` with `xml.etree.ElementTree`.
- Do not modify fields, drawings, tabs, breaks, hyperlinks, or other non-plain-text OOXML merely to normalize fonts.
- Remove generator/tool/account metadata from final outputs.
