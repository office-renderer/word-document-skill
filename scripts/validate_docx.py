#!/usr/bin/env python3
"""Validate a DOCX/DOTX against the bundled Word template's structural invariants.

Uses only the Python standard library. It intentionally focuses on package/style/page
structure and does not police every run-level font because direct formatting can be
legitimate for equations, symbols, imported content, and other special cases.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS, "r": R_NS}
W = "{" + W_NS + "}"
R = "{" + R_NS + "}"

PAGINATION_TAGS = ("widowControl", "keepNext", "keepLines", "pageBreakBefore")
BODY_STYLE_NAME = "正文（默认）"
NORMAL_STYLE_NAME = "Normal"
BODY_FIRST_LINE_TWIPS = "640"

CORE_STYLE_NAMES = [
    "Normal",
    "正文（默认）",
    "公文标题",
    "一级标题",
    "二级标题",
    "三级标题",
    "四级标题",
    "heading 1",
    "heading 2",
    "heading 3",
    "heading 4",
    "heading 5",
    "TOC Heading",
    "目录标题",
    "TOC 11",
    "目录一级条目",
    "TOC 21",
    "目录二级条目",
    "TOC 31",
    "目录三级条目",
]


def w_attr(el: ET.Element | None, name: str):
    return None if el is None else el.get(W + name)


def read_xml(zf: zipfile.ZipFile, member: str) -> ET.Element:
    try:
        return ET.fromstring(zf.read(member))
    except KeyError as exc:
        raise ValueError(f"missing OOXML part: {member}") from exc


def element_attrs(el: ET.Element | None, names: tuple[str, ...]) -> dict[str, str | None]:
    return {name: w_attr(el, name) for name in names}


def style_catalog(styles_root: ET.Element) -> tuple[dict[str, ET.Element], dict[str, ET.Element]]:
    by_id: dict[str, ET.Element] = {}
    by_name: dict[str, ET.Element] = {}
    for style in styles_root.findall("w:style", NS):
        sid = w_attr(style, "styleId")
        name = w_attr(style.find("w:name", NS), "val")
        if sid:
            by_id[sid] = style
        if name:
            by_name[name] = style
    return by_id, by_name


def style_signature(style: ET.Element, by_id: dict[str, ET.Element]) -> dict:
    based_id = w_attr(style.find("w:basedOn", NS), "val")
    based_name = None
    if based_id and based_id in by_id:
        based_name = w_attr(by_id[based_id].find("w:name", NS), "val")

    ppr = style.find("w:pPr", NS)
    rpr = style.find("w:rPr", NS)
    spacing = ppr.find("w:spacing", NS) if ppr is not None else None
    ind = ppr.find("w:ind", NS) if ppr is not None else None
    jc = ppr.find("w:jc", NS) if ppr is not None else None
    outline = ppr.find("w:outlineLvl", NS) if ppr is not None else None
    numpr = ppr.find("w:numPr", NS) if ppr is not None else None
    ilvl = numpr.find("w:ilvl", NS) if numpr is not None else None
    numid = numpr.find("w:numId", NS) if numpr is not None else None
    fonts = rpr.find("w:rFonts", NS) if rpr is not None else None
    sz = rpr.find("w:sz", NS) if rpr is not None else None
    szcs = rpr.find("w:szCs", NS) if rpr is not None else None

    tabs = []
    if ppr is not None:
        for tab in ppr.findall("w:tabs/w:tab", NS):
            tabs.append(element_attrs(tab, ("val", "leader", "pos")))

    return {
        "based_on": based_name,
        "p": {
            "spacing": element_attrs(spacing, ("before", "after", "line", "lineRule")),
            "ind": element_attrs(ind, ("left", "right", "firstLine", "hanging", "leftChars", "rightChars", "firstLineChars", "hangingChars")),
            "jc": w_attr(jc, "val"),
            "outline": w_attr(outline, "val"),
            "num": {"ilvl": w_attr(ilvl, "val"), "numId": w_attr(numid, "val")},
            "tabs": tabs,
            "wordWrap": w_attr(ppr.find("w:wordWrap", NS), "val") if ppr is not None else None,
            "adjustRightInd": w_attr(ppr.find("w:adjustRightInd", NS), "val") if ppr is not None else None,
            "snapToGrid": w_attr(ppr.find("w:snapToGrid", NS), "val") if ppr is not None else None,
        },
        "r": {
            "fonts": element_attrs(fonts, ("ascii", "eastAsia", "hAnsi", "cs")),
            "sz": w_attr(sz, "val"),
            "szCs": w_attr(szcs, "val"),
            "bold": rpr.find("w:b", NS) is not None if rpr is not None else False,
            "boldCs": rpr.find("w:bCs", NS) is not None if rpr is not None else False,
        },
    }


def doc_defaults_signature(styles_root: ET.Element) -> dict:
    dd = styles_root.find("w:docDefaults", NS)
    rpr = dd.find("w:rPrDefault/w:rPr", NS) if dd is not None else None
    ppr = dd.find("w:pPrDefault/w:pPr", NS) if dd is not None else None
    fonts = rpr.find("w:rFonts", NS) if rpr is not None else None
    sz = rpr.find("w:sz", NS) if rpr is not None else None
    szcs = rpr.find("w:szCs", NS) if rpr is not None else None
    spacing = ppr.find("w:spacing", NS) if ppr is not None else None
    jc = ppr.find("w:jc", NS) if ppr is not None else None
    return {
        "fonts": element_attrs(fonts, ("ascii", "eastAsia", "hAnsi", "cs", "asciiTheme", "eastAsiaTheme", "hAnsiTheme", "cstheme")),
        "sz": w_attr(sz, "val"),
        "szCs": w_attr(szcs, "val"),
        "spacing": element_attrs(spacing, ("before", "after", "line", "lineRule")),
        "jc": w_attr(jc, "val"),
    }


def section_signature(document_root: ET.Element) -> tuple[dict, int]:
    sections = document_root.findall(".//w:sectPr", NS)
    if not sections:
        raise ValueError("document contains no sectPr")
    sect = sections[-1]
    pgsz = sect.find("w:pgSz", NS)
    pgmar = sect.find("w:pgMar", NS)
    grid = sect.find("w:docGrid", NS)
    return (
        {
            "pgSz": element_attrs(pgsz, ("w", "h", "orient")),
            "pgMar": element_attrs(pgmar, ("top", "right", "bottom", "left", "header", "footer", "gutter")),
            "docGrid": element_attrs(grid, ("type", "linePitch")),
        },
        len(sections),
    )


def footer_signatures(zf: zipfile.ZipFile, document_root: ET.Element) -> dict[str, dict]:
    rels = read_xml(zf, "word/_rels/document.xml.rels")
    rel_targets = {}
    for rel in rels:
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            rel_targets[rid] = target

    sections = document_root.findall(".//w:sectPr", NS)
    sect = sections[-1]
    out = {}
    for ref in sect.findall("w:footerReference", NS):
        ftype = w_attr(ref, "type")
        rid = ref.get(R + "id")
        target = rel_targets.get(rid)
        if not (ftype and target):
            continue
        member = "word/" + target.lstrip("/")
        root = read_xml(zf, member)
        p = root.find("w:p", NS)
        ppr = p.find("w:pPr", NS) if p is not None else None
        spacing = ppr.find("w:spacing", NS) if ppr is not None else None
        ind = ppr.find("w:ind", NS) if ppr is not None else None
        jc = ppr.find("w:jc", NS) if ppr is not None else None

        text = "".join((t.text or "") for t in root.findall(".//w:t", NS))
        field_instr = " ".join((x.text or "").strip() for x in root.findall(".//w:instrText", NS)).strip()
        first_rpr = root.find(".//w:r/w:rPr", NS)
        fonts = first_rpr.find("w:rFonts", NS) if first_rpr is not None else None
        sz = first_rpr.find("w:sz", NS) if first_rpr is not None else None
        szcs = first_rpr.find("w:szCs", NS) if first_rpr is not None else None
        out[ftype] = {
            "spacing": element_attrs(spacing, ("line", "lineRule")),
            "ind": element_attrs(ind, ("left", "right")),
            "jc": w_attr(jc, "val"),
            "text": text,
            "field": field_instr,
            "fonts": element_attrs(fonts, ("ascii", "eastAsia", "hAnsi", "cs")),
            "sz": w_attr(sz, "val"),
            "szCs": w_attr(szcs, "val"),
        }
    return out


def build_profile(path: Path) -> dict:
    if not path.is_file():
        raise ValueError(f"file not found: {path}")
    try:
        with zipfile.ZipFile(path) as zf:
            styles = read_xml(zf, "word/styles.xml")
            document = read_xml(zf, "word/document.xml")
            settings = read_xml(zf, "word/settings.xml")
            by_id, by_name = style_catalog(styles)
            style_profiles = {}
            missing_styles = []
            for name in CORE_STYLE_NAMES:
                style = by_name.get(name)
                if style is None:
                    missing_styles.append(name)
                else:
                    style_profiles[name] = style_signature(style, by_id)
            section, section_count = section_signature(document)
            return {
                "doc_defaults": doc_defaults_signature(styles),
                "styles": style_profiles,
                "missing_styles": missing_styles,
                "section": section,
                "section_count": section_count,
                "even_and_odd_headers": settings.find("w:evenAndOddHeaders", NS) is not None,
                "footers": footer_signatures(zf, document),
            }
    except zipfile.BadZipFile as exc:
        raise ValueError(f"not a valid OOXML zip package: {path}") from exc


def compare(expected, actual, path="") -> list[str]:
    errors: list[str] = []
    if isinstance(expected, dict):
        if not isinstance(actual, dict):
            return [f"{path or '<root>'}: expected mapping, got {type(actual).__name__}"]
        for key, value in expected.items():
            child = f"{path}.{key}" if path else key
            if key not in actual:
                errors.append(f"{child}: missing")
            else:
                errors.extend(compare(value, actual[key], child))
        return errors
    if isinstance(expected, list):
        if expected != actual:
            errors.append(f"{path}: expected {expected!r}, got {actual!r}")
        return errors
    if expected != actual:
        errors.append(f"{path}: expected {expected!r}, got {actual!r}")
    return errors



def _is_false(node: ET.Element | None) -> bool:
    if node is None:
        return False
    return (w_attr(node, "val") or "true").lower() in {"0", "false", "off"}


def _paragraph_text(p: ET.Element) -> str:
    return "".join((t.text or "") for t in p.findall(".//w:t", NS)).strip()


def persistent_rule_errors(path: Path) -> list[str]:
    errors: list[str] = []
    with zipfile.ZipFile(path) as zf:
        styles = read_xml(zf, "word/styles.xml")
        document = read_xml(zf, "word/document.xml")
        by_id, by_name = style_catalog(styles)

        for style in styles.findall("w:style", NS):
            if w_attr(style, "type") != "paragraph":
                continue
            name = w_attr(style.find("w:name", NS), "val") or w_attr(style, "styleId") or "<unnamed>"
            ppr = style.find("w:pPr", NS)
            for tag in PAGINATION_TAGS:
                node = ppr.find(f"w:{tag}", NS) if ppr is not None else None
                if not _is_false(node):
                    errors.append(f"pagination style {name}.{tag}: must be explicitly disabled")

        body_style = by_name.get(BODY_STYLE_NAME)
        body_id = w_attr(body_style, "styleId") if body_style is not None else None
        if body_style is None:
            errors.append(f"required body style missing: {BODY_STYLE_NAME}")
        else:
            ind = body_style.find("w:pPr/w:ind", NS)
            if w_attr(ind, "firstLine") != BODY_FIRST_LINE_TWIPS:
                errors.append(f"body style firstLine: expected {BODY_FIRST_LINE_TWIPS}, got {w_attr(ind, 'firstLine')!r}")
            if w_attr(ind, "hanging") is not None or w_attr(ind, "hangingChars") is not None:
                errors.append("body style must not use hanging indentation")

        for idx, p in enumerate(document.findall(".//w:p", NS), 1):
            ppr = p.find("w:pPr", NS)
            if ppr is None:
                continue
            for tag in PAGINATION_TAGS:
                node = ppr.find(f"w:{tag}", NS)
                if node is not None and not _is_false(node):
                    errors.append(f"paragraph {idx} {tag}: enabled ({_paragraph_text(p)[:24]!r})")

        body = document.find("w:body", NS)
        if body is not None and body_id:
            for idx, p in enumerate(body.findall("w:p", NS), 1):
                text = _paragraph_text(p)
                if len(text) < 12:
                    continue
                ppr = p.find("w:pPr", NS)
                jc = ppr.find("w:jc", NS) if ppr is not None else None
                if w_attr(jc, "val") in {"center", "right"}:
                    continue
                pstyle = ppr.find("w:pStyle", NS) if ppr is not None else None
                sid = w_attr(pstyle, "val")
                style_name = None
                if sid in by_id:
                    style_name = w_attr(by_id[sid].find("w:name", NS), "val")
                elif sid is None:
                    style_name = NORMAL_STYLE_NAME
                if style_name == NORMAL_STYLE_NAME:
                    errors.append(f"body paragraph {idx}: uses Normal/unstyled instead of {BODY_STYLE_NAME} ({text[:24]!r})")
                if sid == body_id and ppr is not None:
                    ind = ppr.find("w:ind", NS)
                    if ind is not None and any(
                        w_attr(ind, key) is not None
                        for key in ("firstLine", "firstLineChars", "hanging", "hangingChars")
                    ):
                        errors.append(f"body paragraph {idx}: direct indentation overrides {BODY_STYLE_NAME} ({text[:24]!r})")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate DOCX/DOTX formatting against the bundled template plus body-indent and pagination rules.")
    parser.add_argument("document", type=Path, help="DOCX/DOTX file to validate")
    parser.add_argument(
        "--template",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "assets" / "公文排版Word模板.dotx",
        help="reference template (defaults to the bundled asset)",
    )
    parser.add_argument("--json", action="store_true", help="emit machine-readable JSON")
    parser.add_argument("--dump-profile", action="store_true", help="print the target structural profile and exit")
    args = parser.parse_args()

    try:
        actual = build_profile(args.document)
        if args.dump_profile:
            print(json.dumps(actual, ensure_ascii=False, indent=2))
            return 0
        expected = build_profile(args.template)
        rule_errors = persistent_rule_errors(args.document)
    except ValueError as exc:
        if args.json:
            print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        else:
            print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    expected_cmp = dict(expected)
    actual_cmp = dict(actual)
    expected_cmp.pop("section_count", None)
    actual_cmp.pop("section_count", None)

    errors = compare(expected_cmp, actual_cmp) + rule_errors
    warnings = []
    if actual.get("section_count", 1) > 1:
        warnings.append(
            f"document contains {actual['section_count']} sections; validator compares the final/default section to the template and leaves intentional special sections to visual review"
        )

    result = {
        "ok": not errors,
        "document": str(args.document),
        "template": str(args.template),
        "errors": errors,
        "warnings": warnings,
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        if result["ok"]:
            print("PASS: document matches the bundled template's structural formatting invariants.")
        else:
            print(f"FAIL: {len(errors)} formatting mismatch(es) found.")
            for err in errors:
                print(f"  - {err}")
        for warning in warnings:
            print(f"WARNING: {warning}")

    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
