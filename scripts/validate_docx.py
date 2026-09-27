#!/usr/bin/env python3
"""Fast structural validation for the Word skill; use --strict only when needed."""

from __future__ import annotations

import argparse
import io
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_RULES = ROOT / "config" / "rules.json"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W = "{" + W_NS + "}"
R = "{" + R_NS + "}"
NS = {"w": W_NS, "r": R_NS}

PAGINATION_TAGS = ("widowControl", "keepNext", "keepLines", "pageBreakBefore")


def load_rules(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def w_attr(el: ET.Element | None, name: str):
    return None if el is None else el.get(W + name)


class Package:
    def __init__(self, path: Path):
        self.path = path
        self.zf = zipfile.ZipFile(path, "r")
        self.names = set(self.zf.namelist())
        self.cache: dict[str, ET.Element] = {}

    def close(self):
        self.zf.close()

    def xml(self, name: str) -> ET.Element:
        if name not in self.cache:
            if name not in self.names:
                raise ValueError(f"missing OOXML part: {name}")
            try:
                self.cache[name] = ET.fromstring(self.zf.read(name))
            except ET.ParseError as exc:
                raise ValueError(f"invalid XML in {name}: {exc}") from exc
        return self.cache[name]


def style_catalog(styles_root: ET.Element):
    by_id, by_name = {}, {}
    for style in styles_root.findall("w:style", NS):
        sid = w_attr(style, "styleId")
        name = w_attr(style.find("w:name", NS), "val")
        if sid:
            by_id[sid] = style
        if name:
            by_name[name] = style
    return by_id, by_name


def attrs(el: ET.Element | None, names):
    return {name: w_attr(el, name) for name in names}


def style_signature(style: ET.Element, by_id: dict) -> dict:
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
    fonts = rpr.find("w:rFonts", NS) if rpr is not None else None
    sz = rpr.find("w:sz", NS) if rpr is not None else None
    szcs = rpr.find("w:szCs", NS) if rpr is not None else None

    tabs = []
    if ppr is not None:
        for tab in ppr.findall("w:tabs/w:tab", NS):
            tabs.append(attrs(tab, ("val", "leader", "pos")))

    return {
        "based_on": based_name,
        "spacing": attrs(spacing, ("before", "after", "line", "lineRule")),
        "ind": attrs(
            ind,
            (
                "left", "right", "firstLine", "hanging",
                "leftChars", "rightChars", "firstLineChars", "hangingChars",
            ),
        ),
        "jc": w_attr(jc, "val"),
        "outline": w_attr(outline, "val"),
        "tabs": tabs,
        "fonts": attrs(fonts, ("ascii", "eastAsia", "hAnsi", "cs")),
        "sz": w_attr(sz, "val"),
        "szCs": w_attr(szcs, "val"),
        "bold": rpr.find("w:b", NS) is not None if rpr is not None else False,
        "boldCs": rpr.find("w:bCs", NS) is not None if rpr is not None else False,
    }


def doc_defaults_signature(styles: ET.Element) -> dict:
    dd = styles.find("w:docDefaults", NS)
    rpr = dd.find("w:rPrDefault/w:rPr", NS) if dd is not None else None
    ppr = dd.find("w:pPrDefault/w:pPr", NS) if dd is not None else None
    fonts = rpr.find("w:rFonts", NS) if rpr is not None else None
    sz = rpr.find("w:sz", NS) if rpr is not None else None
    szcs = rpr.find("w:szCs", NS) if rpr is not None else None
    spacing = ppr.find("w:spacing", NS) if ppr is not None else None
    jc = ppr.find("w:jc", NS) if ppr is not None else None
    return {
        "fonts": attrs(
            fonts,
            (
                "ascii", "eastAsia", "hAnsi", "cs",
                "asciiTheme", "eastAsiaTheme", "hAnsiTheme", "cstheme",
            ),
        ),
        "sz": w_attr(sz, "val"),
        "szCs": w_attr(szcs, "val"),
        "spacing": attrs(spacing, ("before", "after", "line", "lineRule")),
        "jc": w_attr(jc, "val"),
    }


def section_signature(document: ET.Element) -> tuple[dict, int]:
    sections = document.findall(".//w:sectPr", NS)
    if not sections:
        raise ValueError("document contains no sectPr")
    sect = sections[-1]
    return (
        {
            "pgSz": attrs(sect.find("w:pgSz", NS), ("w", "h", "orient")),
            "pgMar": attrs(
                sect.find("w:pgMar", NS),
                ("top", "right", "bottom", "left", "header", "footer", "gutter"),
            ),
            "docGrid": attrs(sect.find("w:docGrid", NS), ("type", "linePitch")),
        },
        len(sections),
    )


def footer_signatures(pkg: Package, document: ET.Element) -> dict:
    rels = pkg.xml("word/_rels/document.xml.rels")
    targets = {}
    for rel in rels:
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            targets[rid] = target

    sections = document.findall(".//w:sectPr", NS)
    sect = sections[-1]
    result = {}

    for ref in sect.findall("w:footerReference", NS):
        ftype = w_attr(ref, "type")
        rid = ref.get(R + "id")
        target = targets.get(rid)
        if not ftype or not target:
            continue

        root = pkg.xml("word/" + target.lstrip("/"))
        p = root.find("w:p", NS)
        ppr = p.find("w:pPr", NS) if p is not None else None
        rpr = root.find(".//w:r/w:rPr", NS)

        result[ftype] = {
            "jc": w_attr(ppr.find("w:jc", NS), "val") if ppr is not None else None,
            "ind": attrs(
                ppr.find("w:ind", NS) if ppr is not None else None,
                ("left", "right"),
            ),
            "spacing": attrs(
                ppr.find("w:spacing", NS) if ppr is not None else None,
                ("line", "lineRule"),
            ),
            "text": "".join((t.text or "") for t in root.findall(".//w:t", NS)),
            "field": " ".join(
                (x.text or "").strip()
                for x in root.findall(".//w:instrText", NS)
            ).strip(),
            "fonts": attrs(
                rpr.find("w:rFonts", NS) if rpr is not None else None,
                ("ascii", "eastAsia", "hAnsi", "cs"),
            ),
            "sz": w_attr(rpr.find("w:sz", NS), "val") if rpr is not None else None,
            "szCs": w_attr(rpr.find("w:szCs", NS), "val") if rpr is not None else None,
        }

    return result


def template_errors(target: Package, template: Package, rules: dict) -> tuple[list[str], list[str]]:
    errors, warnings = [], []

    t_styles = target.xml("word/styles.xml")
    r_styles = template.xml("word/styles.xml")
    t_doc = target.xml("word/document.xml")
    r_doc = template.xml("word/document.xml")
    t_settings = target.xml("word/settings.xml")
    r_settings = template.xml("word/settings.xml")

    if doc_defaults_signature(t_styles) != doc_defaults_signature(r_styles):
        errors.append("TPL-DEFAULTS: document defaults differ from template")

    t_section, count = section_signature(t_doc)
    r_section, _ = section_signature(r_doc)
    if t_section != r_section:
        errors.append("TPL-SECTION: final/default section page setup differs from template")
    if count > 1:
        warnings.append(
            f"TPL-SECTIONS: document has {count} sections; only the final/default section is compared"
        )

    t_even_odd = t_settings.find("w:evenAndOddHeaders", NS) is not None
    r_even_odd = r_settings.find("w:evenAndOddHeaders", NS) is not None
    if t_even_odd != r_even_odd:
        errors.append("TPL-FOOTER-MODE: even/odd header-footer setting differs from template")

    if footer_signatures(target, t_doc) != footer_signatures(template, r_doc):
        errors.append("TPL-FOOTERS: footer/page-number structure differs from template")

    t_by_id, t_by_name = style_catalog(t_styles)
    r_by_id, r_by_name = style_catalog(r_styles)

    for name in rules["validation"]["template_compare_styles"]:
        if name not in t_by_name:
            errors.append(f"TPL-STYLE-MISSING: {name}")
            continue
        if name not in r_by_name:
            continue
        if style_signature(t_by_name[name], t_by_id) != style_signature(
            r_by_name[name], r_by_id
        ):
            errors.append(f"TPL-STYLE: style differs from template: {name}")

    return errors, warnings


def is_false(node: ET.Element | None) -> bool:
    if node is None:
        return False
    return (w_attr(node, "val") or "true").lower() in {"0", "false", "off"}


def text_of(p: ET.Element) -> str:
    return "".join((t.text or "") for t in p.findall(".//w:t", NS)).strip()


def quick_rule_errors(pkg: Package, rules: dict) -> list[str]:
    errors = []

    for name in rules["validation"]["quick_required_parts"]:
        if name not in pkg.names:
            errors.append(f"PKG-MISSING: {name}")
        elif name.endswith((".xml", ".rels")):
            try:
                pkg.xml(name)
            except ValueError as exc:
                errors.append(f"PKG-XML: {exc}")

    remove_parts = set(rules["metadata"]["remove_parts"])
    rel_types = set(rules["metadata"]["relationship_types"])
    for part in sorted(remove_parts):
        if part in pkg.names:
            errors.append(f"META-PART: {part}")

    if "_rels/.rels" in pkg.names:
        root = pkg.xml("_rels/.rels")
        for rel in root:
            target = rel.get("Target", "").lstrip("/")
            if rel.get("Type", "") in rel_types or target in remove_parts:
                errors.append(f"META-REL: {target}")

    if "[Content_Types].xml" in pkg.names:
        root = pkg.xml("[Content_Types].xml")
        for child in root:
            part_name = child.get("PartName", "").lstrip("/")
            if part_name in remove_parts:
                errors.append(f"META-CONTENT-TYPE: {part_name}")

    if "word/styles.xml" not in pkg.names or "word/document.xml" not in pkg.names:
        return errors

    styles = pkg.xml("word/styles.xml")
    document = pkg.xml("word/document.xml")
    by_id, by_name = style_catalog(styles)

    body_name = rules["body"]["style"]
    body_style = by_name.get(body_name)
    if body_style is None:
        errors.append(f"FMT-BODY-STYLE-MISSING: {body_name}")
        body_id = None
    else:
        body_id = w_attr(body_style, "styleId")
        ind = body_style.find("w:pPr/w:ind", NS)
        expected = str(rules["body"]["first_line_twips"])
        if w_attr(ind, "firstLine") != expected:
            errors.append(
                f"FMT-BODY-INDENT: expected {expected}, got {w_attr(ind, 'firstLine')!r}"
            )

    for style in styles.findall("w:style", NS):
        if w_attr(style, "type") != "paragraph":
            continue
        name = w_attr(style.find("w:name", NS), "val") or w_attr(style, "styleId") or "?"
        ppr = style.find("w:pPr", NS)
        for tag in PAGINATION_TAGS:
            node = ppr.find(f"w:{tag}", NS) if ppr is not None else None
            if not is_false(node):
                errors.append(f"FMT-PAGINATION-STYLE: {name}.{tag}")

    for idx, p in enumerate(document.findall(".//w:p", NS), 1):
        ppr = p.find("w:pPr", NS)
        if ppr is None:
            continue
        for tag in PAGINATION_TAGS:
            node = ppr.find(f"w:{tag}", NS)
            if node is not None and not is_false(node):
                errors.append(f"FMT-PAGINATION-PARA: paragraph {idx}.{tag}")

    body = document.find("w:body", NS)
    if body is not None and body_id:
        min_chars = int(rules["body"]["auto_map_min_chars"])
        auto_from = set(rules["body"]["auto_map_from_styles"])
        for idx, p in enumerate(body.findall("w:p", NS), 1):
            text = text_of(p)
            if len(text) < min_chars:
                continue
            ppr = p.find("w:pPr", NS)
            jc = ppr.find("w:jc", NS) if ppr is not None else None
            if w_attr(jc, "val") in {"center", "right"}:
                continue
            pstyle = ppr.find("w:pStyle", NS) if ppr is not None else None
            sid = w_attr(pstyle, "val")
            name = (
                w_attr(by_id[sid].find("w:name", NS), "val")
                if sid in by_id
                else "Normal" if sid is None else sid
            )
            if name in auto_from:
                errors.append(
                    f"FMT-BODY-USAGE: paragraph {idx} uses {name!r} instead of {body_name!r}"
                )
            if sid == body_id and ppr is not None:
                direct = ppr.find("w:ind", NS)
                if direct is not None and any(
                    w_attr(direct, key) is not None
                    for key in ("firstLine", "firstLineChars", "hanging", "hangingChars")
                ):
                    errors.append(
                        f"FMT-BODY-DIRECT-INDENT: paragraph {idx} overrides body indent"
                    )

    max_width = int(rules["table"]["max_width_twips"])
    min_font = int(rules["table"]["font_half_points"])

    for t_idx, tbl in enumerate(document.findall(".//w:tbl", NS), 1):
        tblpr = tbl.find("w:tblPr", NS)
        tblw = tblpr.find("w:tblW", NS) if tblpr is not None else None
        layout = tblpr.find("w:tblLayout", NS) if tblpr is not None else None
        tblind = tblpr.find("w:tblInd", NS) if tblpr is not None else None

        if w_attr(tblw, "type") != "dxa":
            errors.append(f"TBL-WIDTH-TYPE: table {t_idx}")
        else:
            try:
                if int(w_attr(tblw, "w") or "0") > max_width:
                    errors.append(f"TBL-WIDTH: table {t_idx}")
            except ValueError:
                errors.append(f"TBL-WIDTH: table {t_idx}")

        if w_attr(layout, "type") != "fixed":
            errors.append(f"TBL-LAYOUT: table {t_idx}")
        if tblind is not None and w_attr(tblind, "w") not in {None, "0"}:
            errors.append(f"TBL-INDENT: table {t_idx}")

        grid = tbl.find("w:tblGrid", NS)
        if grid is not None:
            total = 0
            for col in grid.findall("w:gridCol", NS):
                try:
                    total += int(w_attr(col, "w") or "0")
                except ValueError:
                    pass
            if total > max_width:
                errors.append(f"TBL-GRID-WIDTH: table {t_idx} = {total}")

        for c_idx, tc in enumerate(tbl.findall(".//w:tc", NS), 1):
            tcpr = tc.find("w:tcPr", NS)
            valign = tcpr.find("w:vAlign", NS) if tcpr is not None else None
            if w_attr(valign, "val") != "center":
                errors.append(f"TBL-VERTICAL-ALIGN: table {t_idx} cell {c_idx}")

            for p_idx, p in enumerate(tc.findall("w:p", NS), 1):
                ppr = p.find("w:pPr", NS)
                jc = ppr.find("w:jc", NS) if ppr is not None else None
                ind = ppr.find("w:ind", NS) if ppr is not None else None

                if w_attr(jc, "val") != "center":
                    errors.append(
                        f"TBL-HORIZONTAL-ALIGN: table {t_idx} cell {c_idx} paragraph {p_idx}"
                    )

                required_zero = (
                    "left", "right", "firstLine",
                    "leftChars", "rightChars", "firstLineChars",
                    "start", "end", "startChars", "endChars",
                )
                hanging = (
                    w_attr(ind, "hanging") if ind is not None else None,
                    w_attr(ind, "hangingChars") if ind is not None else None,
                )
                if (
                    ind is None
                    or any(w_attr(ind, key) != "0" for key in required_zero)
                    or any(value not in {None, "0"} for value in hanging)
                ):
                    errors.append(
                        f"TBL-PARA-INDENT: table {t_idx} cell {c_idx} paragraph {p_idx}"
                    )

                for r_idx, run in enumerate(p.findall(".//w:r", NS), 1):
                    if not "".join((t.text or "") for t in run.findall(".//w:t", NS)):
                        continue
                    rpr = run.find("w:rPr", NS)
                    sz = w_attr(rpr.find("w:sz", NS), "val") if rpr is not None else None
                    szcs = w_attr(rpr.find("w:szCs", NS), "val") if rpr is not None else None
                    for value in (sz, szcs):
                        if value is None:
                            continue
                        try:
                            if int(value) < min_font:
                                errors.append(
                                    f"TBL-FONT: table {t_idx} cell {c_idx} paragraph {p_idx} run {r_idx}"
                                )
                                break
                        except ValueError:
                            errors.append(
                                f"TBL-FONT: table {t_idx} cell {c_idx} paragraph {p_idx} run {r_idx}"
                            )
                            break

    return errors


def strict_errors(pkg: Package) -> list[str]:
    errors = []

    bad = pkg.zf.testzip()
    if bad:
        errors.append(f"STRICT-ZIP: corrupt member {bad}")

    for name in pkg.names:
        if not name.endswith((".xml", ".rels")):
            continue
        data = pkg.zf.read(name)
        try:
            root = ET.fromstring(data)
        except ET.ParseError as exc:
            errors.append(f"STRICT-XML: {name}: {exc}")
            continue

        ignorable = root.get("{" + MC_NS + "}Ignorable")
        if ignorable:
            declared = set()
            try:
                for _event, ns in ET.iterparse(io.BytesIO(data), events=("start-ns",)):
                    prefix, _uri = ns
                    declared.add(prefix or "")
            except ET.ParseError:
                continue
            for prefix in ignorable.split():
                if prefix not in declared:
                    errors.append(
                        f"STRICT-IGNORABLE: {name} references undeclared prefix {prefix!r}"
                    )

    try:
        from docx import Document
        Document(pkg.path)
    except Exception as exc:
        errors.append(f"STRICT-REOPEN: python-docx could not reopen file: {exc}")

    return errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("document", type=Path)
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    rules = load_rules(args.rules)
    template_path = args.template or ROOT / rules["template_path"]

    errors, warnings = [], []
    target = template = None

    try:
        target = Package(args.document)
        template = Package(template_path)
        errors.extend(quick_rule_errors(target, rules))
        e, w = template_errors(target, template, rules)
        errors.extend(e)
        warnings.extend(w)
        if args.strict:
            errors.extend(strict_errors(target))
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        errors.append(f"PKG-OPEN: {exc}")
    finally:
        if target is not None:
            target.close()
        if template is not None:
            template.close()

    result = {
        "ok": not errors,
        "mode": "strict" if args.strict else "quick",
        "document": str(args.document),
        "template": str(template_path),
        "errors": errors,
        "warnings": warnings,
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        if errors:
            print(f"FAIL [{result['mode']}]: {len(errors)} issue(s)")
            for item in errors:
                print(f"  - {item}")
        else:
            print(f"PASS [{result['mode']}]")
        for item in warnings:
            print(f"WARNING: {item}")

    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
