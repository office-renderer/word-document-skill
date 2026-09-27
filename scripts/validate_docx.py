#!/usr/bin/env python3
"""Read-only Word validation with structured auto-repair instructions."""

from __future__ import annotations

import argparse
import io
import json
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

PAGINATION_OOXML = {
    "widow_control": "widowControl",
    "keep_with_next": "keepNext",
    "keep_together": "keepLines",
    "page_break_before": "pageBreakBefore",
}


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def w_attr(el: ET.Element | None, name: str):
    return None if el is None else el.get(W + name)


def issue(code: str, message: str, fix_group: str | None = None, **context):
    return {
        "code": code,
        "message": message,
        "auto_fixable": fix_group is not None,
        "fix_group": fix_group,
        "context": context,
    }


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


def section_signature(document: ET.Element) -> dict:
    sections = document.findall(".//w:sectPr", NS)
    if not sections:
        raise ValueError("document contains no sectPr")
    sect = sections[-1]
    return {
        "pgSz": attrs(sect.find("w:pgSz", NS), ("w", "h", "orient")),
        "pgMar": attrs(
            sect.find("w:pgMar", NS),
            ("top", "right", "bottom", "left", "header", "footer", "gutter"),
        ),
        "docGrid": attrs(sect.find("w:docGrid", NS), ("type", "linePitch")),
    }


def footer_signature(pkg: Package, document: ET.Element) -> dict:
    """Return only footer semantics that affect the rendered page-number layout."""
    rels = pkg.xml("word/_rels/document.xml.rels")
    targets = {}
    for rel in rels:
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            targets[rid] = target

    sect = document.findall(".//w:sectPr", NS)[-1]
    result = {}

    for ref in sect.findall("w:footerReference", NS):
        ftype = w_attr(ref, "type")
        rid = ref.get(R + "id")
        target = targets.get(rid)
        if not ftype or not target:
            continue

        member = "word/" + target.lstrip("/")
        if member not in pkg.names:
            continue

        root = pkg.xml(member)
        p = root.find("w:p", NS)
        ppr = p.find("w:pPr", NS) if p is not None else None

        field_text = " ".join(
            (x.text or "").strip()
            for x in root.findall(".//w:instrText", NS)
        )
        field_tokens = field_text.upper().split()
        has_page_field = "PAGE" in field_tokens

        literal_text = "".join(
            (t.text or "") for t in root.findall(".//w:t", NS)
        )
        literal_text = "".join(literal_text.split())

        result[ftype] = {
            "jc": w_attr(ppr.find("w:jc", NS), "val") if ppr is not None else None,
            "ind": attrs(
                ppr.find("w:ind", NS) if ppr is not None else None,
                ("left", "right"),
            ),
            "has_page_field": has_page_field,
            "dash_count": literal_text.count("—"),
        }

    return result


def is_false(node: ET.Element | None) -> bool:
    if node is None:
        return False
    return (w_attr(node, "val") or "true").lower() in {"0", "false", "off"}


def text_of(p: ET.Element) -> str:
    return "".join((t.text or "") for t in p.findall(".//w:t", NS)).strip()


def is_vmerge_continuation(tc: ET.Element) -> bool:
    """True for a vertical-merge continuation cell that Word does not render independently."""
    tcpr = tc.find("w:tcPr", NS)
    vmerge = tcpr.find("w:vMerge", NS) if tcpr is not None else None
    if vmerge is None:
        return False
    value = w_attr(vmerge, "val")
    return value in {None, "", "continue"}


def quick_issues(target: Package, template: Package, rules: dict) -> tuple[list[dict], list[str]]:
    issues = []
    warnings = []

    for name in rules["validation"]["quick_required_parts"]:
        if name not in target.names:
            issues.append(issue("PKG-MISSING", f"Missing OOXML part: {name}", part=name))
        elif name.endswith((".xml", ".rels")):
            try:
                target.xml(name)
            except ValueError as exc:
                issues.append(issue("PKG-XML", str(exc), part=name))

    if any(x["code"].startswith("PKG-") for x in issues):
        return issues, warnings

    remove_parts = set(rules["metadata"]["remove_parts"])
    rel_types = set(rules["metadata"]["relationship_types"])

    for part in sorted(remove_parts):
        if part in target.names:
            issues.append(issue("META-PART", f"Metadata part remains: {part}", "metadata", part=part))

    root_rels = target.xml("_rels/.rels")
    for rel in root_rels:
        rel_target = rel.get("Target", "").lstrip("/")
        if rel.get("Type", "") in rel_types or rel_target in remove_parts:
            issues.append(issue("META-REL", f"Metadata relationship remains: {rel_target}", "metadata", target=rel_target))

    content_types = target.xml("[Content_Types].xml")
    for child in content_types:
        part_name = child.get("PartName", "").lstrip("/")
        if part_name in remove_parts:
            issues.append(issue("META-CONTENT-TYPE", f"Metadata content type remains: {part_name}", "metadata", part=part_name))

    document = target.xml("word/document.xml")
    styles = target.xml("word/styles.xml")
    settings = target.xml("word/settings.xml")
    by_id, by_name = style_catalog(styles)

    ref_document = template.xml("word/document.xml")
    ref_settings = template.xml("word/settings.xml")

    if section_signature(document) != section_signature(ref_document):
        issues.append(issue("PAGE-SETUP", "Final/default section page setup differs from template", "page"))

    target_even_odd = settings.find("w:evenAndOddHeaders", NS) is not None
    ref_even_odd = ref_settings.find("w:evenAndOddHeaders", NS) is not None
    if target_even_odd != ref_even_odd:
        issues.append(issue("FOOTER-MODE", "Odd/even footer setting differs from template", "footer"))

    try:
        if footer_signature(target, document) != footer_signature(template, ref_document):
            issues.append(issue("FOOTER-STRUCTURE", "Footer/page-number structure differs from template", "footer"))
    except ValueError as exc:
        issues.append(issue("FOOTER-STRUCTURE", str(exc), "footer"))

    section_count = len(document.findall(".//w:sectPr", NS))
    if section_count > 1:
        warnings.append(
            f"Document has {section_count} sections; page/footer checks use the final/default section."
        )

    body_name = rules["body"]["style"]
    body_style = by_name.get(body_name)
    if body_style is None:
        issues.append(issue("FMT-BODY-STYLE-MISSING", f"Missing body style: {body_name}", "body"))
        body_id = None
    else:
        body_id = w_attr(body_style, "styleId")
        ind = body_style.find("w:pPr/w:ind", NS)
        expected = str(rules["body"]["first_line_twips"])
        if w_attr(ind, "firstLine") != expected:
            issues.append(issue("FMT-BODY-INDENT", f"Body firstLine should be {expected} twips", "body", actual=w_attr(ind, "firstLine")))

    disabled_tags = [PAGINATION_OOXML[name] for name in rules["pagination"]["disable"]]

    for style in styles.findall("w:style", NS):
        if w_attr(style, "type") != "paragraph":
            continue
        style_name = w_attr(style.find("w:name", NS), "val") or w_attr(style, "styleId") or "?"
        ppr = style.find("w:pPr", NS)
        for tag in disabled_tags:
            node = ppr.find(f"w:{tag}", NS) if ppr is not None else None
            if not is_false(node):
                issues.append(issue("FMT-PAGINATION-STYLE", f"{style_name}.{tag} is not disabled", "pagination", style=style_name, property=tag))

    for idx, p in enumerate(document.findall(".//w:p", NS), 1):
        ppr = p.find("w:pPr", NS)
        if ppr is None:
            continue
        for tag in disabled_tags:
            node = ppr.find(f"w:{tag}", NS)
            if node is not None and not is_false(node):
                issues.append(issue("FMT-PAGINATION-PARA", f"Paragraph {idx} enables {tag}", "pagination", paragraph=idx, property=tag))

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
            style_name = (
                w_attr(by_id[sid].find("w:name", NS), "val")
                if sid in by_id
                else "Normal" if sid is None else sid
            )
            if style_name in auto_from:
                issues.append(issue("FMT-BODY-USAGE", f"Paragraph {idx} uses {style_name!r} instead of {body_name!r}", "body", paragraph=idx))

            if sid == body_id and ppr is not None:
                direct = ppr.find("w:ind", NS)
                if direct is not None and any(
                    w_attr(direct, key) is not None
                    for key in ("firstLine", "firstLineChars", "hanging", "hangingChars")
                ):
                    issues.append(issue("FMT-BODY-DIRECT-INDENT", f"Paragraph {idx} overrides body indentation", "body", paragraph=idx))

    max_width = int(rules["table"]["max_width_twips"])
    min_font = int(rules["table"]["font_pt"] * 2)

    for t_idx, tbl in enumerate(document.findall(".//w:tbl", NS), 1):
        tblpr = tbl.find("w:tblPr", NS)
        tblw = tblpr.find("w:tblW", NS) if tblpr is not None else None
        layout = tblpr.find("w:tblLayout", NS) if tblpr is not None else None
        tblind = tblpr.find("w:tblInd", NS) if tblpr is not None else None

        if w_attr(tblw, "type") != "dxa":
            issues.append(issue("TBL-WIDTH-TYPE", f"Table {t_idx} width type is not dxa", "table", table=t_idx))
        else:
            try:
                if int(w_attr(tblw, "w") or "0") > max_width:
                    issues.append(issue("TBL-WIDTH", f"Table {t_idx} exceeds {max_width} twips", "table", table=t_idx))
            except ValueError:
                issues.append(issue("TBL-WIDTH", f"Table {t_idx} has invalid width", "table", table=t_idx))

        if w_attr(layout, "type") != rules["table"]["layout"]:
            issues.append(issue("TBL-LAYOUT", f"Table {t_idx} layout is not {rules['table']['layout']}", "table", table=t_idx))

        if tblind is not None and w_attr(tblind, "w") not in {None, "0"}:
            issues.append(issue("TBL-INDENT", f"Table {t_idx} has nonzero table indentation", "table", table=t_idx))

        grid = tbl.find("w:tblGrid", NS)
        if grid is not None:
            total = 0
            for col in grid.findall("w:gridCol", NS):
                try:
                    total += int(w_attr(col, "w") or "0")
                except ValueError:
                    pass
            if total > max_width:
                issues.append(issue("TBL-GRID-WIDTH", f"Table {t_idx} grid width is {total}", "table", table=t_idx, width=total))

        for c_idx, tc in enumerate(tbl.findall(".//w:tc", NS), 1):
            # A vMerge continuation is not an independently rendered cell. Its hidden
            # paragraph formatting must not fail the visible-table validation.
            if is_vmerge_continuation(tc):
                continue

            tcpr = tc.find("w:tcPr", NS)
            valign = tcpr.find("w:vAlign", NS) if tcpr is not None else None
            if w_attr(valign, "val") != rules["table"]["vertical_alignment"]:
                issues.append(issue("TBL-VERTICAL-ALIGN", f"Table {t_idx} cell {c_idx} is not vertically centered", "table", table=t_idx, cell=c_idx))

            for p_idx, p in enumerate(tc.findall("w:p", NS), 1):
                ppr = p.find("w:pPr", NS)
                jc = ppr.find("w:jc", NS) if ppr is not None else None
                ind = ppr.find("w:ind", NS) if ppr is not None else None

                if w_attr(jc, "val") != rules["table"]["horizontal_alignment"]:
                    issues.append(issue("TBL-HORIZONTAL-ALIGN", f"Table {t_idx} cell {c_idx} paragraph {p_idx} is not centered", "table", table=t_idx, cell=c_idx, paragraph=p_idx))

                if rules["table"]["zero_indentation"]:
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
                        issues.append(issue("TBL-PARA-INDENT", f"Table {t_idx} cell {c_idx} paragraph {p_idx} has indentation", "table", table=t_idx, cell=c_idx, paragraph=p_idx))

                for r_idx, run in enumerate(p.findall(".//w:r", NS), 1):
                    if not "".join((t.text or "") for t in run.findall(".//w:t", NS)):
                        continue
                    rpr = run.find("w:rPr", NS)
                    sz = w_attr(rpr.find("w:sz", NS), "val") if rpr is not None else None
                    szcs = w_attr(rpr.find("w:szCs", NS), "val") if rpr is not None else None
                    values = [v for v in (sz, szcs) if v is not None]
                    for value in values:
                        try:
                            if int(value) < min_font:
                                issues.append(issue("TBL-FONT", f"Table {t_idx} cell {c_idx} paragraph {p_idx} has text below {rules['table']['font_pt']} pt", "table", table=t_idx, cell=c_idx, paragraph=p_idx, run=r_idx))
                                break
                        except ValueError:
                            issues.append(issue("TBL-FONT", f"Table {t_idx} cell {c_idx} paragraph {p_idx} has invalid font size", "table", table=t_idx, cell=c_idx, paragraph=p_idx, run=r_idx))
                            break

    return issues, warnings


def strict_issues(pkg: Package) -> list[dict]:
    issues = []

    bad = pkg.zf.testzip()
    if bad:
        issues.append(issue("STRICT-ZIP", f"Corrupt ZIP member: {bad}"))

    for name in pkg.names:
        if not name.endswith((".xml", ".rels")):
            continue
        data = pkg.zf.read(name)
        try:
            root = ET.fromstring(data)
        except ET.ParseError as exc:
            issues.append(issue("STRICT-XML", f"{name}: {exc}", part=name))
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
                    issues.append(issue("STRICT-IGNORABLE", f"{name} references undeclared prefix {prefix!r}", part=name, prefix=prefix))

    try:
        from docx import Document
        Document(pkg.path)
    except Exception as exc:
        issues.append(issue("STRICT-REOPEN", f"python-docx could not reopen file: {exc}"))

    return issues


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only Word validator.")
    parser.add_argument("document", type=Path)
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES)
    parser.add_argument("--template", type=Path)
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--report", type=Path, help="write structured JSON report")
    parser.add_argument("--json", action="store_true", help="also print full JSON")
    args = parser.parse_args()

    rules = load_json(args.rules)
    template_path = args.template or ROOT / rules["template_path"]

    issues = []
    warnings = []
    target = template = None

    try:
        target = Package(args.document)
        template = Package(template_path)
        q, w = quick_issues(target, template, rules)
        issues.extend(q)
        warnings.extend(w)
        if args.strict:
            issues.extend(strict_issues(target))
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        issues.append(issue("PKG-OPEN", str(exc)))
    finally:
        if target is not None:
            target.close()
        if template is not None:
            template.close()

    result = {
        "ok": not issues,
        "mode": "strict" if args.strict else "quick",
        "document": str(args.document),
        "template": str(template_path),
        "issues": issues,
        "warnings": warnings,
        "auto_fixable": any(x["auto_fixable"] for x in issues),
        "fix_groups": sorted({
            x["fix_group"] for x in issues
            if x["auto_fixable"] and x["fix_group"]
        }),
    }

    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(
            json.dumps(result, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    elif result["ok"]:
        print(f"PASS [{result['mode']}]")
    else:
        print(f"FAIL [{result['mode']}]: {len(issues)} issue(s)")
        for item in issues:
            repair = f" -> {item['fix_group']}" if item["auto_fixable"] else ""
            print(f"  - {item['code']}{repair}: {item['message']}")

    for warning in warnings:
        print(f"WARNING: {warning}")

    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
