#!/usr/bin/env python3
"""Apply Word formatting rules and repair validator-reported issues."""

from __future__ import annotations

import argparse
import json
import os
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_RULES = ROOT / "config" / "rules.json"
TWIP_EMU = 635

ALL_GROUPS = {"body", "pagination", "table", "page", "footer", "metadata"}
PARAGRAPH_ALIGNMENTS = {"center": WD_ALIGN_PARAGRAPH.CENTER}
VERTICAL_ALIGNMENTS = {"center": WD_CELL_VERTICAL_ALIGNMENT.CENTER}


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def groups_from_report(path: Path) -> set[str]:
    report = load_json(path)
    groups = {
        item.get("fix_group")
        for item in report.get("issues", [])
        if item.get("auto_fixable") and item.get("fix_group")
    }
    return groups & ALL_GROUPS


def apply_pagination_rules(fmt, rules: dict) -> None:
    for name in rules["pagination"]["disable"]:
        setattr(fmt, name, False)


def clear_direct_pagination_overrides(fmt, rules: dict) -> None:
    for name in rules["pagination"]["disable"]:
        if getattr(fmt, name) is True:
            setattr(fmt, name, False)


def set_indentation_zero(paragraph_or_style) -> None:
    fmt = paragraph_or_style.paragraph_format
    fmt.left_indent = Pt(0)
    fmt.right_indent = Pt(0)
    fmt.first_line_indent = Pt(0)

    if hasattr(paragraph_or_style, "_p"):
        ppr = paragraph_or_style._p.get_or_add_pPr()
    else:
        ppr = paragraph_or_style._element.get_or_add_pPr()

    ind = ppr.get_or_add_ind()
    for name in (
        "left", "right", "firstLine",
        "leftChars", "rightChars", "firstLineChars",
        "start", "end", "startChars", "endChars",
    ):
        ind.set(qn(f"w:{name}"), "0")

    for name in ("hanging", "hangingChars"):
        ind.attrib.pop(qn(f"w:{name}"), None)


def copy_style_format(target_doc, template_doc, name: str):
    try:
        source = template_doc.styles[name]
    except KeyError:
        return None

    try:
        target = target_doc.styles[name]
    except KeyError:
        target = target_doc.styles.add_style(name, source.type)

    source_el = source._element
    target_el = target._element

    for tag in ("w:pPr", "w:rPr"):
        old = target_el.find(qn(tag))
        if old is not None:
            target_el.remove(old)
        new = source_el.find(qn(tag))
        if new is not None:
            target_el.append(deepcopy(new))

    if source.base_style is not None:
        try:
            target.base_style = target_doc.styles[source.base_style.name]
        except KeyError:
            pass

    return target


def iter_table_paragraphs(table):
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for nested in cell.tables:
                yield from iter_table_paragraphs(nested)


def iter_all_paragraphs(doc):
    yield from doc.paragraphs
    for table in doc.tables:
        yield from iter_table_paragraphs(table)
    for section in doc.sections:
        for part in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ):
            yield from part.paragraphs
            for table in part.tables:
                yield from iter_table_paragraphs(table)


def usable_width_twips(doc, configured_max: int) -> int:
    widths = []
    for section in doc.sections:
        if (
            section.page_width is None
            or section.left_margin is None
            or section.right_margin is None
        ):
            continue
        value = int(
            (section.page_width - section.left_margin - section.right_margin)
            / TWIP_EMU
        )
        if value > 0:
            widths.append(value)
    return min([configured_max, *widths]) if widths else configured_max


TBLW_LATER_TAGS = tuple(
    qn(tag) for tag in (
        "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
        "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
        "w:tblDescription", "w:tblPrChange",
    )
)


def ensure_tbl_width(tbl_pr):
    node = tbl_pr.find(qn("w:tblW"))
    if node is not None:
        return node
    node = OxmlElement("w:tblW")
    insert_at = len(tbl_pr)
    for index, child in enumerate(tbl_pr):
        if child.tag in TBLW_LATER_TAGS:
            insert_at = index
            break
    tbl_pr.insert(insert_at, node)
    return node


def normalize_table(table, table_style, rules: dict, max_width: int) -> None:
    table.autofit = rules["table"]["layout"] != "fixed"

    tbl_pr = table._tbl.tblPr
    tbl_w = ensure_tbl_width(tbl_pr)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(max_width))

    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is not None:
        tbl_ind.set(qn("w:type"), "dxa")
        tbl_ind.set(qn("w:w"), "0")

    layout = tbl_pr.get_or_add_tblLayout()
    layout.set(qn("w:type"), rules["table"]["layout"])

    grid_cols = list(table._tbl.tblGrid.gridCol_lst)
    widths = []
    for col in grid_cols:
        try:
            widths.append(int(col.get(qn("w:w")) or "0"))
        except ValueError:
            widths.append(0)

    total = sum(widths)
    if widths and total > max_width and total > 0:
        scaled = []
        used = 0
        for index, width in enumerate(widths):
            if index == len(widths) - 1:
                new_width = max_width - used
            else:
                new_width = max(1, round(width * max_width / total))
                used += new_width
            scaled.append(new_width)
        for col, width in zip(grid_cols, scaled):
            col.set(qn("w:w"), str(width))

    h_align = PARAGRAPH_ALIGNMENTS[rules["table"]["horizontal_alignment"]]
    v_align = VERTICAL_ALIGNMENTS[rules["table"]["vertical_alignment"]]

    seen_cells = set()
    for row in table.rows:
        for cell in row.cells:
            key = id(cell._tc)
            if key in seen_cells:
                continue
            seen_cells.add(key)

            cell.vertical_alignment = v_align
            tcw = cell._tc.get_or_add_tcPr().get_or_add_tcW()
            tcw.set(qn("w:type"), "auto")
            tcw.set(qn("w:w"), "0")

            for paragraph in cell.paragraphs:
                paragraph.style = table_style
                paragraph.alignment = h_align
                if rules["table"]["zero_indentation"]:
                    set_indentation_zero(paragraph)
                clear_direct_pagination_overrides(paragraph.paragraph_format, rules)
                for run in paragraph.runs:
                    if run.text:
                        current = run.font.size.pt if run.font.size is not None else None
                        if current is None or current < rules["table"]["font_pt"]:
                            run.font.size = Pt(rules["table"]["font_pt"])

            for nested in cell.tables:
                normalize_table(nested, table_style, rules, max_width)


def copy_page_setup(doc, template_doc) -> None:
    source = template_doc.sections[-1]
    attrs = (
        "page_width", "page_height", "orientation",
        "top_margin", "bottom_margin", "left_margin", "right_margin",
        "header_distance", "footer_distance", "gutter",
    )
    for section in doc.sections:
        for name in attrs:
            setattr(section, name, getattr(source, name))


def replace_story_content(target_story, source_story) -> None:
    target_el = target_story._element
    for child in list(target_el):
        target_el.remove(child)
    for child in source_story._element:
        target_el.append(deepcopy(child))


def copy_footers(doc, template_doc) -> None:
    doc.settings.odd_and_even_pages_header_footer = (
        template_doc.settings.odd_and_even_pages_header_footer
    )
    source = template_doc.sections[-1]

    for section in doc.sections:
        section.footer.is_linked_to_previous = False
        section.even_page_footer.is_linked_to_previous = False
        replace_story_content(section.footer, source.footer)
        replace_story_content(section.even_page_footer, source.even_page_footer)


def strip_package_metadata(path: Path, rules: dict) -> None:
    remove_parts = set(rules["metadata"]["remove_parts"])
    rel_types = set(rules["metadata"]["relationship_types"])

    fd, temp_name = tempfile.mkstemp(
        prefix=path.stem + ".metadata-",
        suffix=path.suffix,
        dir=path.parent,
    )
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(temp_path, "w") as zout:
            for info in zin.infolist():
                name = info.filename
                if name in remove_parts:
                    continue

                data = zin.read(name)

                if name == "_rels/.rels":
                    root = etree.fromstring(data)
                    for rel in list(root):
                        rel_type = rel.get("Type", "")
                        target = rel.get("Target", "").lstrip("/")
                        if rel_type in rel_types or target in remove_parts:
                            root.remove(rel)
                    data = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )

                elif name == "[Content_Types].xml":
                    root = etree.fromstring(data)
                    for child in list(root):
                        if child.get("PartName", "").lstrip("/") in remove_parts:
                            root.remove(child)
                    data = etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True
                    )

                zout.writestr(info, data)

        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def quick_safety_check(path: Path, rules: dict) -> None:
    with zipfile.ZipFile(path, "r") as zf:
        names = set(zf.namelist())
        for name in rules["validation"]["quick_required_parts"]:
            if name not in names:
                raise RuntimeError(f"missing required OOXML part: {name}")
            if name.endswith((".xml", ".rels")):
                etree.fromstring(zf.read(name))
        for name in rules["metadata"]["remove_parts"]:
            if name in names:
                raise RuntimeError(f"metadata part remains: {name}")

    Document(path)


def format_document(src: Path, dst: Path, rules: dict, groups: set[str]) -> None:
    doc = Document(src)
    template = Document(ROOT / rules["template_path"])

    body_name = rules["body"]["style"]
    body_style = None

    if "body" in groups:
        body_style = copy_style_format(doc, template, body_name)
        if body_style is None:
            raise RuntimeError(f"template is missing required body style: {body_name}")
        body_style.paragraph_format.first_line_indent = Pt(
            rules["body"]["first_line_twips"] / 20
        )

        auto_from = set(rules["body"]["auto_map_from_styles"])
        min_chars = int(rules["body"]["auto_map_min_chars"])
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            style_name = paragraph.style.name if paragraph.style is not None else "Normal"
            if (
                len(text) >= min_chars
                and style_name in auto_from | {body_name}
                and paragraph.alignment
                not in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT)
            ):
                paragraph.style = body_style
                paragraph.paragraph_format.left_indent = None
                paragraph.paragraph_format.right_indent = None
                paragraph.paragraph_format.first_line_indent = None

    if "pagination" in groups:
        for style in doc.styles:
            if style.type == WD_STYLE_TYPE.PARAGRAPH:
                apply_pagination_rules(style.paragraph_format, rules)
        for paragraph in iter_all_paragraphs(doc):
            clear_direct_pagination_overrides(paragraph.paragraph_format, rules)

    if "table" in groups:
        if body_style is None:
            try:
                body_style = doc.styles[body_name]
            except KeyError:
                body_style = copy_style_format(doc, template, body_name)

        table_name = rules["table"]["paragraph_style"]
        try:
            table_style = doc.styles[table_name]
        except KeyError:
            table_style = doc.styles.add_style(table_name, WD_STYLE_TYPE.PARAGRAPH)
            table_style.base_style = body_style

        table_style.font.size = Pt(rules["table"]["font_pt"])
        table_style.paragraph_format.alignment = PARAGRAPH_ALIGNMENTS[
            rules["table"]["horizontal_alignment"]
        ]
        if rules["table"]["zero_indentation"]:
            set_indentation_zero(table_style)
        apply_pagination_rules(table_style.paragraph_format, rules)

        max_width = usable_width_twips(doc, int(rules["table"]["max_width_twips"]))
        for table in doc.tables:
            normalize_table(table, table_style, rules, max_width)

    if "page" in groups:
        copy_page_setup(doc, template)

    if "footer" in groups:
        copy_footers(doc, template)

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst)

    if "metadata" in groups or not groups:
        strip_package_metadata(dst, rules)
    else:
        # python-docx may create core properties on save; final files must always be clean.
        strip_package_metadata(dst, rules)

    quick_safety_check(dst, rules)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Apply Word formatting or targeted fixes from a validator report."
    )
    parser.add_argument("document", type=Path)
    parser.add_argument("--rules", type=Path, default=DEFAULT_RULES)
    parser.add_argument("--fix", type=Path, help="validator JSON report for targeted repair")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--in-place", action="store_true")
    group.add_argument("--out", type=Path)
    args = parser.parse_args()

    rules = load_json(args.rules)
    groups = groups_from_report(args.fix) if args.fix else set(ALL_GROUPS)

    if args.fix and not groups:
        print("No auto-fixable formatting groups found in validator report.")
        return 2

    if args.in_place:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir) / args.document.name
            format_document(args.document, temp_path, rules, groups)
            os.replace(temp_path, args.document)
        target = args.document
    else:
        format_document(args.document, args.out, rules, groups)
        target = args.out

    print("Formatted:", target)
    print("Applied groups:", ", ".join(sorted(groups)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
