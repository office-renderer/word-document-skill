#!/usr/bin/env python3
"""Normalize body indentation and Word paragraph pagination options."""

from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)

PAGINATION = ("widowControl", "keepNext", "keepLines", "pageBreakBefore")
BODY_STYLE = "正文（默认）"
NORMAL_STYLE = "Normal"
BODY_FIRST_LINE = "640"
TABLE_MAX_WIDTH = 8845
TABLE_FONT_HALF_POINTS = "28"


def wa(el, name):
    return None if el is None else el.get(W + name)


def false_flag(ppr, tag):
    node = ppr.find(f"w:{tag}", NS)
    if node is None:
        node = ET.Element(W + tag)
        order = {
            "pStyle": 0,
            "keepNext": 1,
            "keepLines": 2,
            "pageBreakBefore": 3,
            "framePr": 4,
            "widowControl": 5,
            "numPr": 6,
            "spacing": 20,
            "ind": 21,
            "jc": 25,
            "outlineLvl": 29,
        }
        target = order[tag]
        pos = len(ppr)
        for i, child in enumerate(list(ppr)):
            child_name = child.tag.rsplit("}", 1)[-1]
            if order.get(child_name, 99) > target:
                pos = i
                break
        ppr.insert(pos, node)
    node.set(W + "val", "0")


def ensure_ppr(parent):
    ppr = parent.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(W + "pPr")
        rpr = parent.find("w:rPr", NS)
        if rpr is not None:
            parent.insert(list(parent).index(rpr), ppr)
        else:
            parent.append(ppr)
    return ppr


def maps(styles):
    by_id, by_name = {}, {}
    for s in styles.findall("w:style", NS):
        sid, name = wa(s, "styleId"), wa(s.find("w:name", NS), "val")
        if sid:
            by_id[sid] = s
        if name:
            by_name[name] = s
    return by_id, by_name


def text_of(p):
    return "".join((t.text or "") for t in p.findall(".//w:t", NS)).strip()



PPR_ORDER = {
    "pStyle": 0, "keepNext": 1, "keepLines": 2, "pageBreakBefore": 3,
    "framePr": 4, "widowControl": 5, "numPr": 6, "spacing": 20,
    "ind": 21, "jc": 25, "rPr": 27, "outlineLvl": 29,
}
TBLPR_ORDER = {
    "tblStyle": 0, "tblpPr": 1, "tblOverlap": 2, "bidiVisual": 3,
    "tblStyleRowBandSize": 4, "tblStyleColBandSize": 5, "tblW": 6,
    "jc": 7, "tblCellSpacing": 8, "tblInd": 9, "tblBorders": 10,
    "shd": 11, "tblLayout": 12, "tblCellMar": 13, "tblLook": 14,
}
MARGIN_ORDER = {"top": 0, "start": 1, "left": 1, "bottom": 2, "end": 3, "right": 3}
TCPR_ORDER = {
    "cnfStyle": 0, "tcW": 1, "gridSpan": 2, "hMerge": 3, "vMerge": 4,
    "tcBorders": 5, "shd": 6, "noWrap": 7, "tcMar": 8,
    "textDirection": 9, "tcFitText": 10, "vAlign": 11, "hideMark": 12,
}
RPR_ORDER = {
    "rStyle": 0, "rFonts": 1, "b": 2, "bCs": 3, "i": 4, "iCs": 5,
    "caps": 6, "smallCaps": 7, "strike": 8, "dstrike": 9,
    "outline": 10, "shadow": 11, "emboss": 12, "imprint": 13,
    "noProof": 14, "snapToGrid": 15, "vanish": 16, "webHidden": 17,
    "color": 18, "spacing": 19, "w": 20, "kern": 21, "position": 22,
    "sz": 23, "szCs": 24, "highlight": 25, "u": 26,
}


def ensure_ordered(parent, tag, order):
    node = parent.find(f"w:{tag}", NS)
    if node is not None:
        return node
    node = ET.Element(W + tag)
    target = order.get(tag, 999)
    pos = len(parent)
    for i, child in enumerate(list(parent)):
        name = child.tag.rsplit("}", 1)[-1]
        if order.get(name, 999) > target:
            pos = i
            break
    parent.insert(pos, node)
    return node


def ensure_rpr(run):
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(W + "rPr")
        run.insert(0, rpr)
    return rpr


def set_run_size(run, half_points):
    rpr = ensure_rpr(run)
    ensure_ordered(rpr, "sz", RPR_ORDER).set(W + "val", half_points)
    ensure_ordered(rpr, "szCs", RPR_ORDER).set(W + "val", half_points)


def set_margin_zero(margins, tag):
    node = margins.find(f"w:{tag}", NS)
    if node is None:
        node = ET.Element(W + tag)
        target = MARGIN_ORDER.get(tag, 99)
        pos = len(margins)
        for i, child in enumerate(list(margins)):
            name = child.tag.rsplit("}", 1)[-1]
            if MARGIN_ORDER.get(name, 99) > target:
                pos = i
                break
        margins.insert(pos, node)
    node.set(W + "w", "0")
    node.set(W + "type", "dxa")


def normalize_tables(doc):
    for tbl in doc.findall(".//w:tbl", NS):
        tblpr = tbl.find("w:tblPr", NS)
        if tblpr is None:
            tblpr = ET.Element(W + "tblPr")
            tbl.insert(0, tblpr)

        tblind = ensure_ordered(tblpr, "tblInd", TBLPR_ORDER)
        tblind.set(W + "w", "0")
        tblind.set(W + "type", "dxa")

        layout = ensure_ordered(tblpr, "tblLayout", TBLPR_ORDER)
        layout.set(W + "type", "fixed")

        tbl_cell_mar = ensure_ordered(tblpr, "tblCellMar", TBLPR_ORDER)
        set_margin_zero(tbl_cell_mar, "left")
        set_margin_zero(tbl_cell_mar, "right")

        grid = tbl.find("w:tblGrid", NS)
        widths = []
        if grid is not None:
            for col in grid.findall("w:gridCol", NS):
                try:
                    widths.append(int(wa(col, "w") or "0"))
                except ValueError:
                    widths.append(0)

        total = sum(widths)
        if widths and total > TABLE_MAX_WIDTH:
            scaled = []
            used = 0
            for i, width in enumerate(widths):
                if i == len(widths) - 1:
                    new_width = TABLE_MAX_WIDTH - used
                else:
                    new_width = max(1, round(width * TABLE_MAX_WIDTH / total))
                    used += new_width
                scaled.append(new_width)
            for col, width in zip(grid.findall("w:gridCol", NS), scaled):
                col.set(W + "w", str(width))
            target_width = sum(scaled)
        elif widths and total > 0:
            target_width = total
        else:
            target_width = TABLE_MAX_WIDTH

        tblw = ensure_ordered(tblpr, "tblW", TBLPR_ORDER)
        tblw.set(W + "w", str(min(target_width, TABLE_MAX_WIDTH)))
        tblw.set(W + "type", "dxa")

        for tc in tbl.findall(".//w:tc", NS):
            tcpr = tc.find("w:tcPr", NS)
            if tcpr is None:
                tcpr = ET.Element(W + "tcPr")
                tc.insert(0, tcpr)
            tcw = ensure_ordered(tcpr, "tcW", TCPR_ORDER)
            tcw.set(W + "w", "0")
            tcw.set(W + "type", "auto")

            tc_mar = ensure_ordered(tcpr, "tcMar", TCPR_ORDER)
            set_margin_zero(tc_mar, "left")
            set_margin_zero(tc_mar, "right")

            valign = ensure_ordered(tcpr, "vAlign", TCPR_ORDER)
            valign.set(W + "val", "center")

            for p in tc.findall("w:p", NS):
                ppr = p.find("w:pPr", NS)
                if ppr is None:
                    ppr = ET.Element(W + "pPr")
                    p.insert(0, ppr)

                ind = ensure_ordered(ppr, "ind", PPR_ORDER)
                for key in (
                    "left", "right", "firstLine", "hanging",
                    "leftChars", "rightChars", "firstLineChars", "hangingChars",
                    "start", "end", "startChars", "endChars",
                ):
                    ind.set(W + key, "0")

                jc = ensure_ordered(ppr, "jc", PPR_ORDER)
                jc.set(W + "val", "center")

                for run in p.findall("w:r", NS):
                    if "".join((t.text or "") for t in run.findall(".//w:t", NS)):
                        set_run_size(run, TABLE_FONT_HALF_POINTS)


def normalize(src: Path, dst: Path):
    with zipfile.ZipFile(src) as zin:
        styles = ET.fromstring(zin.read("word/styles.xml"))
        doc = ET.fromstring(zin.read("word/document.xml"))
        by_id, by_name = maps(styles)

        body_style = by_name.get(BODY_STYLE)
        if body_style is None:
            raise ValueError(f"missing required style: {BODY_STYLE}")
        body_id = wa(body_style, "styleId")

        # Explicitly turn all four pagination checkboxes off for every paragraph style.
        for s in styles.findall("w:style", NS):
            if wa(s, "type") != "paragraph":
                continue
            ppr = ensure_ppr(s)
            for tag in PAGINATION:
                false_flag(ppr, tag)

        # Keep the body style's canonical two-character first-line indent.
        ppr = ensure_ppr(body_style)
        ind = ppr.find("w:ind", NS)
        if ind is None:
            ind = ET.SubElement(ppr, W + "ind")
        ind.set(W + "firstLine", BODY_FIRST_LINE)
        ind.attrib.pop(W + "hanging", None)
        ind.attrib.pop(W + "hangingChars", None)

        # A direct paragraph setting must never turn pagination back on.
        for p in doc.findall(".//w:p", NS):
            ppr = p.find("w:pPr", NS)
            if ppr is None:
                continue
            for tag in PAGINATION:
                if ppr.find(f"w:{tag}", NS) is not None:
                    false_flag(ppr, tag)

        normalize_tables(doc)

        # Direct w:body children exclude table-cell paragraphs. Normalize ordinary prose.
        body = doc.find("w:body", NS)
        if body is not None:
            for p in body.findall("w:p", NS):
                txt = text_of(p)
                if len(txt) < 12:
                    continue
                ppr = p.find("w:pPr", NS)
                jc = ppr.find("w:jc", NS) if ppr is not None else None
                if wa(jc, "val") in {"center", "right"}:
                    continue
                pstyle = ppr.find("w:pStyle", NS) if ppr is not None else None
                sid = wa(pstyle, "val")
                name = wa(by_id[sid].find("w:name", NS), "val") if sid in by_id else (NORMAL_STYLE if sid is None else None)
                if name not in {NORMAL_STYLE, BODY_STYLE}:
                    continue
                if ppr is None:
                    ppr = ET.Element(W + "pPr")
                    p.insert(0, ppr)
                if pstyle is None:
                    pstyle = ET.Element(W + "pStyle")
                    ppr.insert(0, pstyle)
                pstyle.set(W + "val", body_id)
                direct_ind = ppr.find("w:ind", NS)
                if direct_ind is not None:
                    for key in ("firstLine", "firstLineChars", "hanging", "hangingChars"):
                        direct_ind.attrib.pop(W + key, None)
                    if not direct_ind.attrib:
                        ppr.remove(direct_ind)

        replacements = {
            "word/styles.xml": ET.tostring(styles, encoding="utf-8", xml_declaration=True),
            "word/document.xml": ET.tostring(doc, encoding="utf-8", xml_declaration=True),
        }
        with zipfile.ZipFile(dst, "w") as zout:
            for info in zin.infolist():
                zout.writestr(info, replacements.get(info.filename, zin.read(info.filename)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("document", type=Path)
    group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument("--in-place", action="store_true")
    group.add_argument("--out", type=Path)
    args = ap.parse_args()

    if args.in_place:
        with tempfile.TemporaryDirectory() as td:
            tmp = Path(td) / args.document.name
            normalize(args.document, tmp)
            shutil.copy2(tmp, args.document)
        target = args.document
    else:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        normalize(args.document, args.out)
        target = args.out

    print(f"Normalized: {target}")


if __name__ == "__main__":
    main()
