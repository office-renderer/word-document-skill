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
