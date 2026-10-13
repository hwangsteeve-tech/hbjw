"""저장 직전에 PPT 전체에 적용하는 공통 스타일.

- apply_font: 템플릿·빌더가 쓰는 KoPubWorld 글꼴을 어느 PC에나 있는 '맑은 고딕'으로 바꾼다.
  (Bold 글꼴은 굵게 표시로 옮기고, 좁은 자간(spc 음수)은 없앤다)
- mark_blanks: 작성자가 채워야 할 '【 … 】' 부분을 붉은색 굵은 글씨로 표시한다.
"""
from __future__ import annotations

import copy
import re

from lxml import etree
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

BODY_FONT = "맑은 고딕"
BLANK_COLOR = "E00000"
_OLD = "KoPubWorld"
_BLANK = re.compile(r"(【[^】]*】)")


def _fix_tree(root):
    for tag in ("a:latin", "a:ea", "a:cs"):
        for el in root.iter(qn(tag)):
            face = el.get("typeface", "")
            if _OLD in face:
                el.set("typeface", BODY_FONT)
                if "Bold" in face:
                    el.getparent().set("b", "1")
    for tag in ("a:rPr", "a:defRPr", "a:endParaRPr"):
        for el in root.iter(qn(tag)):
            spc = el.get("spc")
            if spc and spc.lstrip("-").isdigit() and int(spc) < 0:
                del el.attrib["spc"]


def apply_font(prs):
    for slide in prs.slides:
        _fix_tree(slide._element)
    for master in prs.slide_masters:
        _fix_tree(master._element)
        for layout in master.slide_layouts:
            _fix_tree(layout._element)
        theme = master.part.part_related_by(RT.THEME)
        root = etree.fromstring(theme.blob)
        for font in root.iter(qn("a:latin"), qn("a:ea")):
            if font.getparent().tag in (qn("a:majorFont"), qn("a:minorFont")):
                font.set("typeface", BODY_FONT)
        theme._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _split_run(r):
    text = r.findtext(qn("a:t")) or ""
    parts = [p for p in _BLANK.split(text) if p]
    if len(parts) == 1 and not _BLANK.fullmatch(parts[0]):
        return
    parent = r.getparent()
    idx = list(parent).index(r)
    for k, part in enumerate(parts):
        nr = copy.deepcopy(r)
        nr.find(qn("a:t")).text = part
        if _BLANK.fullmatch(part):
            rpr = nr.find(qn("a:rPr"))
            if rpr is None:
                rpr = etree.SubElement(nr, qn("a:rPr"))
                nr.remove(rpr)
                nr.insert(0, rpr)
            rpr.set("b", "1")
            for f in rpr.findall(qn("a:solidFill")):
                rpr.remove(f)
            sf = etree.Element(qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", BLANK_COLOR)
            # solidFill 은 ln 다음, latin/ea 앞에 와야 한다
            ln = rpr.find(qn("a:ln"))
            rpr.insert(list(rpr).index(ln) + 1 if ln is not None else 0, sf)
        parent.insert(idx + k, nr)
    parent.remove(r)


def mark_blanks(prs):
    for slide in prs.slides:
        for r in list(slide._element.iter(qn("a:r"))):
            if "【" in (r.findtext(qn("a:t")) or ""):
                _split_run(r)


def finalize(prs):
    apply_font(prs)
    mark_blanks(prs)
