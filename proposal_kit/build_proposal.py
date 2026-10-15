"""H9 정성제안서 PPT 생성기.

YAML 내용 파일을 읽어 templates/h9_base.pptx 양식(세로 A4, KoPubWorld돋움체)으로
표지 · 요구사항 분석표 · 목차 · 간지 · 본문 슬라이드를 만든다.

    python -m proposal_kit.build_proposal proposals/xxx/content.yaml -o out.pptx

본문 슬라이드는 위에서 아래로 블록을 쌓는다. 블록 종류는 README.md 참고.
"""
from __future__ import annotations

import argparse
import copy
import re
import sys
from pathlib import Path

import yaml
from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Cm, Emu, Pt

from proposal_kit.company import load_content, set_pptx_author
from proposal_kit.style import finalize

HERE = Path(__file__).resolve().parent
BASE_TEMPLATE = HERE / "templates" / "h9_base.pptx"

# 템플릿 프로토타입 슬라이드 순서 (h9_base.pptx)
PROTO_COVER, PROTO_REQ, PROTO_TOC, PROTO_DIVIDER, PROTO_BODY = range(5)

FONT_B = "KoPubWorld돋움체 Bold"
FONT_M = "KoPubWorld돋움체 Medium"

BLUE = "0462DA"       # 표 머리, 소제목 막대, 강조 글자
NAVY = "1D4E89"       # 본문 테두리
DARK = "194EA7"       # 흐름도 글자
LIGHT = "D7EAF2"      # 흐름도 상자
LIGHT_LINE = "BADBE8"
GRID = "C8D3E0"

# 본문 영역 (cm)
X0, W = 0.78, 17.59
Y_START, Y_END = 6.55, 26.75

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


# ---------------------------------------------------------------- 텍스트 유틸

_MARK = re.compile(r"(\*\*.+?\*\*)")


def _runs(text: str):
    """'**강조**' 표기를 (문자열, 강조여부) 목록으로 나눈다."""
    out = []
    for part in _MARK.split(str(text)):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            out.append((part[2:-2], True))
        else:
            out.append((part, False))
    return out


def _set_font(run, size, font=FONT_M, color="000000", bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = _rgb(color)
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)
    rPr.set("spc", "-30")


def _rgb(hexstr):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(hexstr)


def _add_runs(par, text, size, font=FONT_M, color="000000", bold=False, em_color=BLUE):
    for chunk, em in _runs(text):
        r = par.add_run()
        r.text = chunk
        _set_font(r, size, FONT_B if (em or bold) else font, em_color if em else color, bold)


def _para_fmt(par, line=1.2, before=0, after=0, align=None):
    par.line_spacing = line
    par.space_before = Pt(before)
    par.space_after = Pt(after)
    if align:
        from pptx.enum.text import PP_ALIGN
        par.alignment = {"c": PP_ALIGN.CENTER, "l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT}[align]


def _bullet(par, char, indent_emu, mar_emu, font="Arial"):
    pPr = par._p.get_or_add_pPr()
    pPr.set("marL", str(mar_emu))
    pPr.set("indent", str(-indent_emu))
    for tag in ("a:buNone", "a:buChar", "a:buFont"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    bf = etree.SubElement(pPr, qn("a:buFont"))
    bf.set("typeface", font)
    bc = etree.SubElement(pPr, qn("a:buChar"))
    bc.set("char", char)


def _no_bullet(par):
    pPr = par._p.get_or_add_pPr()
    if pPr.find(qn("a:buNone")) is None:
        etree.SubElement(pPr, qn("a:buNone"))


def _text_box(slide, x, y, w, h, wrap=True, inset=0.0, anchor="t", autofit=False):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Cm(inset))
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT if autofit else MSO_AUTO_SIZE.NONE
    return tb


def _shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=1.0, adj=None, shadow=False):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(getattr(MSO_SHAPE, kind), Cm(x), Cm(y), Cm(w), Cm(h))
    if adj is not None:
        shp.adjustments[0] = adj
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    if not shadow:
        spPr = shp._element.spPr
        if spPr.find(qn("a:effectLst")) is None:
            etree.SubElement(spPr, qn("a:effectLst"))
    # 기본 스타일 글꼴색 제거용
    shp.text_frame.word_wrap = True
    return shp


def _fill_shape_text(shp, text, size, font=FONT_B, color="000000", align="c", line=1.15, inset=0.12):
    tf = shp.text_frame
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Cm(inset))
    lines = text if isinstance(text, list) else str(text).split("\n")
    for i, ln in enumerate(lines):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _para_fmt(par, line=line, align=align)
        _add_runs(par, ln, size, font, color)


# 한글 줄 수 추정: 글자 폭 ≈ 글꼴 크기의 0.93배
def _est_lines(text, size_pt, width_cm, indent_cm=0.0):
    plain = re.sub(r"\*\*", "", str(text))
    char_cm = size_pt * 0.0353 * 1.0
    cap = max(1, int((width_cm - indent_cm) / char_cm))
    units = sum(1.0 if ord(c) > 0x2000 else 0.55 for c in plain)
    return max(1, -(-int(units) // cap))


def _line_cm(size_pt, spacing=1.2):
    return size_pt * 0.0353 * spacing


# ---------------------------------------------------------------- 슬라이드 복제

def _duplicate(prs, index):
    src = prs.slides[index]
    dst = prs.slides.add_slide(src.slide_layout)
    for shp in list(dst.shapes):
        shp._element.getparent().remove(shp._element)
    for el in src.shapes._spTree.iterchildren():
        tag = etree.QName(el).localname
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        dst.shapes._spTree.append(copy.deepcopy(el))
    for rel in src.part.rels.values():
        if "notesSlide" in rel.reltype or "slideLayout" in rel.reltype:
            continue
        dst.part.rels.get_or_add(rel.reltype, rel._target)
    return dst


def _shape_by_name(slide, name):
    for shp in slide.shapes:
        if shp.name == name:
            return shp
    raise KeyError(name)


def _replace_text(shp, lines, keep_first_run=True):
    """첫 run의 서식을 유지한 채 문단 내용을 교체한다. lines: str 또는 list."""
    if isinstance(lines, str):
        lines = lines.split("\n")
    txBody = shp.text_frame._txBody
    paras = txBody.findall(qn("a:p"))
    proto_p = paras[0]
    proto_r = proto_p.find(qn("a:r"))
    rPr = copy.deepcopy(proto_r.find(qn("a:rPr"))) if proto_r is not None else None
    pPr = proto_p.find(qn("a:pPr"))
    for p_ in paras:
        txBody.remove(p_)
    for ln in lines:
        p_ = etree.SubElement(txBody, qn("a:p"))
        if pPr is not None:
            p_.append(copy.deepcopy(pPr))
        r = etree.SubElement(p_, qn("a:r"))
        if rPr is not None:
            r.append(copy.deepcopy(rPr))
        t = etree.SubElement(r, qn("a:t"))
        t.text = ln


def _delete_slide(prs, index):
    sldIdLst = prs.slides._sldIdLst
    sldId = list(sldIdLst)[index]
    prs.part.drop_rel(sldId.rId)
    sldIdLst.remove(sldId)


# ---------------------------------------------------------------- 고정 슬라이드

def build_cover(prs, meta):
    s = _duplicate(prs, PROTO_COVER)
    title = _shape_by_name(s, "TextBox 5")
    lines = [meta.get("doc_type", "정성제안서")] + list(meta["title_lines"])
    # 첫 줄(문서 종류)과 제목 줄은 서식이 달라서 문단을 각각 복제해 채운다
    txBody = title.text_frame._txBody
    paras = txBody.findall(qn("a:p"))
    proto_small, proto_big = paras[0], paras[1] if len(paras) > 1 else paras[0]
    for p_ in paras:
        txBody.remove(p_)
    for i, ln in enumerate(lines):
        src = proto_small if i == 0 else proto_big
        p_ = copy.deepcopy(src)
        runs = p_.findall(qn("a:r"))
        for r in runs[1:]:
            p_.remove(r)
        runs[0].find(qn("a:t")).text = ln
        txBody.append(p_)
    _replace_text(_shape_by_name(s, "TextBox 6"), meta.get("pm_line", "사업PM ○○○ 파트너"))
    orgs = meta.get("org_lines", [])
    if orgs:
        _replace_text(_shape_by_name(s, "TextBox 7"), orgs)
    return s


def build_toc(prs, chapters):
    s = _duplicate(prs, PROTO_TOC)
    group = [["TextBox 2", "타원 42", "직선 화살표 연결선 44", "TextBox 52", "직선 연결선 70"]]
    proto = [_shape_by_name(s, n)._element for n in group[0]]
    # 기존 4개 항목 제거 후 장 수만큼 다시 배치
    keep = {"TextBox 16"} | set(group[0])
    for shp in list(s.shapes):
        if shp.name not in keep:
            shp._element.getparent().remove(shp._element)
    step = Cm(1.93) if len(chapters) <= 5 else Cm(min(1.93, 12.5 / len(chapters)))
    tree = s.shapes._spTree
    for el in proto:
        tree.remove(el)
    for i, ch in enumerate(chapters):
        for el in proto:
            e = copy.deepcopy(el)
            off = e.find(".//" + qn("a:off"))
            off.set("y", str(int(off.get("y")) + int(step) * i))
            tree.append(e)
        added = list(tree.iterchildren())[-5:]
        label, roman = added[0], added[3]
        _set_xml_text(label, ch["title"])
        _set_xml_text(roman, ch["num"])
    return s


def _set_xml_text(sp_el, text):
    txBody = sp_el.find(qn("p:txBody"))
    paras = txBody.findall(qn("a:p"))
    first = paras[0]
    runs = first.findall(qn("a:r"))
    for p_ in paras[1:]:
        txBody.remove(p_)
    for r in runs[1:]:
        first.remove(r)
    runs[0].find(qn("a:t")).text = text


def build_divider(prs, ch):
    s = _duplicate(prs, PROTO_DIVIDER)
    title = _shape_by_name(s, "TextBox 4")
    _replace_text(title, ch["title"])
    # 폭이 넓은 로마숫자는 제목을 오른쪽으로 민다
    shift = {"Ⅳ": 1.6, "Ⅵ": 1.6, "Ⅶ": 2.4, "Ⅷ": 3.2, "Ⅸ": 1.6}.get(ch["num"], 0)
    if shift:
        title.left += Cm(shift)
        title.width -= Cm(shift)
    _replace_text(_shape_by_name(s, "TextBox 17"), f"{ch['num']}.")
    items = ch.get("toc") or []
    _replace_text(_shape_by_name(s, "TextBox 3"), items or [" "])
    return s


def _compact_pages(pages):
    """['Ⅲ-3', 'Ⅲ-4', 'Ⅲ-5', 'Ⅳ-1'] → 'Ⅲ-3~5, Ⅳ-1' (숫자 쪽번호도 같은 방식)."""
    groups: dict = {}
    for p in pages:
        ch, _, n = str(p).rpartition("-")
        if not n.isdigit():
            groups.setdefault(str(p), [])
            continue
        groups.setdefault(ch, []).append(int(n))
    out = []
    for ch, nums in groups.items():
        if not nums:
            out.append(ch)
            continue
        nums = sorted(set(nums))
        pre = f"{ch}-" if ch else ""
        start = prev = nums[0]
        for n in nums[1:] + [None]:
            if n is not None and n == prev + 1:
                prev = n
                continue
            out.append(f"{pre}{start}" + (f"~{prev}" if prev != start else ""))
            if n is not None:
                start = prev = n
    return ", ".join(out)


def build_req_table(prs, slide, req_index, title=None, compliance=None):
    """요구사항 분석표 슬라이드: 기존 표를 지우고 분류 · ID · 페이지 표를 새로 만든다.

    compliance({ID: {name, comply}})가 있으면 '요구사항 대비 수용여부 조견표'로 만들어
    분류 · ID · 명칭 · 수용 여부 · 해당 페이지를 보인다.
    """
    for shp in list(slide.shapes):
        if shp.has_table:
            shp._element.getparent().remove(shp._element)
    if compliance is not None:
        title = title or "< 요구사항 대비 수용여부 조견표 >"
        rows = [("분류", "고유번호", "요구사항 명칭", "수용 여부", "해당 페이지")]
        for cat, items in req_index:
            for rid, pages in items:
                info = compliance.get(rid) or {}
                rows.append((cat, rid, info.get("name", "【명칭】"), info.get("comply", "수용"),
                             _compact_pages(pages) if pages else "-"))
        widths = [2.3, 2.1, 7.2, 1.7, 3.0]
    else:
        rows = [("요구사항 분류", "ID", "해당 페이지")]
        for cat, items in req_index:
            for rid, pages in items:
                rows.append((cat, rid, ", ".join(str(p) for p in pages) if pages else "-"))
        widths = [5.4, 5.4, 5.5]
    if title:
        for shp in slide.shapes:
            if shp.has_text_frame and "요구사항" in shp.text_frame.text:
                _replace_text(shp, title)
    # 한 장에 최대 45행. 넘치면 두 장으로 나눈다.
    body = rows[1:]
    n_pages = max(1, -(-len(body) // 44))
    per = -(-len(body) // n_pages) if body else 44  # 여러 장이면 고르게 나눈다
    chunks = [body[i:i + per] for i in range(0, max(1, len(body)), per)]
    slides = [slide]
    for _ in chunks[1:]:
        extra = _duplicate(prs, PROTO_REQ)
        for shp in list(extra.shapes):
            if shp.has_table:
                shp._element.getparent().remove(shp._element)
            elif title and shp.has_text_frame and "요구사항" in shp.text_frame.text:
                _replace_text(shp, title)
        slides.append(extra)
    for sl, chunk in zip(slides, chunks):
        data = [rows[0]] + chunk
        row_h = min(0.78, 0.52 * 45 / len(data))
        tbl = _table(sl, 1.4, 1.3, 16.3, data, widths=widths, size=8 if row_h < 0.62 else 9,
                     row_h=row_h, merge_first=True).table
        if compliance is not None:
            # 명칭 칸은 왼쪽 정렬, 해당 페이지 칸은 가운데 정렬
            for i in range(1, len(data)):
                for j, al in ((2, "l"), (4, "ctr")):
                    for par in tbl.cell(i, j).text_frame.paragraphs:
                        par._p.get_or_add_pPr().set("algn", al)
    return slides


# ---------------------------------------------------------------- 본문 슬라이드

class Body:
    def __init__(self, prs, chapter, spec, boost=0, quiet=False):
        self.prs = prs
        self.s = _duplicate(prs, PROTO_BODY)
        self.y = Y_START
        self.spec = spec
        # boost: 여백이 남을 때 글자 크기·행 간격을 키우는 단계 (0 = 기본)
        self.boost = boost
        self.quiet = quiet
        self.overflow = False
        s = self.s
        _replace_text(_shape_by_name(s, "텍스트 개체 틀 9"), spec["header"])
        _replace_text(_shape_by_name(s, "TextBox 28"), f"{chapter['num']}. {chapter['title']}")
        _replace_text(_shape_by_name(s, "직사각형 2"), spec.get("h2", " "))
        h3 = _shape_by_name(s, "직사각형 1")
        _replace_text(h3, spec.get("h3", " "))
        summary = _shape_by_name(s, "사각형: 둥근 모서리 3")
        if spec.get("summary"):
            _replace_text(summary, spec["summary"])
            self._emphasize(summary, spec["summary"])
        else:
            summary._element.getparent().remove(summary._element)
            self.y = 3.9
        # 템플릿의 "제목 기입" 소제목은 블록으로 다시 그리므로 지운다
        for n in ("직선 연결선 4", "TextBox 5"):
            el = _shape_by_name(s, n)._element
            el.getparent().remove(el)
        h3.width = Cm(10.3) if spec.get("req") else Cm(16.5)
        # 소제목이 한 줄을 넘으면 글자를 줄여 위 제목과 겹치지 않게 한다
        h3_text = spec.get("h3", "")
        for size in (14, 12.5, 11):
            if _est_lines(h3_text, size, Emu(h3.width).cm) <= 1:
                break
        if size != 14:
            for r in h3._element.iter(qn("a:rPr")):
                r.set("sz", str(int(size * 100)))
        if spec.get("req"):
            box = _shape(s, "ROUNDED_RECTANGLE", 12.17, 2.35, 6.19, 0.9, fill="DCE9F7", adj=0.05)
            _fill_shape_text(box, spec["req"], 10.5 if len(spec["req"]) < 30 else 9)

    def _emphasize(self, shp, text):
        """요약 박스에서 '**강조**'를 파란색으로 바꾼다."""
        if "**" not in text:
            return
        par = shp.text_frame.paragraphs[0]
        rPr = copy.deepcopy(par.runs[0]._r.find(qn("a:rPr")))
        for r in list(par.runs):
            par._p.remove(r._r)
        for chunk, em in _runs(text):
            r = etree.SubElement(par._p, qn("a:r"))
            rp = copy.deepcopy(rPr)
            if em:
                for f in rp.findall(qn("a:solidFill")):
                    rp.remove(f)
                sf = etree.Element(qn("a:solidFill"))
                c = etree.SubElement(sf, qn("a:srgbClr"))
                c.set("val", BLUE)
                ln = rp.find(qn("a:ln"))
                rp.insert(list(rp).index(ln) + 1 if ln is not None else 0, sf)
            r.append(rp)
            t = etree.SubElement(r, qn("a:t"))
            t.text = chunk

    # 남은 높이
    def room(self):
        return Y_END - self.y

    def sz(self, b, default):
        return b.get("size", default) + 0.5 * self.boost

    def warn(self, need, what):
        if need > self.room() + 0.05:
            self.overflow = True
            if self.quiet:
                return
            print(f"  ! 넘침 {need - self.room():.1f}cm  [{self.spec.get('h3') or self.spec.get('h2')}] {what}",
                  file=sys.stderr)

    # -- 블록들 --------------------------------------------------------
    def subtitle(self, b):
        s = self.s
        _shape(s, "RECTANGLE", 0.84, self.y + 0.04, 0.13, 0.62, fill=BLUE)
        tb = _text_box(s, 1.09, self.y, 17.1, 0.7, anchor="m")
        par = tb.text_frame.paragraphs[0]
        _add_runs(par, b["text"], 12, FONT_B)
        self.y += 0.9

    def bullets(self, b):
        items = b["items"]
        size = self.sz(b, 12)
        sub_size = b.get("sub_size", size - 1)
        x = X0 + 0.15
        w = W - 0.3
        h = 0.25
        for it in items:
            it = it if isinstance(it, dict) else {"text": it}
            h += _est_lines(it["text"], size, w, 0.3) * _line_cm(size, 1.25) + 0.45
            for sub in it.get("sub", []):
                h += _est_lines(sub, sub_size, w, 0.7) * _line_cm(sub_size, 1.25) + 0.12
        h += 0.2 * len(items)  # 글꼴 차이에 대비한 여유
        frame = b.get("frame", True)
        self.warn(h, "bullets")
        if frame:
            _shape(self.s, "RECTANGLE", X0, self.y, W, h + 0.1, line=NAVY, line_w=1)
        tb = _text_box(self.s, x, self.y + 0.1, w, h, inset=0.1)
        tf = tb.text_frame
        first = True
        for it in items:
            it = it if isinstance(it, dict) else {"text": it}
            par = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            _para_fmt(par, line=1.25, before=6, after=6)
            _bullet(par, "•", 108000, 108000)
            _add_runs(par, it["text"], size)
            for sub in it.get("sub", []):
                sp = tf.add_paragraph()
                _para_fmt(sp, line=1.25, after=3)
                _bullet(sp, "–", 180975, 361950, FONT_M)
                _add_runs(sp, sub, sub_size)
        self.y += h + 0.35

    def text(self, b):
        size = self.sz(b, 11)
        h = _est_lines(b["text"], size, W) * _line_cm(size, 1.3) + 0.2
        tb = _text_box(self.s, X0, self.y, W, h, inset=0.05)
        par = tb.text_frame.paragraphs[0]
        _para_fmt(par, line=1.3, align=b.get("align"))
        _add_runs(par, b["text"], size, FONT_B if b.get("bold") else FONT_M,
                  b.get("color", "000000"))
        self.y += h + 0.15

    def caption(self, b):
        tb = _text_box(self.s, X0, self.y, W, 0.6, anchor="m")
        par = tb.text_frame.paragraphs[0]
        _para_fmt(par, align="c")
        _add_runs(par, f"< {b['text']} >", 11, FONT_B)
        self.y += 0.65

    def table(self, b):
        if b.get("caption"):
            self.caption({"text": b["caption"]})
        rows = [b["header"]] + b["rows"]
        size = self.sz(b, 10.5)
        widths = b.get("widths")
        total = W - 0.2
        if not widths:
            widths = [total / len(b["header"])] * len(b["header"])
        else:
            k = total / sum(widths)
            widths = [w_ * k for w_ in widths]
        heights = []
        for r_i, row in enumerate(rows):
            lines = max(_est_lines(c, size + (1 if r_i == 0 else 0), widths[j] - 0.2) for j, c in enumerate(row))
            pad = 0.4 + 0.12 * self.boost
            heights.append(max((0.95 if r_i else 0.85) + 0.2 * self.boost, lines * _line_cm(size, 1.3) + pad))
        need = sum(heights)
        self.warn(need, f"table {b.get('caption', '')}")
        _table(self.s, X0 + 0.1, self.y, total, rows, widths=widths, size=size, heights=heights,
               merge_first=b.get("merge_first", False), first_col_fill=b.get("first_col_fill", True),
               align=b.get("align"))
        self.y += need + 0.3
        if b.get("note"):
            self.note({"text": b["note"]})

    def flow(self, b):
        """단계 카드(가로): steps: [{title, desc|items}]  — 예시 슬라이드 120 하단 형식."""
        steps = b["steps"]
        n = len(steps)
        gap = 0.3
        w = (W - 0.2 - gap * (n - 1)) / n
        desc_size = self.sz(b, 10)
        body_h = 0.3
        for st in steps:
            items = st.get("items") or [st.get("desc", "")]
            hh = sum(_est_lines(t, desc_size, w - 0.4) * _line_cm(desc_size, 1.3) + 0.12 for t in items)
            body_h = max(body_h, hh + 0.25)
        body_h = max(body_h, 1.6 + 0.3 * self.boost)
        h = 0.75 + body_h
        self.warn(h, "flow")
        for i, st in enumerate(steps):
            x = X0 + 0.1 + i * (w + gap)
            _shape(self.s, "ROUNDED_RECTANGLE", x, self.y, w, h, fill="FFFFFF", line="0070C0", adj=0.04)
            head = _shape(self.s, "ROUND_2_SAME_RECTANGLE", x, self.y, w, 0.72, fill=BLUE, adj=0.12)
            _fill_shape_text(head, st["title"], 10, FONT_B, "FFFFFF")
            tb = _text_box(self.s, x + 0.1, self.y + 0.8, w - 0.2, body_h, inset=0.05)
            tf = tb.text_frame
            items = st.get("items") or ([st["desc"]] if st.get("desc") else [])
            for k, t in enumerate(items):
                par = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                _para_fmt(par, line=1.3, after=2)
                if len(items) > 1 or st.get("items"):
                    _bullet(par, "•", 88900, 88900)
                _add_runs(par, t, desc_size)
            if i < n - 1:
                tri = _shape(self.s, "ISOSCELES_TRIANGLE", x + w + 0.03, self.y + h / 2 - 0.12, 0.24, 0.22,
                             fill=BLUE)
                tri.rotation = 90
        self.y += h + 0.35

    def process(self, b):
        """행 단위 추진 프로세스 — 예시 슬라이드 13 형식.
        rows: [{label, items:[...], result}]"""
        rows = b["rows"]
        label_w, result_w = 2.4, 2.8
        mid_x = X0 + 0.2 + label_w + 0.3
        mid_w = W - 0.2 - label_w - 0.3 - result_w - 0.35
        box_h = 1.0 + 0.15 * self.boost
        cols = b.get("cols", 2)
        total = 0
        layout = []
        for r in rows:
            n = len(r["items"])
            lines = -(-n // cols)
            hh = max(1.5, lines * (box_h + 0.25) + 0.3)
            layout.append(hh)
            total += hh + 0.2
        self.warn(total, "process")
        y = self.y
        for r, hh in zip(rows, layout):
            _shape(self.s, "RECTANGLE", X0 + 0.1, y, W - result_w - 0.45, hh, line=LIGHT_LINE)
            lb = _text_box(self.s, X0 + 0.15, y, label_w, hh, anchor="m", inset=0.05)
            par = lb.text_frame.paragraphs[0]
            _para_fmt(par, align="c", line=1.15)
            _add_runs(par, r["label"], 10, FONT_B, DARK)
            _shape(self.s, "RECTANGLE", X0 + 0.2 + label_w + 0.05, y + 0.2, 0.02, hh - 0.4, fill=LIGHT_LINE)
            bw = (mid_w - 0.25 * (cols - 1)) / cols
            for k, it in enumerate(r["items"]):
                cx = mid_x + (k % cols) * (bw + 0.25)
                cy = y + 0.3 + (k // cols) * (box_h + 0.25)
                full = (k == len(r["items"]) - 1 and len(r["items"]) % cols == 1 and cols > 1)
                bx = _shape(self.s, "ROUNDED_RECTANGLE", cx, cy, mid_w if full else bw, box_h, fill=LIGHT,
                            line="215F9A", adj=0.03)
                _fill_shape_text(bx, it, 10, FONT_B, "404040")
            if r.get("result"):
                rb = _shape(self.s, "RECTANGLE", X0 + W - result_w - 0.1, y, result_w, hh, fill=LIGHT,
                            line=LIGHT_LINE)
                _fill_shape_text(rb, r["result"], 10.5, FONT_B, DARK)
            y += hh + 0.2
        self.y = y + 0.15

    def boxes(self, b):
        """그리드 상자: items: [{title, desc}] , cols"""
        items = b["items"]
        cols = b.get("cols", 3)
        gap = 0.25
        w = (W - 0.2 - gap * (cols - 1)) / cols
        size = self.sz(b, 10)
        rows_n = -(-len(items) // cols)
        body_h = 0.3
        for it in items:
            body = it.get("desc", "")
            lines = body if isinstance(body, list) else [body]
            hh = sum(_est_lines(t, size, w - 0.3) * _line_cm(size, 1.3) + 0.1 for t in lines)
            body_h = max(body_h, hh + 0.2)
        body_h = max(body_h, 1.8 + 0.3 * self.boost)
        h = 0.7 + body_h
        need = rows_n * (h + gap)
        self.warn(need, "boxes")
        for i, it in enumerate(items):
            x = X0 + 0.1 + (i % cols) * (w + gap)
            y = self.y + (i // cols) * (h + gap)
            _shape(self.s, "RECTANGLE", x, y, w, h, fill="F4F8FC", line=LIGHT_LINE)
            hd = _shape(self.s, "RECTANGLE", x, y, w, 0.65, fill=b.get("head_fill", "DCE9F7"))
            _fill_shape_text(hd, it["title"], 10, FONT_B, DARK)
            tb = _text_box(self.s, x + 0.1, y + 0.72, w - 0.2, body_h, inset=0.05)
            tf = tb.text_frame
            body = it.get("desc", "")
            lines = body if isinstance(body, list) else [body]
            for k, t in enumerate(lines):
                par = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                _para_fmt(par, line=1.3, after=2)
                if isinstance(body, list):
                    _bullet(par, "•", 88900, 88900)
                _add_runs(par, t, size)
        self.y += need + 0.15

    def layers(self, b):
        """계층 아키텍처: layers: [{name, items:[...], side}]"""
        layers = b["layers"]
        side_w = 3.6 if any(l.get("side") for l in layers) else 0
        main_w = W - 0.2 - (side_w + 0.25 if side_w else 0)
        lh = b.get("height", 1.55)
        need = len(layers) * (lh + 0.18)
        self.warn(need, "layers")
        palette = ["1F4E9A", "2E6FC4", "0E8FD6", "2A6EBB", "1C3F7A", "3A7BD5"]
        for i, l in enumerate(layers):
            y = self.y + i * (lh + 0.18)
            col = l.get("color", palette[i % len(palette)])
            _shape(self.s, "RECTANGLE", X0 + 0.1, y, main_w, lh, fill=col)
            nm = _text_box(self.s, X0 + 0.15, y, 2.9, lh, anchor="m", inset=0.05)
            par = nm.text_frame.paragraphs[0]
            _para_fmt(par, align="c", line=1.1)
            _add_runs(par, l["name"], 10, FONT_B, "FFFFFF", em_color="FFE699")
            items = l.get("items", [])
            if items:
                iw = (main_w - 3.2 - 0.15 * (len(items) - 1) - 0.15) / len(items)
                for k, it in enumerate(items):
                    bx = _shape(self.s, "ROUNDED_RECTANGLE", X0 + 3.2 + k * (iw + 0.15), y + 0.2, iw, lh - 0.4,
                                fill="FFFFFF", adj=0.06)
                    _fill_shape_text(bx, it, 9.5, FONT_B, "1F3864", inset=0.05)
            if l.get("side"):
                sb = _shape(self.s, "RECTANGLE", X0 + 0.1 + main_w + 0.25, y, side_w, lh, fill="F4F8FC",
                            line=LIGHT_LINE)
                _fill_shape_text(sb, l["side"], 9.5, FONT_M, "1F3864", inset=0.08)
        self.y += need + 0.2

    def note(self, b):
        tb = _text_box(self.s, X0, self.y - 0.1, W, 0.5)
        par = tb.text_frame.paragraphs[0]
        _add_runs(par, b["text"], 9.5, FONT_M, "666666")
        self.y += 0.5

    def callout(self, b):
        size = self.sz(b, 10.5)
        h = _est_lines(b["text"], size, W - 0.6) * _line_cm(size, 1.35) + 0.45
        self.warn(h, "callout")
        bx = _shape(self.s, "ROUNDED_RECTANGLE", X0, self.y, W, h, fill="FFF4E5", line="F4B183", adj=0.06)
        _fill_shape_text(bx, b["text"], size, FONT_B, "7F3F00", align="l", line=1.3, inset=0.25)
        self.y += h + 0.3

    def spacer(self, b):
        self.y += float(b.get("h", 0.3))

    def kpi(self, b):
        """핵심 수치 타일: items: [{value, unit, label}]"""
        items = b["items"]
        n = len(items)
        gap = 0.3
        w = (W - 0.2 - gap * (n - 1)) / n
        h = b.get("height", 2.4)
        self.warn(h, "kpi")
        for i, it in enumerate(items):
            x = X0 + 0.1 + i * (w + gap)
            _shape(self.s, "RECTANGLE", x, self.y, w, h, fill="EEF4FB", line=LIGHT_LINE)
            _shape(self.s, "RECTANGLE", x, self.y, 0.12, h, fill=BLUE)
            tb = _text_box(self.s, x + 0.3, self.y + 0.15, w - 0.45, 1.1, anchor="m")
            par = tb.text_frame.paragraphs[0]
            _add_runs(par, str(it["value"]), b.get("size", 24), FONT_B, NAVY)
            if it.get("unit"):
                _add_runs(par, " " + it["unit"], 11, FONT_B, NAVY)
            lb = _text_box(self.s, x + 0.3, self.y + 1.25, w - 0.45, h - 1.35, inset=0.02)
            par = lb.text_frame.paragraphs[0]
            _para_fmt(par, line=1.25)
            _add_runs(par, it["label"], 9.5, FONT_M, "404040")
        self.y += h + 0.35

    def themes(self, b):
        """번호 카드(차별화 요소 등): items: [{title, text, points:[...]}] — 번호는 순서대로 01, 02 …"""
        items = b["items"]
        size = self.sz(b, 10)
        num_w = 1.5
        tx = X0 + 0.1 + num_w + 0.3
        tw = W - 0.2 - num_w - 0.5
        total = 0
        layout = []
        for it in items:
            hh = 0.75 + _est_lines(it.get("text", ""), size, tw) * _line_cm(size, 1.3) + 0.15
            pts = it.get("points", [])
            ph = sum(_est_lines(p, size - 0.5, tw - 0.6) * _line_cm(size - 0.5, 1.25) + 0.08 for p in pts)
            hh += (ph + 0.35) if pts else 0
            hh = max(hh, num_w + 0.3)
            layout.append((hh, ph))
            total += hh + 0.25
        self.warn(total, "themes")
        y = self.y
        for i, (it, (hh, ph)) in enumerate(zip(items, layout), start=b.get("start", 1)):
            _shape(self.s, "RECTANGLE", X0 + 0.1, y, W - 0.2, hh, fill="FFFFFF", line=LIGHT_LINE)
            nb = _shape(self.s, "RECTANGLE", X0 + 0.1, y, num_w, num_w, fill=NAVY)
            _fill_shape_text(nb, f"{b.get('prefix', '')}{i:02d}", 18, FONT_B, "FFFFFF")
            tb = _text_box(self.s, tx, y + 0.12, tw, 0.65, anchor="m")
            _add_runs(tb.text_frame.paragraphs[0], it["title"], 12, FONT_B, "000000")
            yy = y + 0.8
            if it.get("text"):
                th = _est_lines(it["text"], size, tw) * _line_cm(size, 1.3) + 0.1
                t2 = _text_box(self.s, tx, yy, tw, th, inset=0.02)
                p = t2.text_frame.paragraphs[0]
                _para_fmt(p, line=1.3)
                _add_runs(p, it["text"], size, FONT_M, "333333")
                yy += th + 0.05
            pts = it.get("points", [])
            if pts:
                _shape(self.s, "RECTANGLE", tx, yy, tw, ph + 0.25, fill="EEF4FB")
                pb = _text_box(self.s, tx + 0.15, yy + 0.1, tw - 0.3, ph + 0.1, inset=0.02)
                for k, t in enumerate(pts):
                    p = pb.text_frame.paragraphs[0] if k == 0 else pb.text_frame.add_paragraph()
                    _para_fmt(p, line=1.25, after=1)
                    _bullet(p, "•", 108000, 108000)
                    _add_runs(p, t, size - 0.5, FONT_M, "1F3864")
            y += hh + 0.25
        self.y = y + 0.1

    def point(self, b):
        """제목이 있는 강조 상자 (설계 관점 · 제안 요지 등): {title, text, color: green|blue|navy}"""
        size = self.sz(b, 10)
        style = {"green": ("EAF6EF", "1E7B4F", "333333"), "blue": ("EEF4FB", BLUE, "333333"),
                 "navy": (NAVY, "FFE699", "FFFFFF")}[b.get("color", "green")]
        lines = _est_lines(b["text"], size, W - 0.8) * _line_cm(size, 1.35)
        h = 0.75 + lines + 0.3
        self.warn(h, "point")
        _shape(self.s, "RECTANGLE", X0, self.y, W, h, fill=style[0])
        tb = _text_box(self.s, X0 + 0.35, self.y + 0.15, W - 0.7, 0.55, anchor="m")
        _add_runs(tb.text_frame.paragraphs[0], b["title"], 10.5, FONT_B, style[1])
        t2 = _text_box(self.s, X0 + 0.35, self.y + 0.7, W - 0.7, lines + 0.2, inset=0.02)
        p = t2.text_frame.paragraphs[0]
        _para_fmt(p, line=1.35)
        _add_runs(p, b["text"], size, FONT_M, style[2])
        self.y += h + 0.3


def _table(slide, x, y, w, rows, widths, size=9, row_h=None, heights=None, merge_first=False,
           first_col_fill=True, align=None):
    n_r, n_c = len(rows), len(rows[0])
    gf = slide.shapes.add_table(n_r, n_c, Cm(x), Cm(y), Cm(w), Cm(row_h * n_r if row_h else sum(heights)))
    tbl = gf.table
    # 기본 표 스타일 해제
    tblPr = gf._element.graphic.graphicData.tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        tblPr.set(attr, "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        tblPr.remove(sid)
    for j, cw in enumerate(widths):
        tbl.columns[j].width = Cm(cw)
    for i in range(n_r):
        tbl.rows[i].height = Cm(row_h if row_h else heights[i])
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.margin_left = cell.margin_right = Cm(0.1)
            cell.margin_top = cell.margin_bottom = Cm(0.05)
            from pptx.enum.text import MSO_ANCHOR
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            lines = str(val).split("\n")
            for k, ln in enumerate(lines):
                par = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                left = (i > 0 and (align == "l" or (align is None and j == n_c - 1 and n_c > 1)))
                _para_fmt(par, line=1.15, align="l" if left else "c")
                if i == 0:
                    _add_runs(par, ln, size + 0.5, FONT_B, "FFFFFF", em_color="FFE699")
                else:
                    _add_runs(par, ln, size, FONT_B if (j == 0 and first_col_fill) else FONT_M, "222222")
            tcPr = cell._tc.get_or_add_tcPr()
            for side in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcPr, qn(side))
                ln.set("w", "10160")
                sf = etree.SubElement(ln, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr")).set("val", GRID)
            fill = BLUE if i == 0 else ("EEF3F9" if (j == 0 and first_col_fill) else "FFFFFF")
            sf = etree.SubElement(tcPr, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", fill)
    if merge_first:
        start = 1
        for i in range(2, n_r + 1):
            if i == n_r or rows[i][0] != rows[start][0]:
                if i - 1 > start:
                    for k in range(start + 1, i):
                        for r in tbl.cell(k, 0).text_frame.paragraphs[0].runs:
                            r.text = ""
                    tbl.cell(start, 0).merge(tbl.cell(i - 1, 0))
                start = i
    return gf


def _chapter_page_no(slide, label):
    """레이아웃의 연속 쪽번호를 가리고 하단 중앙에 장별 쪽번호(예: Ⅲ-5)를 넣는다."""
    _shape(slide, "RECTANGLE", 8.3, 26.75, 2.45, 0.75, fill="FFFFFF")
    tb = _text_box(slide, 7.5, 26.8, 4.05, 0.65, anchor="m")
    par = tb.text_frame.paragraphs[0]
    _para_fmt(par, align="c")
    _add_runs(par, label, 10, FONT_M, "404040")


# ---------------------------------------------------------------- 조립


def _draw(prs, ch, spec, boost, quiet):
    body = Body(prs, ch, spec, boost=boost, quiet=quiet)
    for blk in spec.get("blocks", []):
        getattr(body, blk.get("type", "bullets"))(blk)
    return body


def _render_body(prs, ch, spec, auto_fit=True, target=0.85, max_boost=4):
    """본문 슬라이드를 그린다. 여백이 많으면(채움 < target) 넘치지 않는 범위에서
    글자 크기·표 행 간격을 한 단계씩 키워 다시 그린다."""
    body = _draw(prs, ch, spec, 0, quiet=False)
    if not auto_fit or body.overflow:
        return body
    best = 0
    for boost in range(1, max_boost + 1):
        if (body.y - Y_START) / (Y_END - Y_START) >= target:
            break
        _delete_slide(prs, len(prs.slides) - 1)
        trial = _draw(prs, ch, spec, boost, quiet=True)
        if trial.overflow:
            _delete_slide(prs, len(prs.slides) - 1)
            return _draw(prs, ch, spec, best, quiet=True)
        body, best = trial, boost
    return body

def build(content_path: Path, out_path: Path, template: Path = BASE_TEMPLATE, report=False):
    data = load_content(content_path)
    prs = Presentation(str(template))
    n_proto = len(prs.slides)
    meta = data["meta"]
    chapters = data["chapters"]

    build_cover(prs, meta)
    req_slide = _duplicate(prs, PROTO_REQ) if data.get("req_table", True) else None
    build_toc(prs, chapters)

    n_ids = sum(len(c["ids"]) for c in data.get("req_categories", []))
    req_extra = max(0, -(-n_ids // 44) - 1) if req_slide is not None else 0
    req_pages: dict[str, list] = {}
    fill = []
    by_chapter = data.get("page_numbering") == "chapter"
    for ch in chapters:
        ch.setdefault("toc", [])
        build_divider(prs, ch)
        for k, spec in enumerate(ch.get("slides", []), start=1):
            body = _render_body(prs, ch, spec, auto_fit=data.get("auto_fit", True))
            page = len(prs.slides) - n_proto + req_extra
            if by_chapter:
                page = f"{ch['num']}-{k}"
                _chapter_page_no(body.s, page)
            fill.append((page, spec.get("h3") or spec.get("h2"), (body.y - Y_START) / (Y_END - Y_START)))
            for rid in re.split(r"[,\s]+", spec.get("req", "") or ""):
                if rid:
                    req_pages.setdefault(rid, []).append(page)

    if req_slide is not None:
        cats = data.get("req_categories", [])
        index = []
        for cat in cats:
            ids = cat["ids"]
            index.append((cat["name"], [(rid, sorted(set(req_pages.get(rid, [])))) for rid in ids]))
        missing = sorted(set(req_pages) - {r for c in cats for r in c["ids"]})
        if missing:
            index.append(("기타", [(rid, sorted(set(req_pages[rid]))) for rid in missing]))
        comp = None
        if data.get("req_compliance"):
            import yaml
            cpath = (content_path.parent / data["req_compliance"])
            comp = (yaml.safe_load(cpath.read_text(encoding="utf-8")) or {}).get("requirements") or {}
        build_req_table(prs, req_slide, index, title=data.get("req_table_title"), compliance=comp)
        unmapped = [rid for c in cats for rid in c["ids"] if rid not in req_pages]
        if unmapped:
            print("  ! 본문에 연결되지 않은 요구사항:", ", ".join(unmapped), file=sys.stderr)

    for _ in range(n_proto):
        _delete_slide(prs, 0)
    # 요구사항 분석표가 여러 장이면 표지 다음으로 옮긴다
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    extra_req = [i for i, s in enumerate(prs.slides) if s.slide_layout.name == "빈 화면"]
    for k, idx in enumerate(extra_req[1:], start=2):
        el = ids[idx]
        sldIdLst.remove(el)
        sldIdLst.insert(k, el)
    if report:
        for page, name, ratio in fill:
            flag = "  ← 여백 많음" if ratio < 0.8 else ""
            print(f"  p{page:>3} {ratio*100:5.0f}%  {name}{flag}")
    finalize(prs)
    set_pptx_author(prs)
    prs.save(str(out_path))
    print(f"저장: {out_path}  ({len(prs.slides)}장)")


def main(argv=None):
    ap = argparse.ArgumentParser(description="H9 정성제안서 PPT 생성")
    ap.add_argument("content", type=Path, help="내용 YAML 파일")
    ap.add_argument("-o", "--out", type=Path, help="출력 pptx 경로")
    ap.add_argument("--template", type=Path, default=BASE_TEMPLATE)
    ap.add_argument("--report", action="store_true", help="슬라이드별 본문 채움 비율 출력")
    args = ap.parse_args(argv)
    out = args.out or args.content.with_suffix(".pptx")
    build(args.content, out, args.template, args.report)


if __name__ == "__main__":
    main()
