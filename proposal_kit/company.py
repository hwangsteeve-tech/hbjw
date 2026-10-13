"""회사 고유 정보(company/h9_company.yaml)로 일반현황 · 연혁 · 재무 슬라이드를 만든다.

content.yaml 의 slides 목록에 아래처럼 쓰면 빌드할 때 슬라이드로 펼쳐진다.

    - include: company.general     # 일반현황 (회사명·대표자·사업분야·주소·연락처·설립·사업기간)
      header: 1. 일반현황
      h2: 1.1 일반현황 및 연혁
      h3: 1.1.1 회사 개요
      extra_rows: [[입찰자격, "소프트웨어사업자(1468) ..."]]   # 공고별로 덧붙일 행 (선택)
    - include: company.history     # 주요 연혁
    - include: company.finance     # 자본금 및 매출액 (최근 3년) + 재무 지표

header · h2 · h3 · summary · req 는 기본값이 있고 지정하면 덮어쓴다. extra_blocks 로 블록을 뒤에 더할 수 있다.
"""
from __future__ import annotations

import copy
from pathlib import Path

import yaml

COMPANY_FILE = Path(__file__).parent / "company" / "h9_company.yaml"


def load_company(path: Path = COMPANY_FILE) -> dict:
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def document_author() -> str:
    return load_company().get("document_author", "")


def set_pptx_author(prs):
    """PPT 파일 정보의 만든 이 · 마지막으로 수정한 사람을 회사 정보의 작성자로 바꾼다."""
    cp = prs.core_properties
    cp.author = cp.last_modified_by = document_author()


def set_xlsx_author(wb):
    wb.properties.creator = wb.properties.lastModifiedBy = document_author()


def _won(v):
    return "-" if not v else f"{v:,}"


def _num(v):
    return "-" if v is None else f"{v:,}" if isinstance(v, int) else str(v)


def _general(co, spec):
    g = co["general"]
    rows = [
        ["회사명", f"{g['name']} ({g['name_en']})"],
        ["대표자", g["ceo"]],
        ["사업분야", g["business"]],
        ["주소", g["address"]],
        ["연락처", f"TEL {g['tel']} / FAX {g['fax']}"],
        ["회사설립년도", g["founded"]],
        ["해당부문 사업기간", g["biz_period"]],
    ] + spec.get("extra_rows", [])
    return {
        "header": "1. 일반현황", "h2": "1.1 일반현황 및 연혁", "h3": "1.1.1 회사 개요",
        "summary": "(사용자 중심 디지털 서비스 전문기업) 2010년 설립 이후 **UX 기획·디자인부터 플랫폼 개발·AI 구축·운영까지** 수행해 온 S/W개발 및 디자인 전문기업임",
        "blocks": [
            {"type": "callout", "text": "\"우리는 기술을 통해 창의적인 경험을 만들어갑니다. 사려 깊은 디자인과 사용자 중심의 접근 방식을 바탕으로, 사람과 제품, 디지털 서비스 간의 의미 있는 연결을 만드는 것이 우리의 목표입니다.\""},
            {"type": "subtitle", "text": "일반현황"},
            {"type": "table", "header": ["구분", "내용"], "widths": [3.4, 12.6], "rows": rows},
        ],
    }


def _history(co, spec):
    rows = [[str(y), "\n".join(f"ㆍ{t}" for t in items)] for y, items in co["history"]]
    return {
        "header": "1. 일반현황", "h2": "1.1 일반현황 및 연혁", "h3": "1.1.2 주요 연혁 및 수상",
        "summary": "(검증된 서비스 품질) iF · Good Design · ICT AWARD KOREA 등 국내외 수상과 **한국디자인진흥원 우수디자인전문기업** 선정으로 서비스 품질을 입증함",
        "blocks": [{"type": "table", "size": 9.5, "header": ["연도", "주요 연혁"], "widths": [2, 14], "rows": rows}],
    }


def _finance(co, spec):
    f = co["finance"]
    labels = [f"{y}년" for y in f["years"]]
    money = [
        ["자본금"] + [_won(v) for v in f["capital"]],
        ["매출액 (임대료수입)"] + [_won(v) for v in f["revenue_rent"]],
        ["매출액 (용역매출)"] + [_won(v) for v in f["revenue_service"]],
        ["매출액 합계"] + [_won(v) for v in f["revenue_total"]],
    ]
    keys = [("매출액", "sales"), ("매출원가", "cost_of_sales"), ("매출총이익", "gross_profit"),
            ("경상이익", "ordinary_income"), ("당기순이익", "net_income"), ("자산총계", "total_assets"),
            ("부채총계", "total_liabilities"), ("자본총계", "total_equity"), ("부채비율", "debt_ratio"),
            ("유동비율", "current_ratio"), ("매출액순이익률", "net_margin")]
    ratios = [[name] + [_num(v) for v in f[k]] for name, k in keys]
    return {
        "header": "2. 재무상황", "h2": "2.1 재무상황", "h3": "2.1.1 자본금 및 매출액 (최근 3년)",
        "summary": f"(안정적 재무 기반) 최근 3년 매출 **{min(f['sales']) / 100:.0f}억 ~ {max(f['sales']) / 100:.0f}억 원**을 유지하고, "
                   f"부채비율을 {f['debt_ratio'][0]}에서 **{f['debt_ratio'][-1]}**로 낮춰 재무 건전성을 높임",
        "blocks": [
            {"type": "subtitle", "text": "자본금 및 매출액 (단위: 원)"},
            {"type": "table", "size": 9.5, "header": ["구분"] + labels, "widths": [4.6, 3.8, 3.8, 3.8], "align": "c", "rows": money},
            {"type": "subtitle", "text": "주요 재무 지표 (단위: 백만원)"},
            {"type": "table", "size": 9, "header": ["구분"] + labels, "widths": [4.6, 3.8, 3.8, 3.8], "align": "c", "rows": ratios},
            {"type": "note", "text": "※ 공인회계사 확인 재무제표 별도 첨부 (부채·유동비율, 매출액순이익률은 표준 산식 적용)"},
        ],
    }


BUILDERS = {"company.general": _general, "company.history": _history, "company.finance": _finance}


def expand_slide(spec: dict, co: dict | None = None) -> dict:
    """include 슬라이드를 일반 슬라이드 spec 으로 펼친다. include 가 없으면 그대로 돌려준다."""
    key = spec.get("include")
    if not key:
        return spec
    if key not in BUILDERS:
        raise ValueError(f"알 수 없는 include: {key} (사용 가능: {', '.join(BUILDERS)})")
    out = BUILDERS[key](co or load_company(), spec)
    for k in ("header", "h2", "h3", "summary", "req"):
        if k in spec:
            out[k] = spec[k]
    out["blocks"] = out["blocks"] + copy.deepcopy(spec.get("extra_blocks", []))
    return out


def load_content(path: Path) -> dict:
    """content.yaml 을 읽고 include 슬라이드를 펼친다."""
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    co = None
    for ch in data.get("chapters", []):
        slides = ch.get("slides", [])
        if any("include" in s for s in slides):
            co = co or load_company()
            ch["slides"] = [expand_slide(s, co) for s in slides]
    return data
