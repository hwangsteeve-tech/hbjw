"""정성제안서 앞장(RFP Ⅴ "제안서 앞장에 평가항목에 대한 조견표를 명시")에 기술성 평가항목 조견표를 만들어 content.yaml에 넣는다.

조견표는 간지 없는 front 장으로 목차 바로 뒤에 놓인다. 조견표 장 수가 쪽수 계산에 들어가므로
한 번 더 실행해 쪽수를 맞춘다 (main()이 두 번 계산한다). 본문을 고친 뒤에는 다시 실행하고 빌드한다.
    python proposals/2026-12_nmk-museum-web-maint/gen_eval_slides.py
"""
import re
import sys
from pathlib import Path

import yaml

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parents[1]))
from proposal_kit.build_lookup_table import build_rows  # noqa: E402

BEGIN, END = "# >>> 평가항목 조견표 (gen_eval_slides.py가 생성)", "# <<< 평가항목 조견표"

# 평가항목별 제안 포인트 (RFP 양식: 평가항목 · 평가요소 · 제안서 관련목차 · 페이지 · 제안 포인트)
POINTS = {
    "경영상태": "기업신용평가등급 BBB0 (배점의 100%)",
    "사업이해도": "통합 최초 사업 이해, 범위 경계 명확화, 4계층 운영체계",
    "추진전략": "단일 창구 · 도구 · SLA · 공통 개선, 2개월 3차수 인수",
    "시스템 운영 요구사항": "SR 필수 기능 8종, 5분 모니터링, 2h 착수 · 4h 정상화",
    "유지관리 인력 요구사항": "PM SPOC, 누리집별 정/부 담당, 2주 병행 인수인계",
    "보안 요구사항": "누출금지 정보 대장, 승인 단말, 시큐어코딩 · 웹쉘 점검",
    "품질 요구사항": "테스트 환경 검증 배포, 형상관리, 접근성 · 표준 준수",
    "관리방법론": "SLA 7개 지표, 이슈 등급별 당일 보고, 전자적 진척관리",
    "일정계획": "12개월 일정 · 6개 마일스톤 · 시기별 자원 배분",
    "품질보증": "품질 담당 분리 점검, 계획 · 점검 · 시정 · 보고",
    "교육훈련": "13개 기관 SR · 콘텐츠 교육, 운영 노하우 기술이전",
    "유지관리": "범위 · 조직 · 절차 · 기간, 종료 1개월 전 인계",
    "비상대책": "응용 백업 · 롤백, 등급별 대응, 예비 인력 1주 내",
}
SPLIT = [("조직 및 인력 (정량)", "전략 및 방법론", "운영 및 품질"), ("프로젝트 관리", "프로젝트 지원")]
NOTES = [
    ("평가 비율", "기술평가 90 (정성 85 · 정량 5) : 가격평가 10, 기술능력평가 배점 한도의 85% 이상인 자를 협상적격자로 선정함"),
    ("요구사항 수용", "요구사항 32개 전체를 수용하고, 요구사항별 해당 쪽은 요구사항 대비 수용여부 조견표에 표시함"),
]


def q(s):
    return '"' + str(s).replace('"', "'") + '"'


def render(rows, groups):
    out = [BEGIN, "- num: \"\"", "  title: 평가항목 조견표", "  front: true", "  toc: []", "  slides:"]
    for n, grp in enumerate(SPLIT, start=1):
        part = [r for r in rows if r[0] in grp]
        out += [
            "  - header: 기술성 평가항목 조견표",
            "    h2: 평가항목별 제안내용 및 해당 쪽",
            f"    h3: 기술성 평가항목 조견표 ({n}/{len(SPLIT)})",
            "    summary: \"(평가항목 → 제안 쪽) RFP 「기술성 평가기준」의 평가항목마다 **평가요소 · 관련 목차 · 쪽 · 제안 포인트**를 정리함 (쪽은 표지부터 센 전체 쪽수)\"",
            "    blocks:",
            "    - type: table",
            "      size: 8",
            "      header: [평가부문, 평가항목 (배점), 평가요소, 제안서 관련 목차, 쪽, 제안 포인트]",
            "      widths: [2, 2.6, 3.4, 3.4, 1.4, 3.2]",
            "      merge_first: true",
            "      rows:",
        ] + [f"      - [{', '.join(q(c) for c in r)}]" for r in part] + [
            "    - type: point",
            "      color: navy",
            f"      title: {NOTES[n - 1][0]}",
            f"      text: {q(NOTES[n - 1][1])}",
        ]
    out.append(END)
    return out


def once(path, spec, groups):
    text = path.read_text(encoding="utf-8")
    content = yaml.safe_load(text)
    merged = {}
    for r in build_rows(content, spec):
        item = r[0].split("\n")[0]
        m = merged.setdefault(item, {"score": r[0].split("\n")[1], "text": r[3], "pages": [], "toc": []})
        if r[4] != "-":
            m["pages"].append(r[4])
        m["toc"] += [t for t in r[5].split("\n") if t != "-"]
    rows = [[groups[i], f"{i} {m['score']}", m["text"], " / ".join(dict.fromkeys(m["toc"]))[:60],
             ", ".join(m["pages"]) or "-", POINTS.get(i, "")] for i, m in merged.items()]
    block = "\n".join(render(rows, groups))
    text = re.sub(re.escape(BEGIN) + r".*?" + re.escape(END), lambda _: block, text, flags=re.S)
    path.write_text(text, encoding="utf-8")
    return len(rows)


def main():
    path = HERE / "content.yaml"
    spec = yaml.safe_load((HERE / "eval_spec.yaml").read_text(encoding="utf-8"))
    groups = {e["item"]: e["group"] for e in spec["evaluation"]}
    once(path, spec, groups)          # 1차: 조견표 장을 넣는다
    n = once(path, spec, groups)      # 2차: 조견표 장이 들어간 상태로 쪽수를 다시 계산한다
    print("조견표", n, "항목 →", path)


if __name__ == "__main__":
    main()
