"""정성제안서 앞장(RFP Ⅴ "제안서 앞장에 평가항목에 대한 조견표를 명시")에 기술성 평가항목 조견표를 만들어 content.yaml에 넣는다.

조견표는 간지 없는 front 장으로 목차 바로 뒤에 놓인다. 조견표 장 수가 쪽수 계산에 들어가므로
한 번 더 실행해 쪽수를 맞춘다 (main()이 두 번 계산한다). 본문을 고친 뒤에는 다시 실행하고 빌드한다.
    python proposals/2026-11_kgt-ai-pipeline-safety/gen_eval_slides.py
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
    "경영상태": "기업신용평가등급 BBB0 (평점 5점)",
    "사업 이해도": "PoC = 숫자로 검증, 6대 사업 특성, 45개 요구사항 추적",
    "추진전략": "4대 전략, 겨울 준비 · 우천기 실증, 1 · 2차 실증",
    "적용기술": "2단 모델(단말 · 서버), Apache-2.0, PPS 동기 위치 추정, 이중 마스킹",
    "사업추진 방법론": "5단계 PoC (확보 → 학습 → 실증 → 검증 → 개선)",
    "기능 요구사항": "단말 7세트, 9개 기능 · 관제 · GIS · 연계 · API",
    "보안 요구사항": "전송 전 마스킹, 단말 암호화 · 상호인증, 취약점 0건",
    "제약사항": "8개 현장 제약요소 대응, 조건별 평가데이터",
    "성능 요구사항": "대상별 재현율 80~85%, 알림 2초, 화면 2/4초",
    "품질관리": "지표 산식 · 실증 원장 · 재현 가능한 평가",
    "일정계획": "착수 ~ 8월 9개월, 7개 마일스톤",
    "위험 및 이슈관리": "PoC 7대 위험, 예방 · 대안, 상 등급 당일 보고",
    "교육훈련": "4개 역할 그룹, 집합 · 개별, 모델 재학습 이전",
    "기밀보안": "누출금지 정보 관리, 매월 보안교육, 원격지 보안",
    "차별화 전략": "4대 차별화 + [선택] LLM 조치지원 · EOCS 연계",
}
SPLIT = [("경영상태 (계량)", "전략 및 방법론", "기술 및 기능 요구사항"), ("성능 및 품질 요구사항", "프로젝트 관리", "프로젝트 지원", "추가제안")]
NOTES = [
    ("평가 비율", "기술능력평가 90% (계량 5 + 비계량 95, 100점을 90점으로 환산) : 입찰가격평가 10%, 기술평가 배점 한도의 85% 이상인 자를 협상적격자로 선정함"),
    ("요구사항 수용", "과업지시서 요구사항 45개(선택 2개 포함) 전체를 수용하고, 요구사항별 해당 쪽은 요구사항 대비 수용여부 조견표에 표시함"),
]


def q(s):
    return '"' + str(s).replace('"', "'") + '"'


def render(rows, groups):
    out = [BEGIN, "- num: \"\"", "  title: 평가항목 조견표", "  front: true", "  toc: []", "  slides:"]
    for n, grp in enumerate(SPLIT, start=1):
        part = [r for r in rows if r[0] in grp]
        out += [
            "  - header: 기술평가 항목 조견표",
            "    h2: 평가항목별 제안내용 및 해당 쪽",
            f"    h3: 기술평가 항목 조견표 ({n}/{len(SPLIT)})",
            "    summary: \"(평가항목 → 제안 쪽) 제안요청서 「기술평가 항목 및 배점」의 평가항목마다 **평가요소 · 관련 목차 · 쪽 · 제안 포인트**를 정리함 (쪽은 표지부터 센 전체 쪽수)\"",
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
