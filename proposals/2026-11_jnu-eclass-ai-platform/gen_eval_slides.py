"""정성제안서 끝(Ⅶ 붙임)에 기술성 평가항목 조견표 슬라이드를 만들어 content.yaml에 넣는다.

쪽수는 eval_spec.yaml과 build_lookup_table.build_rows로 계산한다. 조견표는 맨 뒤에 있으므로
앞쪽 쪽수는 바뀌지 않는다. 본문을 고친 뒤에는 이 스크립트를 다시 실행하고 빌드한다.
    python proposals/2026-11_jnu-eclass-ai-platform/gen_eval_slides.py
"""
import re
import sys
from pathlib import Path

import yaml

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parents[1]))
from proposal_kit.build_lookup_table import build_rows  # noqa: E402

NOTES = [
    ("평가 비율", "기술평가 90 (정성 85 · 정량 5) : 가격평가 10, 기술능력평가 점수가 배점 한도의 85% 이상인 자를 협상적격자로 선정함 (차등점수제 미적용). 평가위원 점수 중 최고 · 최저를 1개씩 제외한 평균으로 기술평가 점수를 산출함"),
    ("기술 및 기능 · 성능 및 품질 (40점)", "요구사항 131개 전체를 수용하고, 요구사항 대비 수용여부 조견표(본문 앞)에 요구사항별 해당 쪽을 표시함"),
    ("프로젝트 관리 · 지원 (25점)", "120일 일정의 마일스톤 · 품질보증 · 시험운영 · 교육 · 하자보수 12개월 · 기밀보안 · 비상대책을 쪽별로 제시함"),
]

KPIS = [
    [("5", "정량평가 (경영상태)"), ("20", "전략 및 방법론"), ("85", "정성평가 합계")],
    [("25", "기술 및 기능"), ("15", "성능 및 품질")],
    [("10", "프로젝트 관리"), ("15", "프로젝트 지원")],
]

BEGIN, END = "  # >>> 평가항목 조견표 (gen_eval_slides.py가 생성)", "  # <<< 평가항목 조견표"


def q(s):
    return '"' + str(s).replace('"', "'") + '"'


def main():
    path = HERE / "content.yaml"
    text = path.read_text(encoding="utf-8")
    text = re.sub(re.escape(BEGIN) + r".*?" + re.escape(END) + r"\n?", "", text, flags=re.S)
    content = yaml.safe_load(text)
    spec = yaml.safe_load((HERE / "eval_spec.yaml").read_text(encoding="utf-8"))
    groups = {e["item"]: e["group"] for e in spec["evaluation"]}
    merged = {}
    for r in build_rows(content, spec):
        item = r[0].split("\n")[0]
        m = merged.setdefault(item, {"score": r[0].split("\n")[1], "text": r[3], "pages": [], "toc": []})
        if r[4] != "-":
            m["pages"].append(r[4])
        m["toc"] += [t for t in r[5].split("\n") if t != "-"]
    rows = [[groups[i], f"{i} {m['score']}", m["text"], " / ".join(dict.fromkeys(m["toc"]))[:70], ", ".join(m["pages"]) or "-"]
            for i, m in merged.items()]
    # 쪽 넘침을 막으려고 평가 구분 단위로 3장에 나눈다
    split = [("정량평가", "전략 및 방법론"), ("기술 및 기능", "성능 및 품질"), ("프로젝트 관리", "프로젝트 지원")]
    parts = [[r for r in rows if r[0] in grp] for grp in split]
    out = [BEGIN]
    for n, part in enumerate(parts, start=1):
        out += [
            "  - header: 2. 기술성 평가항목 조견표",
            "    h2: 2.1 평가항목별 제안내용 및 해당 쪽",
            f"    h3: 2.1.{n} 기술성 평가항목 조견표 ({n}/{len(parts)})",
            "    summary: \"(평가항목 → 제안 쪽) RFP [붙임1] 「제안서 기술 평가 항목 및 배점표」의 평가항목마다 **제안내용과 해당 쪽**을 정리함 (쪽은 표지부터 센 전체 쪽수)\"",
            "    blocks:",
            "    - type: table",
            "      size: 8.5",
            "      header: [구분, 평가항목 (배점), 평가 기준 요약, 제안내용 (제안서 목차), 쪽]",
            "      widths: [2.2, 3, 4.6, 4.6, 1.6]",
            "      merge_first: true",
            "      rows:",
        ] + [f"      - [{', '.join(q(c) for c in r)}]" for r in part] + [
            "    - type: point",
            "      color: navy",
            f"      title: {NOTES[n - 1][0]}",
            f"      text: {q(NOTES[n - 1][1])}",
            "    - type: kpi",
            "      items:",
        ] + [f"      - {{value: {q(v)}, unit: 점, label: {q(l)}}}" for v, l in KPIS[n - 1]]
    out.append(END)
    # Ⅶ장 slides 끝(파일 끝)에 붙인다
    path.write_text(text.rstrip("\n") + "\n\n" + "\n".join(out) + "\n", encoding="utf-8")
    print("조견표", len(rows), "항목 →", path)


if __name__ == "__main__":
    main()
