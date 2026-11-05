"""정성제안서 앞장에 기술평가 항목 조견표를 만들어 content.yaml에 넣는다 (국립생물자원관 R26BK01808991).

조견표는 간지 없는 front 장으로 목차 바로 뒤에 놓인다. 조견표 장 수가 쪽수 계산에 들어가므로
한 번 더 실행해 쪽수를 맞춘다 (main()이 두 번 계산한다). 본문을 고친 뒤에는 다시 실행하고 빌드한다.
    python proposals/nibr-R26BK01808991/gen_eval_slides.py
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
    "수행실적": "실적 인정 범위 [회신대기]",
    "사업이해도": "수기 → 음성 · AI, 완전 구현 / 프로토타입 범위 구분",
    "추진전략": "4대 전략, 현장 위험 선시험, 기준선 비교",
    "적용기술": "4계층 구성, 내부 오픈소스 AI, 조사 유형 확장",
    "표준 프레임워크 적용": "전자정부 표준프레임워크 4.3 · 공통컴포넌트",
    "개발방법론": "SW 단계형 + AI 반복 실험형, 단계별 산출물",
    "기능요구사항": "고소음 STT · 발화 구조화 · 오프라인 · 격자 분석 · 보고서 · GIS",
    "보안요구사항": "외부 전송 차단, 암호화 · 감사 로그, 취약점 100% 조치",
    "데이터요구사항": "표준 우선, 전환 4단계 · 오류 유형별 처리 · 전담 조직",
    "시스템운영요구사항": "운영 제약 5종 대응, 운영 대시보드",
    "제약사항": "외부 전송 금지 · 자체 GPU · 반출 [회신대기]",
    "성능요구사항": "건당 3~5초, 검색 3초, 동시 100명 3초, LLM 2초",
    "품질요구사항": "가동률 99.5%, 단계별 통과 기준",
    "인터페이스요구사항": "철새정보시스템 연계 API, 현장 · 관리자 화면",
    "관리방법론": "정기 보고 · 위험 7대 · 변경 통제",
    "일정계획": "180일 · 3개 마일스톤, STT 기준 1개월 내",
    "개발장비": "자체 GPU · 독립망 개발실 · 환경 분리",
    "안전보건 확보방안": "안전보건관리 계획서 · 현장 동행 안전",
    "품질보증": "독립 품질보증 · 단계 점검",
    "시험운영": "수기 · 음성 병행 기록 효율 검증",
    "교육훈련": "대상별 실습 · 매뉴얼 · 동영상, 재학습 이전",
    "하자보수": "검수 후 1년 무상, 등급별 대응",
    "기밀보안": "보안서약 · CCTV · 지문 · 독립망 · VPN",
    "비상대책": "일 · 주 · 월 백업, 연 1회 원격지 소산",
}
SPLIT = [("정량평가", "전략 및 방법론"), ("기술 및 기능", "성능 및 품질"), ("프로젝트 관리", "프로젝트 지원")]
NOTES = [
    ("평가 비율", "기술평가 90% (정량 10 + 정성 80) : 가격평가 10%, 기술능력평가 배점 한도의 85% 이상인 자를 협상적격자로 선정함"),
    ("핵심 과업", "고소음 STT · 복합 발화 구조화(기능요구사항-009)는 현장 실증 수준으로 완전 구현하고, LLM · MLOps는 제안요청서 범위대로 프로토타입 · 설계 산출물로 수행함"),
    ("요구사항 수용", "제안요청서 요구사항 69개 전체를 수용하고, 요구사항별 해당 쪽은 요구사항 대비 수용여부 조견표에 표시함"),
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
