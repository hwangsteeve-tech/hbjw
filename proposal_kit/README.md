# H9 정성제안서 PPT 키트

에이치나인 표준 정성제안서 양식으로 PPT를 만드는 도구입니다. 양식은 세로 A4, KoPubWorld돋움체, 파란 계열 디자인입니다.
내용을 YAML 파일로 적으면 표지, 요구사항 분석표, 목차, 간지, 본문 슬라이드가 자동으로 만들어집니다.

```
proposal_kit/
├─ build_proposal.py        생성기
├─ templates/h9_base.pptx   기본 양식 (표지·분석표·목차·간지·본문 원형 5장)
└─ reference/               실제 제출본 예시 (AI학습용데이터 통합 제공 시스템, 220장)
proposals/<연월>_<사업명>/content.yaml   사업별 내용 파일
```

## 실행

```bash
pip install python-pptx pyyaml lxml
python -m proposal_kit.build_proposal proposals/2026-10_pnu-ai-platform/content.yaml \
    -o proposals/2026-10_pnu-ai-platform/pnu_proposal.pptx --report
```

`--report`를 붙이면 슬라이드마다 본문이 얼마나 찼는지 비율을 보여줍니다. 본문이 넘치면 `! 넘침` 경고가 나옵니다.
높이는 글자 수로 추정한 값이라 실제 PowerPoint에서 한 번 더 확인해야 합니다.

## 내용 파일 구조

```yaml
meta:
  doc_type: 정성제안서
  title_lines: [AI 기반 지역혁신, 데이터 플랫폼 구축]   # 표지 제목 (줄 단위)
  pm_line: 사업PM 홍길동 파트너
  org_lines: ["주관기관: 에이치나인", "참여기관: ○○"]

req_categories:                 # 요구사항 분석표 (본문 req 필드에서 쪽수를 자동 계산)
  - {name: 기능, ids: [FUR-001, FUR-002]}

chapters:
  - num: Ⅰ                      # 'no'는 YAML에서 false로 읽히므로 num을 씁니다
    title: 일반현황
    toc: [제안사 일반현황, 제안사의 조직 및 인원]   # 간지에 나오는 절 목록
    slides:
      - header: 1. 제안사 일반현황      # 좌상단 절 제목
        h2: 1.1 일반현황
        h3: 1.1.1 회사 개요
        req: FUR-001, FUR-002           # 우상단 요구사항 ID 상자 (생략 가능)
        summary: "(핵심) 헤드 메시지 **강조**"   # 연파랑 요약 상자
        blocks: [ ... ]                 # 아래 블록을 위에서부터 쌓음
```

본문 글에서 `**강조**`로 감싼 부분은 파란 굵은 글씨로 나옵니다.

## 본문 블록

| type | 용도 | 주요 키 |
|---|---|---|
| `subtitle` | ▎ 소제목 | `text` |
| `bullets` | 테두리 상자 안 글머리표 (• / –) | `items: [{text, sub: [...]}]`, `frame`, `size` |
| `table` | 파란 머리행 표 | `header`, `rows`, `widths`, `caption`, `note`, `size`, `merge_first` |
| `flow` | 가로 단계 카드 (①→②→③) | `steps: [{title, desc 또는 items}]` |
| `process` | 행 단위 추진 프로세스 (예시 13쪽 형식) | `rows: [{label, items, result}]`, `cols` |
| `boxes` | 격자 상자 | `items: [{title, desc}]`, `cols` |
| `layers` | 계층 구조도 (아키텍처·조직도) | `layers: [{name, items, side, color}]`, `height` |
| `callout` | 주황 강조 상자 | `text` |
| `text` / `caption` / `note` / `spacer` | 일반 문단 / `< 표 제목 >` / 회색 각주 / 여백 | |

## 작성 원칙 (예시 제출본 기준)

- 요약 상자는 `(키워드) 결론 문장` 형식으로 쓰고, 문장 끝은 "~함"으로 맺습니다.
- RFP에서 "할 수 있다·고려한다" 같은 모호한 표현은 불가능으로 간주되므로 쓰지 않습니다. 셀 수 있는 것은 숫자로 씁니다.
- 각 본문 슬라이드에 해당 요구사항 ID를 `req`로 달아야 분석표의 쪽수가 채워집니다. 빌드할 때 본문에 연결되지 않은 요구사항 ID를 알려 줍니다.
- 회사 고유 정보(실적, 인력, 신용등급 등)는 `【 】`로 표시하고 확인한 뒤 채웁니다.
- 복잡한 그림(화면 시안, 상세 구성도)은 생성된 PPT에서 직접 추가하고, 여기서는 뼈대와 문안을 만듭니다.
