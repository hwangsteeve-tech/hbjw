# 저장소 안내

- `kjebi_recommender/`: 크제비 알림메일 추천 공고 리포트 (README.md 참고)
- `proposal_kit/`: H9 정성제안서 PPT 생성 키트 (proposal_kit/README.md 참고)

## 제안서 작성 규칙

- 표지 문서명은 기술제안서가 아니라 항상 `정성제안서`로 표시한다 (build_proposal.py에서 고정).
- 쪽번호는 항상 전체 쪽수(표지부터 센 문서 쪽수)를 본문 하단에만 넣는다. 장별(Ⅲ-5) 쪽번호는 쓰지 않으며 표지·목차·장 표지에는 표시하지 않는다 (build_proposal.py에서 고정).
- 정성제안서는 항상 `proposal_kit`의 H9 양식(세로 A4, templates/h9_base.pptx)으로 PPT 파일을 만든다.
  형식 기준은 `proposal_kit/reference/`의 실제 제출본 예시를 따른다.
- 사업별 내용은 `proposals/<연월>_<사업명>/content.yaml`에 작성하고
  `python -m proposal_kit.build_proposal <content.yaml> -o <출력.pptx> --report`로 빌드한다.
- 목차는 해당 RFP의 "제안서 작성 목차(권고)"와 기술평가 항목을 따르고, 본문 슬라이드마다 요구사항 ID(`req`)를 단다.
- 회사 고유 정보(실적·인력·신용등급·인증)는 확인되지 않으면 `【 】`로 남기고 지어내지 않는다.
- 수행실적·유사사업 실적 표(정성·정량제안서, 제안요약서, 제출양식)는 항상 최근 계약일(사업 시작일)이 맨 위에 오도록 내림차순으로 쓴다. 번호를 매기면 최신이 1번이다.
- 사업수행실적 대장은 `proposal_kit/company/h9_company.yaml`의 `records`가 기준이다. 사용자가 새 실적을 주면 바로 추가해 최근 계약일 내림차순으로 저장하고, 공고별 인정 실적은 이 목록에서 골라 쓴다.
- 일반현황·연혁·자본금 및 매출액(최근 3년)·보유 자격(조달청 등록업종·직접생산증명)은 `proposal_kit/company/h9_company.yaml`이 기준이다.
  사용자가 새 등록 정보를 주면 이 파일을 바로 고쳐 저장하고, 제안서의 보유 자격 칸은 이 목록과 맞춘다.
  제안서에는 직접 쓰지 말고 `include: company.general` / `company.history` / `company.finance` 슬라이드로 불러온다.
- 글꼴은 어느 PC에서나 보이는 `맑은 고딕`으로 저장한다 (proposal_kit/style.py가 저장 직전에 일괄 적용).
- 작성자가 채울 빈칸은 `【 】`로 쓰며, 빌드 시 붉은색 굵은 글씨로 표시된다. 빈 표 칸도 `""` 대신 `【성명】`처럼 쓴다.
- 생성되는 PPT·엑셀의 파일 정보(만든 이·마지막 수정한 사람)는 항상 `황보정운`이다 (h9_company.yaml `document_author`).
- 빌드 후 LibreOffice로 PDF를 만들어 넘침·겹침을 눈으로 확인한다.
- 본문 슬라이드는 빈 공간이 없게 채운다: `--report`에서 `← 여백 많음`(채움 80% 미만)이 뜨는 장은 표·흐름도·`kpi`·`point` 등 읽기 쉬운 블록을 더해 80% 이상으로 맞춘다.
- 제안 발표자료는 `proposal_kit/build_presentation.py`(16:9)로 만들고, 형식은 `reference/h9_발표자료_예시_*.pdf`를 따른다.
  내용은 정성제안서를 압축하되 평가항목 조견표의 모든 항목에 해당 페이지가 생기도록 `eval`을 단다.
  발표자료도 빈 공간이 없게 채운다: `--report`의 `← 여백 많음`을 없애고, 렌더링해서 카드·패널 안이 비어 보이면 항목을 보강한다.
