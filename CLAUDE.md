# 저장소 안내

- `kjebi_recommender/`: 크제비 알림메일 추천 공고 리포트 (README.md 참고)
- `proposal_kit/`: H9 정성제안서 PPT 생성 키트 (proposal_kit/README.md 참고)

## 제안서 작성 규칙

- 정성제안서는 항상 `proposal_kit`의 H9 양식(세로 A4, templates/h9_base.pptx)으로 PPT 파일을 만든다.
  형식 기준은 `proposal_kit/reference/`의 실제 제출본 예시를 따른다.
- 사업별 내용은 `proposals/<연월>_<사업명>/content.yaml`에 작성하고
  `python -m proposal_kit.build_proposal <content.yaml> -o <출력.pptx> --report`로 빌드한다.
- 목차는 해당 RFP의 "제안서 작성 목차(권고)"와 기술평가 항목을 따르고, 본문 슬라이드마다 요구사항 ID(`req`)를 단다.
- 회사 고유 정보(실적·인력·신용등급·인증)는 확인되지 않으면 `【 】`로 남기고 지어내지 않는다.
- 빌드 후 LibreOffice로 PDF를 만들어 넘침·겹침을 눈으로 확인한다.
- 제안 발표자료는 `proposal_kit/build_presentation.py`(16:9)로 만들고, 형식은 `reference/h9_발표자료_예시_*.pdf`를 따른다.
  내용은 정성제안서를 압축하되 평가항목 조견표의 모든 항목에 해당 페이지가 생기도록 `eval`을 단다.
