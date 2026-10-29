# 한국가스기술공사 「AI 기반 실시간 가스배관망 안전관리시스템 시범구축 용역」

- 공고 제2026-47호 (R26BK01796859-000), 227,040,000원 (VAT 포함 · 공급가액 206,400,000원), 계약체결일 ~ 2027.8.31
- 제안서 · 입찰서 제출 2026.11.05(목) 10:00 ~ 11.09(월) 10:00 (나라장터, PDF 총 200MB 이내), 참가자격 등록 마감 11.06(금) 18:00
- 제안 발표 2026.11.10(화) 대전 본사 — PM 직접 발표 30분(시연 포함) + 질의 10분
- 평가: 기술 90 (계량 5 경영상태 + 비계량 95, 90점 환산) : 가격 10, 협상적격 85% 이상 / 공동수급 · 하도급 불허

## 제출 파일

| 제출 구분 | 파일 | 만드는 방법 |
|---|---|---|
| 정성제안서 (업체명 표기) | `kgt_정성제안서.pptx` (81쪽) | `content.yaml` → build_proposal |
| 정성제안서 (업체명 미표기) | `블라인드_정성제안서.pptx` | `python -m proposal_kit.make_blind kgt_정성제안서.pptx -o 블라인드_정성제안서.pptx` |
| 정량제안서 | `kgt_정량제안서.pptx` (9쪽, [서식3] 포함) | `quant.yaml` → build_proposal |
| 제안요약서 · 발표자료 (표기) | `kgt_발표자료.pptx` (29쪽, 16:9) | `presentation.yaml` → build_presentation |
| 제안요약서 · 발표자료 (미표기) | `블라인드_발표자료.pptx` | make_blind |
| 기타 제출서류 | `기타제출서류_서식1-4_위임장.docx` | `build_forms.py` ([서식1~4] · 발표 대리인 위임장 · 제출서류 점검표) |
| 서약서 (계약특수조건) | `계약특수조건_각종서약서_에이치나인.xlsx` | 공고 첨부 xlsx 「기본정보」 계약상대자 칸에 회사 정보 입력 |
| 조견표 | `요구사항_수용여부_조견표.xlsx` · `정성제안서_평가항목_조견표.xlsx` | build_lookup_table (정성제안서 앞장에도 수록) |

## 다시 빌드

```bash
D=proposals/2026-11_kgt-ai-pipeline-safety
python $D/gen_eval_slides.py                       # 앞장 평가항목 조견표 · 쪽수 갱신
python -m proposal_kit.build_proposal $D/content.yaml -o $D/kgt_정성제안서.pptx --report
python -m proposal_kit.build_proposal $D/quant.yaml -o $D/kgt_정량제안서.pptx --report
python -m proposal_kit.build_presentation $D/presentation.yaml -o $D/kgt_발표자료.pptx --report
python -m proposal_kit.make_blind $D/kgt_정성제안서.pptx -o $D/블라인드_정성제안서.pptx
python -m proposal_kit.make_blind $D/kgt_발표자료.pptx -o $D/블라인드_발표자료.pptx
python -m proposal_kit.build_lookup_table $D/content.yaml $D/lookup_spec.yaml -o $D/요구사항_수용여부_조견표.xlsx
python -m proposal_kit.build_lookup_table $D/content.yaml $D/eval_spec.yaml -o $D/정성제안서_평가항목_조견표.xlsx
python $D/build_forms.py
```

## 제출 전 채울 것 (붉은 【 】)

- 사업PM · 투입인력 성명 · 등급 · 인원, [서식4] PM 수행경험, 발표 대리인 위임장
- 신용평가등급 확인서 (기관 · 평가일 · 유효기간) 사본 — 정량제안서 첨부
- 단말 · GPU 워크스테이션 제조사 · 모델 · 단가, 통신 요금제 · 단말 전력 소비량
- 제안 목표값(라벨 오류율, 동기 오차, 평가데이터 수량, 백업 보관기간, 동시사용자 등) 확인
