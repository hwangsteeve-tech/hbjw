# 국립생물자원관 R26BK01808991 「AI기반 조사 데이터 구조화 및 분석 기술 개발」 제안 작업

- 사업금액 199,000,000원 (VAT 포함), 계약 후 180일, 인천지방조달청 대행
- 제안서 · 가격 전자입찰서 접수 마감 2026/11/20 10:00 / 발표는 사업관리자(PM)
- 공고 원본: `inbox/R26BK01808991/` (공고서 · 제안요청서 · 과업지시서 · 사업기간 산정서 · 영향평가 결과서)

## 산출물 (`output/`)

| 파일 | 내용 |
|---|---|
| `nibr_정성제안서.pptx` | 정성제안서 77쪽 (세로 A4 H9 양식, 목차 Ⅰ~Ⅵ, 평가항목 조견표 3쪽 포함) |
| `nibr_발표자료.pptx` | 발표자료(제안요약서) 27쪽 (16:9, 평가항목 25개 모두 해당 쪽 표시) |
| `제출서식.docx` | 점검표 · 별지 제1~4호 · [붙임 8] 안전보건관리 계획서 · 별지 제9호 안전보건 서약서 · 별지 제15호 수행실적 총괄표 |
| `요구사항_수용여부_조견표.xlsx` | 요구사항 69개 → 목차 · 쪽 · 수용 여부 |
| `정성제안서_평가항목_조견표.xlsx` | 기술평가 항목 25개 → 제안서 목차 · 쪽 |

## 빌드

```bash
python proposals/nibr-R26BK01808991/gen_eval_slides.py          # 앞쪽 평가항목 조견표 갱신 (쪽 바뀌면 다시)
python -m proposal_kit.build_proposal proposals/nibr-R26BK01808991/content.yaml -o proposals/nibr-R26BK01808991/output/nibr_정성제안서.pptx --report
python -m proposal_kit.build_presentation proposals/nibr-R26BK01808991/presentation.yaml -o proposals/nibr-R26BK01808991/output/nibr_발표자료.pptx --report
python proposals/nibr-R26BK01808991/build_forms.py
python -m proposal_kit.build_lookup_table proposals/nibr-R26BK01808991/content.yaml proposals/nibr-R26BK01808991/lookup_spec.yaml -o proposals/nibr-R26BK01808991/output/요구사항_수용여부_조견표.xlsx
python -m proposal_kit.build_lookup_table proposals/nibr-R26BK01808991/content.yaml proposals/nibr-R26BK01808991/eval_spec.yaml -o proposals/nibr-R26BK01808991/output/정성제안서_평가항목_조견표.xlsx
```

## [회신대기] — 발주기관 질의 후 확정

1. **수행실적 인정 범위**: 민간 발주 구축 실적, 공공 유지보수 실적의 '정보시스템 구축' 인정 여부, 별지 제3호 "생물자원 등 환경 정보화 분야만 기재" 적용 범위
   → 정성 Ⅱ-4.1.1 실적 표, 별지 제3호, 별지 제15호 총괄표(후보 목록에 '인정 여부' 칸)
2. **개발 데이터 원격지 반출 가능 여부 · 기존 음성 데이터 보유 규모**
   → Ⅰ-2.2.1 전제조건, Ⅲ-2.1.1 STT, Ⅳ-2.3.1 위험(대안: 발주기관 내 GPU 반입 학습), Ⅳ-3.1.1 개발장비
3. **지식재산권**: 제안요청서 '발주기관 · 수행사 공동 소유' vs [붙임 9] 저작재산권 양도 협약서
   → Ⅵ-2.1.1, 수용여부 조견표 프로젝트지원요구사항-004 비고

## 【 】 — 작성자가 채울 값

- 사업PM · 투입인력 성명 · 등급 · 투입률, 부서별 인원 (Ⅱ-3, 별지 제4호, 표지 PM)
- STT 목표 수치 (CER · 종명 인식률 — 착수 시 「시험 및 검수계획서」로 확정)
- 개발 GPU 사양, SP 인증 등 품질 인증, 기술 세미나 횟수
- 신용등급 · 기업 규모 (Ⅱ-1 일반현황 추가 행)
- 별지 제2호: 유동부채 · 고정부채 · 유동자산 · 매출원가 (재무제표 천원 값), 총자산 · 자기자본 · 당기순이익은 백만원 참고값 → 천원 값으로 교체
- [붙임 8]: 안전보건관리비, 비상연락망 연락처, 안전 담당자, 작업 참여자 수, 관할 보건소 · 병원
- 별지 제9호: 계약기간 날짜, 서약일
