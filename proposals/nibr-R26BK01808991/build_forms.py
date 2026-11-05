"""국립생물자원관 「AI기반 조사 데이터 구조화 및 분석 기술 개발」 제출서식(DOCX).

제안요청서 [붙임 2] 별지 제1~4호 · 제9호, [붙임 8] 2026년도 안전보건관리 계획서,
수행실적 총괄표(별지 제15호, 정량평가 수행실적용)를 한 파일로 만든다.
회사 정보는 proposal_kit/company/h9_company.yaml 기준, 확인되지 않은 값은 【 】로 남긴다.

    python proposals/nibr-R26BK01808991/build_forms.py
"""
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parents[1]))
from proposal_kit.company import load_company, document_author  # noqa: E402

FONT = "맑은 고딕"
PROJECT = "AI기반 조사 데이터 구조화 및 분석 기술 개발"
co = load_company()
g = co["general"]


def font(run, size=10, bold=False):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.bold = bold
    if "【" in run.text:
        run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        run.bold = True


def para(doc, text, size=10, bold=False, align=None, after=4):
    p = doc.add_paragraph()
    r = p.add_run(text)
    font(r, size, bold)
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    return p


def table(doc, rows, widths=None, head=True, size=9.5):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    widths = widths or [17 / len(rows[0])] * len(rows[0])
    tblPr = t._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(sum(widths) / 2.54 * 1440)))
    grid = t._tbl.tblGrid
    for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[j] / 2.54 * 1440)))
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = t.cell(i, j)
            cell.text = ""
            r = cell.paragraphs[0].add_run(str(val))
            font(r, size, bold=(head and i == 0))
            if widths:
                cell.width = Cm(widths[j])
    doc.add_paragraph()
    return t


def title(doc, form, name):
    para(doc, form, 10, True)
    para(doc, name, 15, True, WD_ALIGN_PARAGRAPH.CENTER, 10)


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def won_k(v):
    return f"{round(v / 1000):,}"


def main():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for m in ("left_margin", "right_margin"):
        setattr(sec, m, Cm(2))
    f = co["finance"]
    years = f["years"]

    # 표지 · 점검표
    para(doc, PROJECT, 16, True, WD_ALIGN_PARAGRAPH.CENTER, 6)
    para(doc, "제출서식 (별지 제1~4호 · 제9호 · [붙임 8] · 수행실적 총괄표)", 12, False, WD_ALIGN_PARAGRAPH.CENTER, 14)
    para(doc, "제출서류 점검표 (입찰공고 R26BK01808991-000 · 제안요청서 기준)", 11, True)
    table(doc, [
        ["번호", "서류", "근거", "파일"],
        ["1", "정성제안서 (200쪽 이내)", "제안요청서 제안서 작성 목차", "nibr_정성제안서"],
        ["2", "발표자료 (제안요약서)", "공고 · 평가참고자료", "nibr_발표자료"],
        ["3", "일반현황 및 연혁 / 재무구조 및 최근 3년간 매출", "별지 제1호 · 제2호", "본 문서 (재무제표 첨부)"],
        ["4", "최근 3년간 사업 실적 / 조직 및 인력현황", "별지 제3호 · 제4호", "본 문서 (실적증명서 첨부)"],
        ["5", "2026년도 안전보건관리 계획서", "[붙임 8] · 평가 (4)", "본 문서"],
        ["6", "안전보건 서약서", "별지 제9호", "본 문서 (대표자 날인)"],
        ["7", "수행실적 총괄표 · 수행실적증명서", "평가기준 [주] 5 · 별지 제15호", "본 문서 (증명서 첨부)"],
        ["8", "요구사항 대비 수용여부 조견표", "제안요청서 요구사항", "요구사항_수용여부_조견표"],
        ["9", "정성 평가항목 조견표", "기술평가 항목", "정성제안서_평가항목_조견표"],
    ], widths=[1.2, 7, 4.8, 4], size=9)
    para(doc, "※ 제안서 · 가격 전자입찰서 접수 마감 2026/11/20 10:00 / 제안서 발표는 사업관리자(PM)가 함", 9)
    para(doc, "※ [회신대기] 수행실적 인정 범위 (민간 구축 · 공공 유지보수, 별지 제3호 '환경 정보화 분야만 기재') — 발주기관 회신 후 확정", 9)
    page_break(doc)

    # 별지 제1호
    title(doc, "[별지 제1호 서식]", "일반현황 및 연혁")
    fy, fm = g["founded"].replace("월", "").split("년")
    table(doc, [
        ["회 사 명", g["name"], "대 표 자", g["ceo"]],
        ["사 업 분 야", g["business"], "전 화 번 호", g["tel"]],
        ["주 소", g["address"], "회사설립년도", f"{fy.strip()}년 {fm.strip()}월"],
        ["해당부문 종사기간", g["biz_period"], "사업자등록번호", g["biz_no"]],
    ], widths=[3, 5.5, 3, 5.5], head=False)
    para(doc, "<주요연혁>", 10, True)
    table(doc, [["연도", "주요 연혁"]] + [[str(y), " / ".join(items)] for y, items in reversed(co["history"])],
          widths=[2, 15], size=8.5)
    page_break(doc)

    # 별지 제2호
    title(doc, "[별지 제2호 서식]", "재무구조 및 최근 3년간 매출")
    para(doc, "(단위 : 천원)", 9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    rev = [round(v / 1000) for v in f["revenue_total"]]
    ni = f["net_income"]; eq = f["total_equity"]
    roe = [f"{n / e * 100:.1f}%" for n, e in zip(ni, eq)]

    def m(vals):  # 백만원 값 → 재무제표 천원 값으로 교체할 자리
        return [f"【{v:,}백만】" for v in vals] + ["【 】", "【 】"]
    table(doc, [
        ["구분", f"{years[0]}년도 (M-2)", f"{years[1]}년도 (M-1)", f"{years[2]}년도 (M)", "합계", "평균"],
        ["1. 총 자 산"] + m(f["total_assets"]),
        ["2. 자 기 자 본"] + m(eq),
        ["3. 유 동 부 채", "【 】", "【 】", "【 】", "【 】", "【 】"],
        ["4. 고 정 부 채", "【 】", "【 】", "【 】", "【 】", "【 】"],
        ["5. 유 동 자 산", "【 】", "【 】", "【 】", "【 】", "【 】"],
        ["6. 당 기 순 이 익"] + m(ni),
        ["7. 매 출 원 가", "【 】", "【 】", "【 】", "【 】", "【 】"],
        ["8. 총 매 출 액"] + [f"{v:,}" for v in rev] + [f"{sum(rev):,}", f"{round(sum(rev) / 3):,}"],
        ["9. 자기자본 이익률 (당기순이익 / 자기자본)"] + roe + ["-", "-"],
        ["10. 부 채 율 ((유동부채+고정부채) / 자기자본)"] + list(f["debt_ratio"]) + ["-", "-"],
    ], widths=[4.6, 2.6, 2.6, 2.6, 2.3, 2.3], size=8.5)
    para(doc, "※ 총매출액은 회사 기준정보(원 단위)를 천원으로 환산 / 【 】은 재무제표(천원) 값으로 기재 — 백만원 값은 참고용", 9)
    para(doc, "※ 자기자본 이익률 · 부채율은 회사 기준정보(백만원) 기준 산출값, 재무제표 확정 후 재확인 / * 재무제표 첨부", 9)
    page_break(doc)

    # 별지 제3호
    title(doc, "[별지 제3호 서식]", "최근 3년간 사업 실적")
    para(doc, "(단위 : 천원)", 9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    table(doc, [["사 업 명", "사업기간", "계약금액", "발주처", "비고"]] +
          [["【사업명 — 실적 인정 범위 회신 후 확정】", "【 ~ 】", "【 】", "【발주처】", "[회신대기]"]] +
          [["【사업명】", "【 ~ 】", "【 】", "【발주처】", "[회신대기]"] for _ in range(4)],
          widths=[6, 3.4, 2.4, 3, 2.2], size=9)
    para(doc, "- 입찰공고일 기준 최근 3년간의 실적 (실적증명서 필히 첨부)", 9)
    para(doc, "- 현재 수행중인 업무도 포함하여 연도순(최근 계약일이 위)으로 기재하며, 생물자원 등 환경 정보화 분야만 기재", 9)
    para(doc, "- 하도급은 발주처가 승인한 경우에 한하며 비고란에 원 도급회사를 기재 / 공동도급은 제안사 지분만 기재", 9)
    para(doc, "[회신대기] '환경 정보화 분야만 기재' 적용 범위 질의 중 — 해당 실적이 없으면 '해당 없음'으로 제출", 9, True)
    page_break(doc)

    # 별지 제4호
    title(doc, "[별지 제4호 서식]", "조직 및 인력현황")
    para(doc, "□ 조 직", 10, True)
    table(doc, [
        ["조직", "주요 업무", "인원"],
        ["대표이사", "경영 · 사업 총괄", "1명"],
        ["사업관리실", "PM · PMO · 품질보증 · 보안", "【 】명"],
        ["기획 · 디자인", "UX 기획 · UI 디자인 · 인터랙션", "【 】명"],
        ["개발", "프론트엔드 · 백엔드 · 연계 · 모바일 웹", "【 】명"],
        ["AI · 데이터", "AI 모델 (STT · NLP · ML/DL) · 데이터 엔지니어링", "【 】명"],
        ["경영지원", "인사 · 재무 · 총무", "【 】명"],
        ["계", "", "【 】명 (【기준월】 기준)"],
    ], widths=[3.6, 9.4, 4], size=9)
    para(doc, "□ 본 사업 투입인력", 10, True)
    table(doc, [
        ["역할", "성명", "등급 · 자격", "주요 업무", "투입률"],
        ["사업PM", "【성명】", "【등급】", "사업 총괄 · 일정 · 위험 · 보고, 제안 발표", "【 】%"],
        ["품질보증", "【성명】", "【등급】", "단계별 품질 점검 · 산출물 검토 (수행팀과 분리)", "【 】%"],
        ["AI 모델", "【성명】", "【등급】", "STT · 소음 전처리 · 발화 구조화 · 예측 모델", "【 】%"],
        ["개발", "【성명】", "【등급】", "현장 입력 · 오프라인 동기화 · 분석 · GIS · 연계", "【 】%"],
        ["UX · 화면", "【성명】", "【등급】", "현장 · 관리자 화면 설계, 웹 접근성", "【 】%"],
        ["품질 · 보안", "【성명】", "【등급】", "단계별 시험 · 취약점 진단 · 시큐어코딩", "【 】%"],
    ], widths=[2.6, 2.4, 2.6, 7.4, 2], size=9)
    para(doc, "※ 제안업체 특성에 맞게 작성 / 투입인력 성명 · 등급 · 투입률은 착수계 제출 시 확정", 9)
    page_break(doc)

    # 붙임 8 안전보건관리 계획서
    para(doc, "[붙임 8]", 10, True)
    para(doc, "2026년도 안전 · 보건관리 계획서", 18, True, WD_ALIGN_PARAGRAPH.CENTER, 6)
    para(doc, f"2026. 【 】.     {g['name']}", 11, False, WD_ALIGN_PARAGRAPH.CENTER, 12)
    para(doc, "1. 도급 · 용역 · 위탁 사업 총괄", 11, True)
    table(doc, [
        ["사업명", PROJECT],
        ["사업기간", "계약 후 180일"],
        ["계약금액", "【 】원 (VAT 포함, 계약 시 기재)"],
        ["안전보건관리비 계상내역", "【 】원 — 안전보건교육 【 】 · 보험 【 】 · 안전장비(안전화 · 구명조끼 등) 【 】"],
    ], widths=[4.4, 12.6], head=False)
    para(doc, "2. 안전보건방침", 11, True)
    table(doc, [
        ["안전보건 경영방침", f"{g['name']}는 근로자의 안전과 건강을 경영의 최우선 가치로 두고, 본 사업의 현장 조사 동행 · 개발실 근무에서 발생할 수 있는 사고를 예방한다.\n"
         "1. 안전보건을 경영의 최우선으로 하며 체계적인 안전관리로 산업재해를 예방한다.\n"
         "2. 안전보건교육 · 훈련으로 임직원의 안전의식을 높이고 잠재 위험을 발굴 · 개선한다.\n"
         "3. 안전보건법규를 준수하고 이해관계자의 의견을 방침에 반영한다.\n"
         "4. 모든 임직원은 방침 준수에 참여하며 안전보건활동의 책임과 의무를 다한다.\n"
         "5. 공급자 · 계약자가 당사 안전보건 방침과 요구사항을 준수하도록 한다.\n"
         "6. 사업주는 방침과 추진목표를 근로자와 공유한다."],
        ["추진 목표", "○ 산업재해 제로(Zero) 및 안전한 작업환경 조성\n○ 근로자의 안전과 생명 보호\n○ 안전경영 중심 체계 구축\n○ 안전보건 예산 확보 및 개선 활동 실시\n○ 원 · 하청 상호 협력 (본 사업은 하도급 없음)"],
    ], widths=[3.4, 13.6], head=False, size=9)
    para(doc, f"2026년 【 】월 【 】일     {g['name']} 대표이사 {g['ceo']} (서명)", 10, False, WD_ALIGN_PARAGRAPH.RIGHT, 10)
    para(doc, "3. 사고발생시 비상연락체계", 11, True)
    table(doc, [
        ["순번", "성명", "소속", "직책", "역할", "전화번호"],
        ["1", g["ceo"], g["name"], "대표이사", "사업주 · 안전보건 총괄", "【 】"],
        ["2", "【성명】", "사업관리실", "사업PM", "관리감독자 · 현장 보고", "【 】"],
        ["3", "【성명】", "【 】", "【 】", "안전보건 담당", "【 】"],
        ["4", "【성명】", "경영지원", "【 】", "재해 신고 · 보험 처리", "【 】"],
        ["5", "【성명】", "국립생물자원관 국가조류연구팀", "【 】", "발주기관 감독", "【 】"],
    ], widths=[1.2, 2.4, 4.2, 2.2, 4, 3], size=9)
    para(doc, "소방서 : 119 / 【관할 보건소】 : 【 】 / 【인근 종합병원】 : 【 】  (현장 조사 지역별로 추가 기재)", 9)
    para(doc, "4. 사업 현황", 11, True)
    para(doc, "◦ 작업장소 : 당사 사무실 · 원격지 개발실 (【주소】), 발주기관 (인천 서구 환경로 42), 현장 조사 동행 (해안 · 습지 · 【조사 지역】)", 9.5)
    para(doc, "◦ 근무형태 : 주간 근무 (야간근무 없음, 현장 조사 동행은 일출 전후 조사 일정에 따라 【 】)", 9.5)
    para(doc, "◦ 작업 참여자 수 : 【 】명", 9.5)
    table(doc, [
        ["구분", "위험요인", "방호조치", "점검주기", "담당자"],
        ["PC · 모니터 (VDT)", "근골격계 질환 · 눈 피로", "50분 작업 후 휴식, 작업대 높이 조정", "월 1회", "【성명】"],
        ["GPU 서버 · 전기설비", "과부하 · 누전 · 화재", "전용 회로 · 누전차단기 · 소화기 비치", "월 1회", "【성명】"],
        ["현장 녹음 장비 · 차량", "해안 · 습지 미끄러짐, 기상 악화, 차량 이동", "2인 1조, 기상 확인, 안전화 · 구명조끼", "동행 시마다", "【성명】"],
    ], widths=[3.2, 3.8, 5.4, 2, 2.6], size=9)
    para(doc, "5. 위험성 평가 계획 (상주 도급인만 해당)", 11, True)
    para(doc, "◦ (해당 여부) 본 사업은 발주기관 상주 작업이 아님 — 현장 조사 동행 · 시범운영 시 동행 전 위험성 평가 실시", 9.5)
    para(doc, "◦ (평가일정) 착수 후 【 】일 이내 최초 평가, 현장 동행 일정 확정 시 수시 평가 / (평가범위) 사무실 · 개발실 · 현장 동행 구간", 9.5)
    table(doc, [
        ["구분", "직위(직급)", "성명", "역할"],
        ["안전총괄책임자", "대표이사", g["ceo"], "위험성평가 실시 총괄"],
        ["안전관리책임자", "사업PM", "【성명】", "총괄관리자 보좌"],
        ["안전 담당자", "【 】", "【성명】", "유해 · 위험요인 파악, 감소대책 수립"],
        ["근로자", "투입인력", "【성명】", "위험성평가 참여"],
    ], widths=[3.6, 3, 3, 7.4], size=9)
    para(doc, "6. 작업 안전점검 및 모니터링 계획", 11, True)
    table(doc, [
        ["점검시기", "점검내용", "점검방법", "점검자", "모니터링 담당자"],
        ["작업 전", "기상 · 물때 확인, 안전장비 지급, 비상연락망 공유", "체크리스트", "사업PM", "안전 담당자"],
        ["작업 중", "2인 1조 유지, 위험 구역 접근 통제, 정시 연락", "현장 확인", "동행 책임자", "안전 담당자"],
        ["작업 후", "인원 · 장비 복귀 확인, 이상 사항 기록", "점검표 기록", "사업PM", "안전 담당자"],
    ], widths=[2, 6.4, 2.6, 2.6, 3.4], size=9)
    para(doc, "7. 안전보건교육 계획", 11, True)
    table(doc, [
        ["일시", "교육내용", "교육대상", "장소", "교육지표"],
        ["착수 시", "안전보건 방침 · 비상연락체계 · 작업안전수칙", "투입인력 전원", "당사 회의실", "이수율 100%"],
        ["현장 동행 전", "해안 · 습지 안전수칙, 기상 악화 · 낙상 대응", "동행 인원", "출발 전 현장", "이수율 100%"],
        ["분기 1회", "VDT 증후군 예방 · 전기 · 화재 안전", "투입인력 전원", "당사 회의실", "이수율 100%"],
    ], widths=[2.2, 6.4, 2.8, 2.6, 3], size=9)
    para(doc, "8. 안전작업 허가 (허가대상 안전작업 해당)", 11, True)
    para(doc, "◦ (허가대상) 해당 없음 — 화기 · 고소 · 밀폐공간 작업 없음 (현장 동행은 사전 승인 후 실시)", 9.5)
    para(doc, "◦ (허가절차 · 기록) 현장 동행 계획서 작성 → 사업PM 검토 → 발주기관 승인, 결과 기록 1년 보관", 9.5)
    para(doc, "9. 안전사고별 비상대책", 11, True)
    table(doc, [
        ["사고 유형", "상황 전파", "초기 대응", "후속 조치"],
        ["화재 (사무실 · 개발실)", "발견자 119 신고 · 사내 전파", "소화기 초기 진화 · 대피 유도 · 전원 차단", "현장 조사 · 보고 · 재발 방지"],
        ["현장 동행 사고 (낙상 · 고립)", "동행자 119 · 사업PM 보고", "응급조치 · 안전 지대 이동", "발주기관 보고 · 원인 분석"],
        ["기상 악화", "기상특보 확인 즉시 전파", "조사 중단 · 즉시 철수", "일정 재조정"],
        ["정전 · 설비 장애", "담당자 전파", "UPS 전환 · 데이터 보호", "복구 · 백업 확인"],
    ], widths=[3.6, 4, 5, 4.4], size=9)
    page_break(doc)

    # 별지 제9호
    para(doc, "[별지 제9호 서식]", 10, True)
    para(doc, "안 전 보 건 서 약 서", 18, True, WD_ALIGN_PARAGRAPH.CENTER, 12)
    para(doc, f"계약건명 : {PROJECT}", 11)
    para(doc, "계약기간 : 계약 후 180일 (【 】 ~ 【 】)", 11)
    para(doc, f"업 체 명 : {g['name']}  (대표자) {g['ceo']}", 11, after=10)
    para(doc, "위 업체는 본 계약을 수행함에 있어 아래의 안전수칙 사항을 비롯하여 안전관련법령 및 관련 규정을 준수하고 위반 시 그에 따른 책임을 인정하며 감수할 것을 서약합니다.", 10.5, after=10)
    para(doc, "< 작 업 안 전 수 칙 >", 11, True, WD_ALIGN_PARAGRAPH.CENTER, 8)
    for t in [
        "1. 근로자는 작업에 대한 지침을 숙지하고 관련 안전교육을 이수하여야 하며, 안전보호구를 반드시 착용하고 안전시설 및 작업기구는 사용 전에 점검한다.",
        "2. 근로자는 관리자의 지시사항을 준수하고 작업장 내 뛰어다니는 등 임의행동을 삼가며, 유류 주변에서는 화기 취급에 절대 주의한다.",
        "3. 근로자는 작업 전과 작업 시 절대 음주를 하지 않아야 하고 작업장 내 담배는 반드시 지정된 장소에서만 피운다.",
        "4. 근로자는 작업 상황을 수시 보고하고 위급상황 발생 시 즉시 작업중단 및 응급조치를 실시하며 관계 기관에 신고하여야 한다.",
        "5. 근로자는 작업 주변을 항상 정리정돈하고 청결하게 유지하여야 하며, 작업 후 발생한 폐기물은 업체에서 적법하게 처리하여야 한다.",
        "6. 기타 세부 안전수칙은 산업안전보건 기준에 관한 규칙(2021.11.19.시행) 등 관련 규정에 따른다.",
    ]:
        para(doc, t, 10, after=6)
    para(doc, "2026년    【 】월    【 】일", 11, False, WD_ALIGN_PARAGRAPH.CENTER, 14)
    para(doc, f"대 표 자 : {g['ceo']}   (인)", 11, False, WD_ALIGN_PARAGRAPH.RIGHT)
    page_break(doc)

    # 수행실적 총괄표 (별지 제15호)
    title(doc, "[별지 제15호 서식]", "수행실적 총괄표")
    para(doc, "(단위 : 원, 입찰공고일 기준 최근 3년 · 최근 계약일이 위)", 9, align=WD_ALIGN_PARAGRAPH.RIGHT)
    recs = [r for r in co["records"] if r.get("kind") in ("구축", "AI", "유지보수")]
    recs.sort(key=lambda r: str(r["start"]), reverse=True)
    rows = [["번호", "사업명", "발주처", "사업기간", "계약금액", "상태", "인정 여부"]]
    for i, r in enumerate(recs, 1):
        rows.append([str(i), r["name"], r["client"], f"{r['start']} ~ {r['end']}", f"{int(r['amount']):,}",
                     r.get("status", ""), "[회신대기]"])
    rows.append(["", "합계 (인정 실적만 합산)", "", "", "【 】", "", ""])
    table(doc, rows, widths=[1, 5.4, 2.6, 3.2, 2.4, 1.2, 1.8], size=7.5)
    para(doc, "※ 정량평가 수행실적: 최근 3년간 행정 및 공공기관 정보시스템 구축 사업수행실적(금액 기준) 합산, 이행 완료 시점 기준 (공고일 다음날 이후 완료분 제외)", 8.5)
    para(doc, "※ 위 목록은 회사 실적 대장(h9_company.yaml)의 구축 · AI · 유지보수 실적 후보 — 민간 발주 구축 · 공공 유지보수 인정 여부 [회신대기], 회신 후 인정 실적만 남겨 확정", 8.5)
    para(doc, "※ 별지 제16호 수행실적증명서 또는 계약서 · 세금계산서(민간 발주), 한국소프트웨어산업협회 실적확인서 첨부", 8.5)

    cp = doc.core_properties
    cp.author = cp.last_modified_by = document_author()
    cp.title = f"{PROJECT} 제출서식"
    out = HERE / "output" / "제출서식.docx"
    doc.save(str(out))
    print("저장:", out)


if __name__ == "__main__":
    main()
