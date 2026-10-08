from datetime import date
from pathlib import Path

import yaml

from kjebi_recommender.__main__ import main
from kjebi_recommender.parser import parse_email_html
from kjebi_recommender.scorer import GRADE_EXCLUDED, GRADE_STRONG, rank

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "sample_email.html"
TODAY = date(2026, 10, 8)


def _parse():
    return parse_email_html(FIXTURE.read_text(encoding="utf-8"))


def _config():
    return yaml.safe_load((ROOT / "config.yaml").read_text(encoding="utf-8"))


def test_parse_bid_fields():
    target, notices = _parse()
    assert target == "2026-10-07"
    assert len(notices) == 13
    n = next(n for n in notices if n.notice_id == "R26BK01759644-000")
    assert n.kind == "bid"
    assert n.title == "인천대학교 AI 기반 교육지원시스템 4단계 구축 용역 계약"
    assert n.contract_method == "제한경쟁"
    assert n.budget == 367_272_727
    assert n.agency == "인천대학교"
    assert n.region == "전국"
    assert n.deadline == date(2026, 10, 19)


def test_parse_bid_demand_agency_is_last():
    _, notices = _parse()
    n = next(n for n in notices if n.notice_id == "R26BK01757403-000")
    assert n.agency == "문화체육관광부 한국예술종합학교"


def test_parse_pre_spec_fields():
    _, notices = _parse()
    n = next(n for n in notices if n.notice_id == "R26BD00282365")
    assert n.kind == "pre-spec"
    assert n.title == "AI 서비스 플랫폼 구축"
    assert n.budget == 3_019_200_680
    assert n.agency == "(재)제주테크노파크"
    assert n.deadline == date(2026, 10, 14)


def test_rank_dedupes_and_grades():
    _, notices = _parse()
    ranked = rank(notices, _config(), TODAY)
    ids = {n.notice_id for n in ranked}
    # 정정공고는 최신 차수만, 같은 제목/기관 중복은 하나만
    assert "R26BK01759604-001" in ids and "R26BK01759604-000" not in ids
    assert len([n for n in ranked if "임팩트 챌린지" in n.title]) == 1

    by_id = {n.notice_id: n for n in ranked}
    assert by_id["R26BD00282365"].grade == GRADE_STRONG          # AI 서비스 플랫폼 구축
    assert by_id["R26BK01759843-000"].grade == GRADE_EXCLUDED    # 장비 설치
    assert by_id["R26BD00282087"].grade == GRADE_EXCLUDED        # 폐기물
    assert any("참여지역 제한" in r for r in by_id["R26BK01761269-000"].reasons)
    assert any("마감 임박" in r for r in by_id["R26BK01759604-001"].reasons)
    assert ranked[0].grade == GRADE_STRONG


def test_cli_with_local_file(tmp_path, capsys):
    rc = main(["--config", str(ROOT / "config.yaml"), "--input", str(FIXTURE),
               "--today", "2026-10-08", "--out", str(tmp_path)])
    assert rc == 0
    out = capsys.readouterr().out
    assert "크제비 추천 공고 리포트 (2026-10-08)" in out
    assert (tmp_path / "kjebi-20261008.md").exists()
