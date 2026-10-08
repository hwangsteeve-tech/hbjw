"""설정(config.yaml) 기준으로 공고에 점수와 추천 등급을 매긴다."""
from __future__ import annotations

import re
from datetime import date

from .parser import Notice

GRADE_STRONG = "강력 추천"
GRADE_REVIEW = "검토 추천"
GRADE_INFO = "참고"
GRADE_EXCLUDED = "제외"
GRADE_ORDER = [GRADE_STRONG, GRADE_REVIEW, GRADE_INFO, GRADE_EXCLUDED]


def _keyword_pattern(keyword: str) -> re.Pattern:
    escaped = re.escape(keyword)
    if keyword.isascii():
        # 영문 약어(UI, AX 등)가 다른 단어 일부에 매칭되지 않도록 경계 지정
        return re.compile(rf"(?<![A-Za-z]){escaped}(?![A-Za-z])", re.IGNORECASE)
    return re.compile(escaped, re.IGNORECASE)


def _won(amount: int) -> str:
    if amount >= 100_000_000:
        return f"{amount / 100_000_000:.1f}억"
    return f"{amount / 10_000:,.0f}만"


def dedupe(notices: list[Notice]) -> list[Notice]:
    """정정공고(-001 등)·재발송 메일로 생긴 중복을 정리한다. 최신 차수만 남긴다."""
    by_id: dict[tuple[str, str], Notice] = {}
    for n in notices:
        key = (n.kind, n.base_id)
        if key not in by_id or n.revision > by_id[key].revision:
            by_id[key] = n
    by_title: dict[tuple[str, str, str], Notice] = {}
    for n in by_id.values():
        key = (n.kind, re.sub(r"\W", "", n.title), n.agency)
        if key not in by_title or n.notice_id > by_title[key].notice_id:
            by_title[key] = n
    return list(by_title.values())


def score_notice(n: Notice, config: dict, today: date) -> Notice:
    score = 0
    reasons: list[str] = []
    title = n.title

    for kw in config.get("exclude", []):
        if _keyword_pattern(kw).search(title):
            n.score, n.grade, n.reasons = 0, GRADE_EXCLUDED, [f"제외 키워드 '{kw}'"]
            return n

    for kw, weight in {**config.get("positive", {}), **config.get("negative", {})}.items():
        if _keyword_pattern(kw).search(title):
            score += weight
            reasons.append(f"{kw} {weight:+d}")

    budget_cfg = config.get("budget", {})
    if n.budget:
        if n.budget < budget_cfg.get("min", 0):
            score -= 3
            reasons.append(f"소규모 예산({_won(n.budget)}) -3")
        elif budget_cfg.get("preferred_min", 0) <= n.budget <= budget_cfg.get("preferred_max", float("inf")):
            score += 2
            reasons.append(f"선호 예산({_won(n.budget)}) +2")

    if n.kind == "bid" and n.deadline:
        days_left = (n.deadline - today).days
        if days_left < 0:
            n.score, n.grade, n.reasons = 0, GRADE_EXCLUDED, ["입찰 마감 지남"]
            return n
        min_days = config.get("deadline", {}).get("min_days_left", 0)
        if days_left < min_days:
            score -= 3
            reasons.append(f"마감 임박(D-{days_left}) -3")

    regions_ok = config.get("regions_ok", [])
    if n.region and regions_ok and not any(r in n.region for r in regions_ok):
        score -= 5
        reasons.append("참여지역 제한 -5")

    if n.revision > 0:
        reasons.append("정정/재공고")
    if n.kind == "pre-spec":
        reasons.append("사전규격: 본공고 전 준비 가능")

    thresholds = config.get("thresholds", {})
    if score >= thresholds.get("strong", 8):
        grade = GRADE_STRONG
    elif score >= thresholds.get("review", 4):
        grade = GRADE_REVIEW
    else:
        grade = GRADE_INFO
    n.score, n.grade, n.reasons = score, grade, reasons
    return n


def rank(notices: list[Notice], config: dict, today: date) -> list[Notice]:
    scored = [score_notice(n, config, today) for n in dedupe(notices)]
    return sorted(scored, key=lambda n: (GRADE_ORDER.index(n.grade), -n.score, -(n.budget or 0)))
