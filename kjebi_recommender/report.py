"""추천 결과를 Markdown / HTML 리포트로 만든다."""
from __future__ import annotations

from datetime import date
from html import escape

from .parser import Notice
from .scorer import GRADE_EXCLUDED, GRADE_INFO, GRADE_REVIEW, GRADE_STRONG

GRADE_ICON = {GRADE_STRONG: "🟢", GRADE_REVIEW: "🟡", GRADE_INFO: "⚪", GRADE_EXCLUDED: "⛔"}


def _budget(n: Notice) -> str:
    if not n.budget:
        return "-"
    if n.budget >= 100_000_000:
        return f"{n.budget / 100_000_000:.1f}억"
    return f"{n.budget / 10_000:,.0f}만"


def _deadline(n: Notice, today: date) -> str:
    if not n.deadline:
        return "-"
    label = "의견마감" if n.kind == "pre-spec" else "마감"
    return f"{label} {n.deadline:%m-%d} (D-{(n.deadline - today).days})"


def _region(n: Notice) -> str:
    return n.region if len(n.region) <= 20 else n.region[:20] + "…"


def build_markdown(notices: list[Notice], today: date, target_dates: list[str]) -> str:
    recommended = [n for n in notices if n.grade in (GRADE_STRONG, GRADE_REVIEW)]
    lines = [
        f"# 크제비 추천 공고 리포트 ({today:%Y-%m-%d})",
        "",
        f"- 대상 게시일: {', '.join(target_dates) or '-'}",
        f"- 전체 {len(notices)}건 중 추천 {len(recommended)}건 "
        f"(강력 {sum(n.grade == GRADE_STRONG for n in notices)} / 검토 {sum(n.grade == GRADE_REVIEW for n in notices)})",
        "",
    ]
    for grade in (GRADE_STRONG, GRADE_REVIEW, GRADE_INFO, GRADE_EXCLUDED):
        group = [n for n in notices if n.grade == grade]
        if not group:
            continue
        lines += [f"## {GRADE_ICON[grade]} {grade} ({len(group)})", ""]
        if grade in (GRADE_INFO, GRADE_EXCLUDED):
            for n in group:
                lines.append(f"- [{n.kind_label}] [{n.title}]({n.url}) · {n.agency} · {_budget(n)} — {', '.join(n.reasons) or '-'}")
            lines.append("")
            continue
        for n in group:
            lines += [
                f"### [{n.title}]({n.url})",
                f"- {n.kind_label} `{n.notice_id}` · 점수 **{n.score}**",
                f"- {n.agency} · {_budget(n)} · {_deadline(n, today)}"
                + (f" · 지역 {_region(n)}" if n.region else ""),
                f"- 근거: {', '.join(n.reasons) or '-'}",
                "",
            ]
    return "\n".join(lines)


def build_html(notices: list[Notice], today: date, target_dates: list[str]) -> str:
    rows = []
    for n in notices:
        if n.grade == GRADE_EXCLUDED:
            continue
        rows.append(
            "<tr>"
            f"<td>{GRADE_ICON[n.grade]} {escape(n.grade)}<br><small>{n.score}점</small></td>"
            f"<td><small>{escape(n.kind_label)}</small><br><a href=\"{escape(n.url)}\">{escape(n.title)}</a>"
            f"<br><small style=\"color:#666\">{escape(', '.join(n.reasons))}</small></td>"
            f"<td>{escape(n.agency)}</td><td style=\"white-space:nowrap\">{_budget(n)}</td>"
            f"<td style=\"white-space:nowrap\">{escape(_deadline(n, today))}</td>"
            "</tr>"
        )
    excluded = sum(n.grade == GRADE_EXCLUDED for n in notices)
    return (
        "<div style=\"font-family:sans-serif;font-size:14px\">"
        f"<h2>크제비 추천 공고 ({today:%Y-%m-%d})</h2>"
        f"<p>대상 게시일: {escape(', '.join(target_dates) or '-')} · 전체 {len(notices)}건 (제외 {excluded}건 생략)</p>"
        "<table border=\"1\" cellpadding=\"6\" cellspacing=\"0\" style=\"border-collapse:collapse\">"
        "<tr style=\"background:#f3f3f3\"><th>등급</th><th>공고</th><th>수요기관</th><th>예산</th><th>마감</th></tr>"
        + "".join(rows)
        + "</table></div>"
    )
