"""사용법:
    python -m kjebi_recommender                      # Gmail에서 최근 1일치 크제비 메일 분석
    python -m kjebi_recommender --days 3 --send      # 최근 3일치 분석 후 리포트 메일 발송
    python -m kjebi_recommender --input mail.eml     # 저장한 메일 파일로 분석

환경변수: GMAIL_ADDRESS, GMAIL_APP_PASSWORD, REPORT_TO(선택, 기본은 본인)
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import date
from pathlib import Path

import yaml

from . import mail_io
from .parser import parse_email_html
from .report import build_html, build_markdown
from .scorer import GRADE_REVIEW, GRADE_STRONG, rank


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="kjebi_recommender", description="크제비 알림메일 추천 공고 선별")
    ap.add_argument("--config", default="config.yaml", type=Path)
    ap.add_argument("--input", nargs="*", type=Path, help="IMAP 대신 로컬 .eml/.html 파일 사용")
    ap.add_argument("--days", type=int, default=1, help="최근 며칠치 메일을 볼지 (기본 1)")
    ap.add_argument("--out", type=Path, help="리포트(.md)를 저장할 디렉터리")
    ap.add_argument("--send", action="store_true", help="리포트를 이메일로 발송")
    ap.add_argument("--today", type=date.fromisoformat, default=date.today(), help="기준일 (YYYY-MM-DD)")
    args = ap.parse_args(argv)

    config = yaml.safe_load(args.config.read_text(encoding="utf-8")) or {}
    user = os.environ.get("GMAIL_ADDRESS", "")
    password = os.environ.get("GMAIL_APP_PASSWORD", "")

    if args.input:
        mails = [mail_io.load_local(p) for p in args.input]
    else:
        if not (user and password):
            print("GMAIL_ADDRESS / GMAIL_APP_PASSWORD 환경변수를 설정하거나 --input 을 사용하세요.", file=sys.stderr)
            return 2
        mails = mail_io.fetch_imap(user, password, args.days,
                                   mailbox=os.environ.get("IMAP_MAILBOX", "INBOX"))

    notices, target_dates = [], []
    for m in mails:
        target, items = parse_email_html(m.html)
        if target and target not in target_dates:
            target_dates.append(target)
        notices.extend(items)
    if not mails:
        print("분석할 크제비 메일이 없습니다.", file=sys.stderr)

    ranked = rank(notices, config, args.today)
    markdown = build_markdown(ranked, args.today, sorted(target_dates))
    print(markdown)

    if args.out:
        args.out.mkdir(parents=True, exist_ok=True)
        (args.out / f"kjebi-{args.today:%Y%m%d}.md").write_text(markdown, encoding="utf-8")

    if summary := os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(summary, "a", encoding="utf-8") as f:
            f.write(markdown + "\n")

    if args.send:
        if not (user and password):
            print("메일 발송에는 GMAIL_ADDRESS / GMAIL_APP_PASSWORD 가 필요합니다.", file=sys.stderr)
            return 2
        picks = sum(n.grade in (GRADE_STRONG, GRADE_REVIEW) for n in ranked)
        subject = f"[크제비 추천] {args.today:%Y-%m-%d} 추천 {picks}건 / 전체 {len(ranked)}건"
        mail_io.send_report(user, password, os.environ.get("REPORT_TO") or user, subject,
                            markdown, build_html(ranked, args.today, sorted(target_dates)))
        print(f"리포트 발송 완료 → {os.environ.get('REPORT_TO') or user}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
