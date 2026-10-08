"""Gmail IMAP으로 크제비 메일을 가져오고, SMTP로 리포트를 보낸다."""
from __future__ import annotations

import email
import imaplib
import smtplib
from dataclasses import dataclass
from datetime import date, timedelta
from email.header import decode_header, make_header
from email.message import EmailMessage
from pathlib import Path

KJEBI_SENDER = "kjebi.com"


@dataclass
class KjebiMail:
    subject: str
    date: str
    html: str


def _decode(value: str | None) -> str:
    return str(make_header(decode_header(value))) if value else ""


def _html_part(msg: email.message.Message) -> str:
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            payload = part.get_payload(decode=True) or b""
            return payload.decode(part.get_content_charset() or "utf-8", errors="replace")
    return ""


def message_from_bytes(raw: bytes) -> KjebiMail:
    msg = email.message_from_bytes(raw)
    return KjebiMail(_decode(msg["Subject"]), _decode(msg["Date"]), _html_part(msg))


def load_local(path: Path) -> KjebiMail:
    """저장해 둔 .eml 또는 .html 파일을 읽는다 (테스트/수동 실행용)."""
    data = path.read_bytes()
    if path.suffix.lower() == ".eml":
        return message_from_bytes(data)
    return KjebiMail(path.name, "", data.decode("utf-8", errors="replace"))


def fetch_imap(user: str, password: str, days: int, host: str = "imap.gmail.com",
               mailbox: str = "INBOX") -> list[KjebiMail]:
    # days=1 이면 오늘 받은 메일만 (IMAP SINCE 는 날짜 단위, 당일 포함)
    since = (date.today() - timedelta(days=max(days, 1) - 1)).strftime("%d-%b-%Y")
    with imaplib.IMAP4_SSL(host) as imap:
        imap.login(user, password)
        imap.select(mailbox, readonly=True)
        status, data = imap.search(None, "FROM", f'"{KJEBI_SENDER}"', "SINCE", since)
        if status != "OK":
            raise RuntimeError(f"IMAP 검색 실패: {status}")
        mails = []
        for num in data[0].split():
            status, msg_data = imap.fetch(num, "(RFC822)")
            if status == "OK" and msg_data and isinstance(msg_data[0], tuple):
                mails.append(message_from_bytes(msg_data[0][1]))
        return mails


def send_report(user: str, password: str, to: str, subject: str, text: str, html: str,
                host: str = "smtp.gmail.com", port: int = 465) -> None:
    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = to
    msg["Subject"] = subject
    msg.set_content(text)
    msg.add_alternative(html, subtype="html")
    with smtplib.SMTP_SSL(host, port) as smtp:
        smtp.login(user, password)
        smtp.send_message(msg)
