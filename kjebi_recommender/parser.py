"""크제비 알림메일 HTML에서 공고 목록을 추출한다."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime

from bs4 import BeautifulSoup

DETAIL_RE = re.compile(r"kjebi\.com/search/smart/(bid|pre-spec|order-plan)/detail/([^/?#\s\"']+)")
KIND_LABEL = {"bid": "입찰공고", "pre-spec": "사전규격", "order-plan": "발주계획"}

BUDGET_RE = re.compile(r"(추정가격|배정예산액|기초금액|사업금액|예산액?)\s*:?\s*([\d,]{4,})")
REGION_RE = re.compile(r"참여지역\s*\[([^\]]*)\]")
POSTED_RE = re.compile(r"(?:게시일시|진행일자|공고일자?)\s*:?\s*(\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2})?)")
DEADLINE_RE = re.compile(r"(?:입찰마감일|의견등록마감일|마감일)\s*:?\s*(\d{4}-\d{2}-\d{2})")
TARGET_DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})\s*에 게시된 자료")
CONTRACT_METHODS = r"일반경쟁|제한경쟁|지명경쟁|수의계약"
TRAILING_LABEL_RE = {
    "bid": re.compile(rf"\s+({CONTRACT_METHODS})$"),
    "pre-spec": re.compile(r"\s+(\S*용역|물품|공사)$"),
}
FIELD_LABELS = ("담당자", "진행일자", "의견등록마감일", "참여지역", "게시일시", "입찰마감일")


@dataclass
class Notice:
    kind: str
    notice_id: str
    title: str
    url: str
    raw: str
    agency: str = ""
    budget: int | None = None
    region: str = ""
    posted: str = ""
    deadline: date | None = None
    contract_method: str = ""
    # 스코어링 결과
    score: int = 0
    grade: str = ""
    reasons: list[str] = field(default_factory=list)

    @property
    def kind_label(self) -> str:
        return KIND_LABEL.get(self.kind, self.kind)

    @property
    def base_id(self) -> str:
        return self.notice_id.split("-")[0]

    @property
    def revision(self) -> int:
        parts = self.notice_id.split("-")
        return int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("*", " ")).strip()


def _item_text(anchor, notice_id: str) -> str:
    """앵커에서 위로 올라가며 공고번호가 포함된 가장 작은 요소의 텍스트를 찾는다."""
    node = anchor
    fallback = ""
    for _ in range(6):
        if node is None:
            break
        text = _clean(node.get_text(" "))
        if notice_id in text and len(text) < 4000:
            return text
        if text and not fallback:
            fallback = text
        node = node.parent
    return fallback


def _parse_fields(n: Notice) -> None:
    t = n.raw
    if m := BUDGET_RE.search(t):
        n.budget = int(m.group(2).replace(",", ""))
    if m := REGION_RE.search(t):
        n.region = m.group(1).strip()
    if m := POSTED_RE.search(t):
        n.posted = m.group(1)
    if m := DEADLINE_RE.search(t):
        n.deadline = datetime.strptime(m.group(1), "%Y-%m-%d").date()

    segments = [s.strip() for s in t.split(" / ")]
    head = segments[0]
    title = head.split(n.notice_id, 1)[1].strip() if n.notice_id in head else head
    if pattern := TRAILING_LABEL_RE.get(n.kind):
        if m := pattern.search(title):
            if n.kind == "bid":
                n.contract_method = m.group(1)
            title = title[: m.start()].strip()
    n.title = title or n.notice_id

    budget_idx = next((i for i, s in enumerate(segments) if BUDGET_RE.match(s)), None)
    if budget_idx is not None:
        candidates = []
        for s in segments[budget_idx + 1:]:
            if s.startswith(FIELD_LABELS):
                break
            candidates.append(s)
        if candidates:
            # 입찰공고는 "공고기관 / 수요기관" 순, 사전규격은 "수요기관 / 조달청" 순
            n.agency = candidates[-1] if n.kind == "bid" else candidates[0]


def parse_email_html(html: str) -> tuple[str | None, list[Notice]]:
    """알림메일 HTML을 파싱해 (검색 대상일, 공고 목록)을 반환한다."""
    soup = BeautifulSoup(html, "html.parser")
    full_text = _clean(soup.get_text(" "))
    target = TARGET_DATE_RE.search(full_text)

    notices: list[Notice] = []
    seen: set[tuple[str, str]] = set()
    for a in soup.find_all("a", href=True):
        m = DETAIL_RE.search(a["href"])
        if not m:
            continue
        kind, notice_id = m.group(1), m.group(2)
        if (kind, notice_id) in seen:
            continue
        seen.add((kind, notice_id))
        raw = _item_text(a, notice_id)
        n = Notice(kind=kind, notice_id=notice_id, title="", url=f"https://www.{m.group(0)}", raw=raw)
        _parse_fields(n)
        notices.append(n)
    return (target.group(1) if target else None), notices
