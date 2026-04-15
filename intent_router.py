"""Outlook intent router for the orchestrator chat flow.

Formatters, date-filter parsing and calendar rendering are ported
verbatim from MiBuddy's `backend/agents/outlook_agent.py` so the
agentcore output matches MiBuddy's look byte-for-byte (section
headers, emoji, table layout, wording).

Intent detection is a lightweight keyword-based classifier (vs.
MiBuddy's LangGraph + Azure-OpenAI classifier). It handles the
common cases the user asks about — "show emails", "meetings",
"search" — while avoiding a big LLM dependency on agentcore.
"""
from __future__ import annotations

import logging
import re
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

from agentcore.services.outlook_orch.outlook_service import OutlookService
from agentcore.services.outlook_orch.token_manager import outlook_token_manager

logger = logging.getLogger(__name__)


# ── Intent keyword patterns (pragmatic replacement for MiBuddy's LLM
#    sub-intent classifier) ──────────────────────────────────────────

_CALENDAR_PAT = re.compile(
    r"\b("
    r"(?:my\s+)?calendar|"
    r"meeting(?:s)?|"
    r"appointment(?:s)?|"
    r"schedule|"
    r"events?"
    r")\b",
    re.IGNORECASE,
)

_EMAIL_SEARCH_PAT = re.compile(
    r"\b(?:search|find|look\s+for|any)\b.*\b(?:email|mail|message)s?\b",
    re.IGNORECASE,
)

_EMAIL_READ_PAT = re.compile(
    r"\b("
    r"(?:recent|latest|last|unread|new)\s+(?:email|mail|message)s?|"
    r"show\s+(?:me\s+)?(?:my\s+)?(?:email|mail|inbox|message)s?|"
    r"check\s+(?:my\s+)?(?:email|mail|inbox)|"
    r"read\s+(?:my\s+)?(?:email|mail|inbox)|"
    r"how\s+many\s+(?:email|mail|message)s?|"
    r"emails?\s+(?:from\s+)?(?:today|yesterday)|"
    r"inbox"
    r")\b",
    re.IGNORECASE,
)

_UNREAD_PAT = re.compile(
    r"\b(unread|unseen|new|haven'?t\s+read)\b", re.IGNORECASE,
)

# Date-filter extraction (ports MiBuddy's `date_filter` concept)
_DATE_FILTER_PATS: List[Tuple[re.Pattern, str]] = [
    (re.compile(r"\btoday\b", re.IGNORECASE), "today"),
    (re.compile(r"\byesterday\b", re.IGNORECASE), "yesterday"),
    (re.compile(r"\bthis\s+week\b", re.IGNORECASE), "this_week"),
    (re.compile(r"\blast\s+(\d+)\s+days?\b", re.IGNORECASE), "last_N_days"),
]


class OutlookIntent:
    """Container for a detected intent + its rendered markdown."""

    def __init__(self, kind: str, markdown: str) -> None:
        self.kind = kind  # "email_list" | "email_search" | "calendar" | "error"
        self.markdown = markdown

    def __repr__(self) -> str:  # pragma: no cover
        return f"<OutlookIntent {self.kind}>"


# ═══════════════════════════════════════════════════════════════════
#   MiBuddy-ported helpers (dates, formatters)
# ═══════════════════════════════════════════════════════════════════

def _format_date(dt: datetime) -> str:
    """MiBuddy: cross-platform date formatting without leading zeros."""
    return dt.strftime("%A, %B {day}, %Y").replace("{day}", str(dt.day))


def _format_short_date(dt: datetime) -> str:
    """MiBuddy: cross-platform short date formatting."""
    return dt.strftime("%b {day}").replace("{day}", str(dt.day))


def _parse_email_date_filter(
    date_filter: str,
) -> Tuple[Optional[str], Optional[str], str]:
    """Port of MiBuddy's `_parse_email_date_filter`.

    Accepts "today", "yesterday", "this_week", "last_N_days", or
    "YYYY-MM-DD" and returns (received_after_iso, received_before_iso,
    human_readable_label).
    """
    today = datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0)

    if date_filter == "today":
        start, end, label = today, today + timedelta(days=1), "today"
    elif date_filter == "yesterday":
        start, end, label = today - timedelta(days=1), today, "yesterday"
    elif date_filter == "this_week":
        start = today - timedelta(days=today.weekday())
        end = today + timedelta(days=1)
        label = "this week"
    elif date_filter.startswith("last_") and date_filter.endswith("_days"):
        m = re.match(r"last_(\d+)_days", date_filter)
        if m:
            n = int(m.group(1))
            start = today - timedelta(days=n)
            end = today + timedelta(days=1)
            label = f"the last {n} days"
        else:
            start, end, label = today, today + timedelta(days=1), "today"
    elif re.match(r"\d{4}-\d{2}-\d{2}", date_filter):
        try:
            start = datetime.strptime(date_filter, "%Y-%m-%d")
            end = start + timedelta(days=1)
            label = _format_date(start)
        except ValueError:
            start, end, label = today, today + timedelta(days=1), "today"
    else:
        start, end, label = today, today + timedelta(days=1), "today"
    return start.isoformat() + "Z", end.isoformat() + "Z", label


def _extract_date_filter(message: str) -> str:
    """Pull a MiBuddy-style date_filter keyword out of free text."""
    for pat, key in _DATE_FILTER_PATS:
        m = pat.search(message)
        if not m:
            continue
        if key == "last_N_days":
            return f"last_{m.group(1)}_days"
        return key
    return ""


# ═══════════════════════════════════════════════════════════════════
#   MiBuddy-ported email-list formatter
# ═══════════════════════════════════════════════════════════════════

def _format_email_list(
    emails: List[Dict[str, Any]],
    *,
    search_scope: str,
) -> str:
    """Port of MiBuddy's `_format_email_list` — table layout that matches
    the screenshot from MiBuddy (# | Sender | Subject | Received | Status).

    The MiBuddy original pipes this through an LLM post-processor that
    adds a one-line preamble; we replicate the same preamble deterministically
    so agentcore doesn't need an extra LLM hop.
    """
    if not emails:
        return (
            "I couldn't find any emails matching your request.\n\n"
            "💡 **Tips:** Try a shorter keyword, search by sender name, or check "
            "if the email was just received (Graph API indexes with a short delay)."
        )

    def _short_date(iso: str) -> str:
        if not iso or len(iso) < 10:
            return iso
        months = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ]
        try:
            yyyy, mm, dd = iso[:10].split("-")
            return f"{months[int(mm) - 1]} {int(dd)}"
        except Exception:
            return iso[:10]

    email_count = len(emails)
    lines = [
        f"Here are your **{email_count}** emails from {search_scope}:",
        "",
        "| # | Sender | Subject | Received | Status |",
        "|---|---|---|---|---|",
    ]
    for i, e in enumerate(emails, 1):
        sender = (
            e.get("from", {}).get("emailAddress", {}).get("name")
            or e.get("from", {}).get("emailAddress", {}).get("address")
            or "Unknown"
        ).replace("|", "\\|")
        subj = (e.get("subject") or "No subject").strip().replace("|", "\\|")
        preview = (
            (e.get("bodyPreview") or "").strip().replace("\n", " ").replace("|", "\\|")
        )
        if len(preview) > 120:
            preview = preview[:117] + "..."
        received = _short_date(e.get("receivedDateTime", ""))
        is_read = e.get("isRead", True)
        status = "🟢 Read" if is_read else "🔵 Unread"
        subject_cell = f"**{subj}**"
        if preview:
            subject_cell += f"<br/>{preview}"
        lines.append(
            f"| {i} | {sender} | {subject_cell} | {received} | {status} |",
        )
    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════
#   MiBuddy-ported calendar formatter (Active / Cancelled sections,
#   table with | Time | Title | columns)
# ═══════════════════════════════════════════════════════════════════

def _fmt_time(dt_str: str) -> str:
    """MiBuddy: extract HH:MM from a Graph dateTime string."""
    try:
        return dt_str[11:16]
    except Exception:
        return dt_str[:16].replace("T", " ")


def _is_ooo_event(ev: Dict[str, Any]) -> bool:
    """MiBuddy: detect OOO placeholders that add noise to schedule views."""
    show_as = str(ev.get("showAs", "")).lower()
    subject = str(ev.get("subject", "")).lower()
    return (
        show_as in {"oof", "workingelsewhere"}
        or "out of office" in subject
        or "ooo" in subject
    )


def _format_calendar(
    events: List[Dict[str, Any]],
    *,
    date_label: str,
    include_ooo: bool = False,
) -> str:
    """Port of MiBuddy's `_format_calendar`."""
    if not events:
        return "No events found for the requested period."

    active: List[Dict[str, Any]] = []
    cancelled: List[Dict[str, Any]] = []
    hidden_ooo = 0
    for ev in events:
        try:
            if not include_ooo and _is_ooo_event(ev):
                hidden_ooo += 1
                continue
            is_cancelled = (
                ev.get("isCancelled", False)
                or ev.get("showAs", "").lower() == "free"
                or str(ev.get("subject", "")).lower().startswith("canceled")
                or str(ev.get("subject", "")).lower().startswith("cancelled")
            )
            entry = {
                "subject": ev.get("subject", "No title"),
                "start": _fmt_time(ev.get("start", {}).get("dateTime", "")),
                "end": _fmt_time(ev.get("end", {}).get("dateTime", "")),
                "online": ev.get("isOnlineMeeting", False),
            }
            (cancelled if is_cancelled else active).append(entry)
        except Exception:
            continue

    if not active and not cancelled:
        if hidden_ooo:
            return (
                f"No non-OOO events found for {date_label}. "
                f"({hidden_ooo} out-of-office entries were hidden.)"
            )
        return "No events found for the requested period."

    header = f"You have the following meetings scheduled for {date_label}:\n\n"
    if hidden_ooo:
        header += (
            f"(Hidden {hidden_ooo} out-of-office entries. Ask for OOO "
            f"explicitly to include them.)\n\n"
        )

    def _make_table(entries: List[Dict[str, Any]]) -> str:
        if not entries:
            return ""
        rows = ["| **Time** | **Title** |", "|----------|-----------|"]
        for e in entries:
            time_range = f"{e['start']} - {e['end']}"
            icon = " 📅" if e["online"] else ""
            title = (e["subject"] + icon).replace("|", "\\|")
            rows.append(f"| {time_range} | {title} |")
        return "\n".join(rows)

    sections: List[str] = []
    if active:
        sections.append("**Active Meetings**\n\n" + _make_table(active))
    if cancelled:
        sections.append("**Cancelled Meetings**\n\n" + _make_table(cancelled))
    return header + "\n\n".join(sections)


# ═══════════════════════════════════════════════════════════════════
#   Intent detection + dispatch
# ═══════════════════════════════════════════════════════════════════

def detect_intent(message: str) -> Optional[str]:
    """Return 'calendar' / 'email_search' / 'email_read' or None."""
    if not message or not message.strip():
        return None
    if _CALENDAR_PAT.search(message):
        return "calendar"
    if _EMAIL_SEARCH_PAT.search(message):
        return "email_search"
    if _EMAIL_READ_PAT.search(message):
        return "email_read"
    return None


def _extract_search_query(message: str) -> str:
    """Strip verbs so 'find emails about X' → 'X'."""
    cleaned = re.sub(
        r"\b(find|search|look\s+for|any|emails?|mails?|messages?|"
        r"about|for|from|the|me|my)\b",
        "",
        message,
        flags=re.IGNORECASE,
    ).strip()
    return cleaned or message


def handle_outlook_intent(user_id: str, message: str) -> Optional[OutlookIntent]:
    """Detect an outlook intent and produce a MiBuddy-formatted markdown reply.

    Returns None if the message isn't outlook-related.
    Returns an OutlookIntent with `kind="error"` when the user isn't
    connected or a Graph call fails.
    """
    intent = detect_intent(message)
    if not intent:
        return None

    access_token = outlook_token_manager.get_token(user_id)
    if not access_token:
        return OutlookIntent(
            "error",
            "Your Outlook account isn't connected, or your previous session "
            "may have expired.\n\n"
            "**To connect / reconnect:**\n"
            "1. Click the **+** button → **Outlook Connector**\n"
            "2. If it's already toggled on, switch it **off** then back **on** "
            "to re-authenticate\n"
            "3. Sign in with your Microsoft account and grant the required "
            "permissions\n\n"
            "Once connected I can read emails, check your calendar, and draft replies.",
        )

    try:
        if intent == "calendar":
            start = datetime.utcnow().isoformat() + "Z"
            end = (datetime.utcnow() + timedelta(days=7)).isoformat() + "Z"
            events = OutlookService.get_calendar_events(
                access_token, start_date=start, end_date=end, top=50,
            )
            if events is None:
                return OutlookIntent("error", "_Failed to retrieve calendar events._")
            include_ooo = bool(
                re.search(r"\b(ooo|out\s*-?of\s*-?office)\b", message, re.IGNORECASE),
            )
            return OutlookIntent(
                "calendar",
                _format_calendar(
                    events, date_label="the upcoming week", include_ooo=include_ooo,
                ),
            )

        if intent == "email_search":
            query = _extract_search_query(message)
            emails = OutlookService.search_emails(access_token, query, top=10)
            if emails is None:
                return OutlookIntent("error", "_Failed to search emails._")
            return OutlookIntent(
                "email_search",
                _format_email_list(
                    emails, search_scope=f"mailbox search for '{query}'",
                ),
            )

        # email_read — with MiBuddy-style date filter + unread filter
        date_filter = _extract_date_filter(message)
        unread_only = bool(_UNREAD_PAT.search(message))

        received_after: Optional[str] = None
        received_before: Optional[str] = None
        date_label = "your inbox"
        if date_filter:
            received_after, received_before, date_label = _parse_email_date_filter(
                date_filter,
            )

        # MiBuddy default: if no date mentioned but user said "unread",
        # don't restrict by date. Otherwise default to "today" for "show emails".
        if not date_filter and not unread_only:
            received_after, received_before, date_label = _parse_email_date_filter(
                "today",
            )

        emails = OutlookService.get_emails(
            access_token,
            top=10,
            folder="inbox",
            unread_only=unread_only,
            received_after=received_after,
            received_before=received_before,
        )
        if emails is None:
            return OutlookIntent("error", "_Failed to retrieve emails._")

        # Build MiBuddy-style scope label ("emails from today", "unread emails
        # from the last 7 days", etc.)
        if unread_only and date_label and date_label != "your inbox":
            scope = f"unread emails from {date_label}"
        elif unread_only:
            scope = "unread emails"
        elif date_label and date_label != "your inbox":
            scope = f"emails from {date_label}"
        else:
            scope = "your inbox"

        return OutlookIntent(
            "email_list",
            _format_email_list(emails, search_scope=scope),
        )
    except Exception as e:
        logger.error(f"[outlook_intent_router] {e}", exc_info=True)
        return OutlookIntent("error", f"_Outlook error: {e}_")
