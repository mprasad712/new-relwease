"""Outlook intent router for the orchestrator chat flow.

MiBuddy uses a full LangGraph supervisor with an Azure-OpenAI-backed
IntentClassifier + Outlook sub-agent (hundreds of lines across several
files). Here we deliver the same UX via a lightweight keyword-based
router that:

  1. Detects if a user's chat message is about email or calendar.
  2. Picks the right Graph operation (read / search / calendar).
  3. Calls the stored Outlook token directly via `OutlookService`.
  4. Formats the Graph JSON into a compact markdown block the LLM
     can cite or the UI can render as-is.

Callers should invoke `handle_outlook_intent(user_id, message)` at the
start of their chat handler; when it returns a non-None markdown string
the chat handler can either short-circuit and return it directly (no
LLM call), or prepend it to the system prompt as "tool output" so the
LLM can incorporate it into its answer.
"""
from __future__ import annotations

import logging
import re
from datetime import datetime, timedelta
from typing import List, Optional

from agentcore.services.outlook_orch.outlook_service import OutlookService
from agentcore.services.outlook_orch.token_manager import outlook_token_manager

logger = logging.getLogger(__name__)


# ── keyword → intent mapping ─────────────────────────────────────────
# Patterns are matched case-insensitively against the raw user message.

_CALENDAR_PAT = re.compile(
    r"\b("
    r"(?:my\s+)?calendar|"
    r"meeting(?:s)?|"
    r"appointment(?:s)?|"
    r"schedule|"
    r"what(?:'s| is)\s+on\s+my\s+(?:calendar|schedule)|"
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
    r"inbox"
    r")\b",
    re.IGNORECASE,
)


class OutlookIntent:
    """Container for a detected intent + its resolved data."""

    def __init__(self, kind: str, markdown: str) -> None:
        self.kind = kind          # "email_list" | "email_search" | "calendar" | "error"
        self.markdown = markdown  # rendered summary for chat

    def __repr__(self) -> str:  # pragma: no cover
        return f"<OutlookIntent {self.kind}>"


# ── Formatting helpers ──────────────────────────────────────────────

def _format_email_list(emails: List[dict], header: str) -> str:
    if not emails:
        return f"**{header}**\n\n_No emails found._"
    lines = [f"**{header}**", ""]
    for i, e in enumerate(emails, 1):
        subj = (e.get("subject") or "(no subject)").strip()
        sender = (
            e.get("from", {}).get("emailAddress", {}).get("name")
            or e.get("from", {}).get("emailAddress", {}).get("address")
            or "Unknown"
        )
        dt = e.get("receivedDateTime", "")
        preview = (e.get("bodyPreview") or "").strip().replace("\n", " ")
        if len(preview) > 160:
            preview = preview[:157] + "..."
        lines.append(f"{i}. **{subj}** — _{sender}_ · `{dt}`")
        if preview:
            lines.append(f"   {preview}")
    return "\n".join(lines)


def _format_calendar(events: List[dict], header: str) -> str:
    if not events:
        return f"**{header}**\n\n_No events._"
    lines = [f"**{header}**", "", "| Time | Title | Location |", "|---|---|---|"]
    for ev in events:
        subj = (ev.get("subject") or "(no title)").strip()
        start = ev.get("start", {}).get("dateTime", "")[:16].replace("T", " ")
        end = ev.get("end", {}).get("dateTime", "")[11:16]
        location = (
            ev.get("location", {}).get("displayName", "").strip() or "—"
        )
        lines.append(f"| {start} – {end} | {subj} | {location} |")
    return "\n".join(lines)


# ── Main dispatcher ─────────────────────────────────────────────────

def detect_intent(message: str) -> Optional[str]:
    """Return an intent name (calendar / email_search / email_read) or None."""
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
    """Strip obvious verbs so "find emails about X" → "X"."""
    cleaned = re.sub(
        r"\b(find|search|look\s+for|any|emails?|mails?|messages?|about|for|from|the|me|my)\b",
        "",
        message,
        flags=re.IGNORECASE,
    ).strip()
    return cleaned or message


def handle_outlook_intent(user_id: str, message: str) -> Optional[OutlookIntent]:
    """Detect an outlook intent and produce a chat-ready markdown block.

    Returns None if the message isn't outlook-related or if the user
    isn't connected. Never raises — any errors are captured and
    returned as an OutlookIntent of kind "error".
    """
    intent = detect_intent(message)
    if not intent:
        return None

    access_token = outlook_token_manager.get_token(user_id)
    if not access_token:
        return OutlookIntent(
            "error",
            "**Outlook isn't connected yet.** "
            "Click **Connect Outlook** in the sidebar to link your account.",
        )

    try:
        if intent == "calendar":
            # Default: next 7 days
            events = OutlookService.get_calendar_events(
                access_token,
                start_date=datetime.utcnow().isoformat() + "Z",
                end_date=(datetime.utcnow() + timedelta(days=7)).isoformat() + "Z",
                top=20,
            )
            if events is None:
                return OutlookIntent("error", "_Failed to retrieve calendar events._")
            return OutlookIntent("calendar", _format_calendar(events, "Upcoming events"))

        if intent == "email_search":
            query = _extract_search_query(message)
            emails = OutlookService.search_emails(access_token, query, top=10)
            if emails is None:
                return OutlookIntent("error", "_Failed to search emails._")
            return OutlookIntent(
                "email_search",
                _format_email_list(emails, f"Results for: _{query}_"),
            )

        # email_read
        emails = OutlookService.get_emails(
            access_token, top=10, folder="inbox",
        )
        if emails is None:
            return OutlookIntent("error", "_Failed to retrieve emails._")
        return OutlookIntent("email_list", _format_email_list(emails, "Recent emails"))

    except Exception as e:
        logger.error(f"[outlook_intent_router] {e}", exc_info=True)
        return OutlookIntent("error", f"_Outlook error: {e}_")
