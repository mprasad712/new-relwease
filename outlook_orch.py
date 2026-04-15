"""Orchestrator Outlook integration — port of MiBuddy's `/outlook/...` routes.

Kept separate from agentcore's existing `outlook_chat` router (different
use-case, must not be touched). Mounted under `/api/outlook-orch`.

Flow:
  1. Frontend opens popup → GET /api/outlook-orch/auth/login
       → 302 to login.microsoftonline.com with PKCE challenge
  2. Microsoft redirects back to /api/outlook-orch/auth/callback
       → exchange code (+ code_verifier) for access_token
       → encrypt + store token in OutlookTokenManager keyed by user_id
       → set `outlook_session` HTTP-only cookie
       → postMessage success to opener window, then close popup
  3. GET /api/outlook-orch/status → is this user connected?
  4. POST /api/outlook-orch/{get_emails|search_emails|get_calendar|send_email}
     — all take `access_token` in body like MiBuddy, but server can also
     resolve it from the stored token if the frontend passes nothing.
"""
from __future__ import annotations

import base64
import hashlib
import html
import logging
import secrets
import threading
import time
from typing import Any, Dict, Optional
from urllib.parse import quote_plus, urlparse
from uuid import uuid4

import requests
from fastapi import APIRouter, HTTPException, Request, status
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from pydantic import BaseModel

from agentcore.api.utils import CurrentActiveUser, DbSession
from agentcore.services.auth.utils import get_current_user_by_jwt
from agentcore.services.outlook_orch.intent_router import handle_outlook_intent
from agentcore.services.outlook_orch.outlook_service import OutlookService
from agentcore.services.outlook_orch.token_manager import outlook_token_manager


async def _resolve_current_user_from_request(request: Request, db):
    """Resolve the current user from cookie, Authorization header, OR query
    param. The FastAPI-standard `CurrentActiveUser` dependency only reads the
    Authorization header, which popup windows (opened via `window.open()`) do
    not send. This helper mirrors the pattern used by `teams.py`'s popup
    OAuth login so cookies work for browser-navigation flows.
    """
    auth_header = request.headers.get("Authorization", "")
    bearer_token = (
        auth_header.split(" ", 1)[1] if auth_header.startswith("Bearer ") else None
    )
    cookie_token = request.cookies.get("access_token_lf")
    query_token = request.query_params.get("token")
    resolved_token = cookie_token or bearer_token or query_token
    if not resolved_token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Missing authentication token",
        )
    return await get_current_user_by_jwt(resolved_token, db)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/outlook-orch", tags=["Outlook (Orchestrator)"])

# ── OAuth state storage (in-process, like MiBuddy) ────────────────────
_PKCE_STATE_TTL = 600  # seconds
_pkce_states: Dict[str, Dict[str, Any]] = {}
_pkce_lock = threading.Lock()

# session_id → user_id (for the outlook_session cookie)
_outlook_cookie_sessions: Dict[str, str] = {}
_session_lock = threading.Lock()

_OUTLOOK_SCOPES = " ".join([
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.Read.Shared",
    "https://graph.microsoft.com/Mail.ReadBasic",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Calendars.Read",
    "https://graph.microsoft.com/Calendars.Read.Shared",
    "https://graph.microsoft.com/Calendars.ReadBasic",
    "offline_access",
])

_AUTH_SUCCESS_HTML = """<!DOCTYPE html>
<html><head><title>Connecting...</title></head>
<body>
<script>
  try { if (window.opener) { window.opener.postMessage({type:'OUTLOOK_AUTH_SUCCESS'}, '*'); } } catch(e) {}
  window.close();
</script>
<p>Outlook connected. You may close this window.</p>
</body></html>"""

_AUTH_ERROR_HTML = """<!DOCTYPE html>
<html><head><title>Auth Error</title></head>
<body>
<script>
  try { if (window.opener) { window.opener.postMessage({type:'OUTLOOK_AUTH_ERROR',error:'{error}'}, '*'); } } catch(e) {}
  window.close();
</script>
<p>Authentication failed: {error}. You may close this window.</p>
</body></html>"""


def _build_redirect_uri(request: Request) -> str:
    """Build the redirect URI deterministically from the incoming request.

    An explicit `OUTLOOK_ORCH_REDIRECT_URI` env var takes precedence so prod
    deployments behind a reverse proxy don't depend on request.url.
    """
    import os
    override = os.environ.get("OUTLOOK_ORCH_REDIRECT_URI", "").strip().strip("'\"")
    if override:
        return override
    base = str(request.base_url).rstrip("/")
    return f"{base}/api/outlook-orch/auth/callback"


def _is_request_secure(request: Request) -> bool:
    if request.url.scheme == "https":
        return True
    fwd = request.headers.get("x-forwarded-proto", "")
    return "https" in fwd.lower()


def _html_escape(s: str) -> str:
    return html.escape(s, quote=True)


# ══════════════════════════════════════════════════════════════════════
#   OAUTH: LOGIN + CALLBACK
# ══════════════════════════════════════════════════════════════════════

@router.get("/auth/login")
async def outlook_auth_login(
    request: Request,
    session: DbSession,
) -> RedirectResponse:
    """Kick off Authorization Code + PKCE flow.

    Opened via `window.open()`, so we resolve the user from the
    `access_token_lf` cookie (popups don't send Authorization headers).
    """
    current_user = await _resolve_current_user_from_request(request, session)

    # Purge stale PKCE states
    cutoff = time.time() - _PKCE_STATE_TTL
    with _pkce_lock:
        stale = [k for k, v in _pkce_states.items() if v.get("created_at", 0) < cutoff]
        for k in stale:
            _pkce_states.pop(k, None)

    user_id = str(current_user.id)

    code_verifier = secrets.token_urlsafe(64)
    code_challenge = base64.urlsafe_b64encode(
        hashlib.sha256(code_verifier.encode("ascii")).digest(),
    ).rstrip(b"=").decode("ascii")

    state = secrets.token_urlsafe(32)
    with _pkce_lock:
        _pkce_states[state] = {
            "code_verifier": code_verifier,
            "user_id": user_id,
            "created_at": time.time(),
        }

    redirect_uri = _build_redirect_uri(request)
    logger.info(f"Outlook OAuth login: redirect_uri={redirect_uri}")

    # Tenant comes from the same resolver the service uses
    from agentcore.services.outlook_orch.outlook_service import _get_credentials
    tenant, client_id, _ = _get_credentials()

    auth_url = (
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"
        f"?client_id={client_id}"
        f"&response_type=code"
        f"&redirect_uri={quote_plus(redirect_uri)}"
        f"&scope={quote_plus(_OUTLOOK_SCOPES)}"
        f"&state={state}"
        f"&code_challenge={code_challenge}"
        f"&code_challenge_method=S256"
        f"&prompt=select_account"
    )
    return RedirectResponse(url=auth_url, status_code=302)


@router.get("/auth/callback")
async def outlook_auth_callback(
    request: Request,
    code: Optional[str] = None,
    state: Optional[str] = None,
    error: Optional[str] = None,
    error_description: Optional[str] = None,
) -> HTMLResponse:
    """Complete PKCE flow and set the outlook_session cookie."""
    def err(reason: str) -> HTMLResponse:
        return HTMLResponse(content=_AUTH_ERROR_HTML.format(error=_html_escape(reason)))

    if error:
        logger.warning(f"Outlook OAuth error from MS: {error} - {error_description}")
        return err(error)
    if not code or not state:
        return err("missing_code_or_state")

    with _pkce_lock:
        state_data = _pkce_states.pop(state, None)
    if not state_data:
        return err("invalid_state")
    if time.time() - state_data.get("created_at", 0) > _PKCE_STATE_TTL:
        return err("state_expired")

    code_verifier = state_data["code_verifier"]
    user_id = state_data.get("user_id")

    redirect_uri = _build_redirect_uri(request)

    token_resp = OutlookService.exchange_code_for_token(
        auth_code=code, redirect_uri=redirect_uri, code_verifier=code_verifier,
    )
    if not token_resp:
        return err("token_exchange_failed")

    access_token = token_resp.get("access_token")
    expires_in = int(token_resp.get("expires_in", 3600))
    if not access_token:
        return err("no_access_token_in_response")

    # If we somehow lost the user_id, recover from Graph /me
    if not user_id:
        user_info = OutlookService.validate_access_token(access_token)
        if user_info:
            user_id = user_info.get("id") or user_info.get("userPrincipalName")
    if not user_id:
        return err("cannot_identify_user")

    outlook_token_manager.store_token(user_id, access_token, expires_in)

    session_id = str(uuid4())
    with _session_lock:
        _outlook_cookie_sessions[session_id] = user_id

    response = HTMLResponse(content=_AUTH_SUCCESS_HTML)
    response.set_cookie(
        key="outlook_session",
        value=session_id,
        httponly=True,
        secure=_is_request_secure(request),
        samesite="lax",
        max_age=expires_in,
        path="/",
    )
    return response


# ══════════════════════════════════════════════════════════════════════
#   STATUS / DISCONNECT
# ══════════════════════════════════════════════════════════════════════

@router.get("/status")
async def outlook_status(
    request: Request, current_user: CurrentActiveUser,
) -> JSONResponse:
    user_id = str(current_user.id)
    session_id = request.cookies.get("outlook_session")
    with _session_lock:
        cookie_user = _outlook_cookie_sessions.get(session_id) if session_id else None
    token_connected = outlook_token_manager.is_connected(user_id)
    is_connected = bool(session_id and cookie_user == user_id and token_connected)
    return JSONResponse(content={"connected": is_connected}, status_code=200)


@router.post("/disconnect")
async def outlook_disconnect(
    request: Request, current_user: CurrentActiveUser,
) -> JSONResponse:
    user_id = str(current_user.id)
    outlook_token_manager.delete_token(user_id)
    session_id = request.cookies.get("outlook_session")
    if session_id:
        with _session_lock:
            _outlook_cookie_sessions.pop(session_id, None)
    response = JSONResponse(
        content={"message": "Outlook disconnected successfully"}, status_code=200,
    )
    response.delete_cookie("outlook_session", path="/")
    return response


# ══════════════════════════════════════════════════════════════════════
#   DATA ENDPOINTS  (access_token may come from body OR server-stored)
# ══════════════════════════════════════════════════════════════════════

def _resolve_token(
    body_token: Optional[str], current_user
) -> Optional[str]:
    """Prefer token from request body (MiBuddy-style), else use server store."""
    if body_token:
        return body_token
    return outlook_token_manager.get_token(str(current_user.id))


class ValidateTokenReq(BaseModel):
    access_token: Optional[str] = None


class GetEmailsReq(BaseModel):
    access_token: Optional[str] = None
    top: int = 10
    skip: int = 0
    folder: str = "inbox"
    search: Optional[str] = None
    unread_only: bool = False
    received_after: Optional[str] = None
    received_before: Optional[str] = None


class SearchEmailsReq(BaseModel):
    access_token: Optional[str] = None
    query: str
    top: int = 10


class GetCalendarReq(BaseModel):
    access_token: Optional[str] = None
    start_date: Optional[str] = None
    end_date: Optional[str] = None
    top: int = 10


class SendEmailReq(BaseModel):
    access_token: Optional[str] = None
    to: list[str]
    subject: str
    body: str = ""
    body_type: str = "HTML"


@router.post("/validate_token")
async def validate_token(req: ValidateTokenReq, current_user: CurrentActiveUser):
    access_token = _resolve_token(req.access_token, current_user)
    if not access_token:
        return JSONResponse(
            content={"valid": False, "error": "No access token available"},
            status_code=401,
        )
    user_info = OutlookService.validate_access_token(access_token)
    if user_info:
        return {"valid": True, "user": user_info}
    return JSONResponse(
        content={"valid": False, "error": "Invalid or expired token"},
        status_code=401,
    )


@router.post("/get_emails")
async def get_emails(req: GetEmailsReq, current_user: CurrentActiveUser):
    access_token = _resolve_token(req.access_token, current_user)
    if not access_token:
        return JSONResponse(content={"error": "Not connected to Outlook"}, status_code=401)
    emails = OutlookService.get_emails(
        access_token,
        top=req.top,
        skip=req.skip,
        folder=req.folder,
        search=req.search,
        unread_only=req.unread_only,
        received_after=req.received_after,
        received_before=req.received_before,
    )
    if emails is None:
        return JSONResponse(content={"error": "Failed to retrieve emails"}, status_code=500)
    return {"success": True, "emails": emails, "count": len(emails)}


@router.post("/search_emails")
async def search_emails(req: SearchEmailsReq, current_user: CurrentActiveUser):
    access_token = _resolve_token(req.access_token, current_user)
    if not access_token:
        return JSONResponse(content={"error": "Not connected to Outlook"}, status_code=401)
    if not req.query:
        return JSONResponse(content={"error": "query is required"}, status_code=400)
    emails = OutlookService.search_emails(access_token, req.query, req.top)
    if emails is None:
        return JSONResponse(content={"error": "Failed to search emails"}, status_code=500)
    return {"success": True, "emails": emails, "count": len(emails)}


@router.post("/get_calendar")
async def get_calendar(req: GetCalendarReq, current_user: CurrentActiveUser):
    access_token = _resolve_token(req.access_token, current_user)
    if not access_token:
        return JSONResponse(content={"error": "Not connected to Outlook"}, status_code=401)
    events = OutlookService.get_calendar_events(
        access_token, req.start_date, req.end_date, req.top,
    )
    if events is None:
        return JSONResponse(
            content={"error": "Failed to retrieve calendar events"}, status_code=500,
        )
    return {"success": True, "events": events, "count": len(events)}


class IntentCheckReq(BaseModel):
    message: str


@router.post("/intent")
async def outlook_intent_check(
    req: IntentCheckReq, current_user: CurrentActiveUser,
):
    """Frontend precheck — does this message look like an Outlook query?

    If yes and the user is connected, returns the Graph-backed markdown
    reply the UI can render as-is. If no, returns ``{matched: False}`` so
    the caller continues with the normal LLM chat flow.
    """
    result = handle_outlook_intent(str(current_user.id), req.message)
    if result is None:
        return {"matched": False}
    return {"matched": True, "kind": result.kind, "markdown": result.markdown}


@router.post("/send_email")
async def send_email(req: SendEmailReq, current_user: CurrentActiveUser):
    access_token = _resolve_token(req.access_token, current_user)
    if not access_token:
        return JSONResponse(content={"error": "Not connected to Outlook"}, status_code=401)
    if not req.to or not req.subject:
        return JSONResponse(
            content={"error": "'to' and 'subject' are required"}, status_code=400,
        )
    ok = OutlookService.send_email(access_token, req.to, req.subject, req.body, req.body_type)
    if ok:
        return {"success": True, "message": "Email sent successfully"}
    return JSONResponse(content={"error": "Failed to send email"}, status_code=500)
