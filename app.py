import json
import os
import re
import secrets
import urllib.error
import urllib.request
from base64 import urlsafe_b64decode
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, RedirectResponse
from google.auth.transport.requests import Request as GoogleRequest
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from pydantic import BaseModel, Field


ROOT = Path(__file__).resolve().parent


def load_env_file() -> None:
    env_path = ROOT / ".env"
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


load_env_file()
os.environ.setdefault("OAUTHLIB_RELAX_TOKEN_SCOPE", "1")


ARCHIVE_PATH = Path(os.getenv("GMAIL_ARCHIVE_PATH", ROOT / "gmail_archive.json"))
HTML_PATH = ROOT / "aipi hackathon - email drafter addon.html"
CLIENT_SECRET_FILE = Path(os.getenv("GOOGLE_CLIENT_SECRET_FILE", ROOT / "credentials.json"))
TOKEN_FILE = Path(os.getenv("GOOGLE_TOKEN_FILE", ROOT / "token.json"))
MAX_CONTEXT_MESSAGES = int(os.getenv("MAX_CONTEXT_MESSAGES", "5"))
MAX_BODY_CHARS = int(os.getenv("MAX_BODY_CHARS", "1200"))
GOOGLE_REDIRECT_URI = os.getenv("GOOGLE_REDIRECT_URI", "http://127.0.0.1:8000/api/google/callback")
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
MAX_LIVE_GMAIL_CANDIDATES = int(os.getenv("MAX_LIVE_GMAIL_CANDIDATES", "25"))

SEARCH_STOPWORDS = {
    "about",
    "after",
    "and",
    "because",
    "before",
    "between",
    "draft",
    "email",
    "feedback",
    "format",
    "from",
    "have",
    "into",
    "live",
    "message",
    "notes",
    "regarding",
    "scenario",
    "situation",
    "that",
    "this",
    "tone",
    "want",
    "with",
    "write",
    "your",
}

class DraftRequest(BaseModel):
    scenario: str = Field(min_length=1)
    tone: str = Field(default="Professional and measured")
    format: str = Field(default="formal email")
    situation: str = Field(min_length=1)
    thread: str = Field(default="")
    notes: str = Field(default="")
    use_gmail_archive: bool = Field(default=False)
    use_live_gmail: bool = Field(default=False)


class DraftResponse(BaseModel):
    draft: str
    context_matches: list[dict[str, Any]]


app = FastAPI(title="DraftEase API")
app.state.google_oauth_state = None
app.state.google_oauth_sessions = {}

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def normalize_base_url(url: str) -> str:
    return url.rstrip("/")


def load_archive() -> list[dict[str, Any]]:
    if not ARCHIVE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=f"Gmail archive not found at {ARCHIVE_PATH}",
        )

    try:
        payload = json.loads(ARCHIVE_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=500, detail=f"Invalid archive JSON: {exc}") from exc

    messages = payload.get("messages", [])
    if not isinstance(messages, list):
        raise HTTPException(status_code=500, detail="Archive format is invalid: 'messages' must be a list")
    return messages


def save_google_credentials(creds: Credentials) -> None:
    TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")


def load_saved_google_credentials() -> Credentials | None:
    if not TOKEN_FILE.exists():
        return None

    creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), GOOGLE_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(GoogleRequest())
        save_google_credentials(creds)
    if not creds.valid:
        return None
    return creds


def build_google_flow(state: str | None = None) -> Flow:
    if not CLIENT_SECRET_FILE.exists():
        raise HTTPException(status_code=500, detail=f"Google client secrets file not found at {CLIENT_SECRET_FILE}")

    flow = Flow.from_client_secrets_file(
        str(CLIENT_SECRET_FILE),
        scopes=GOOGLE_SCOPES,
        state=state,
        autogenerate_code_verifier=True,
    )
    flow.redirect_uri = GOOGLE_REDIRECT_URI
    return flow


def build_gmail_service(creds: Credentials):
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def decode_gmail_body(data: str) -> str:
    if not data:
        return ""
    padded = data + "=" * (-len(data) % 4)
    try:
        return urlsafe_b64decode(padded.encode("utf-8")).decode("utf-8", errors="replace")
    except Exception:
        return ""


def extract_text_from_payload(payload: dict[str, Any]) -> str:
    mime_type = payload.get("mimeType", "")
    body = payload.get("body") or {}
    parts = payload.get("parts") or []
    text_chunks: list[str] = []
    html_chunks: list[str] = []

    data = body.get("data")
    if mime_type == "text/plain" and data:
        text_chunks.append(decode_gmail_body(data))
    elif mime_type == "text/html" and data:
        html_chunks.append(decode_gmail_body(data))

    for part in parts:
        nested = extract_text_from_payload(part)
        if nested.strip():
            text_chunks.append(nested)

    if text_chunks:
        return "\n".join(chunk for chunk in text_chunks if chunk.strip())

    if html_chunks:
        cleaned_html = []
        for chunk in html_chunks:
            text = re.sub(r"<[^>]+>", " ", chunk)
            text = re.sub(r"\s+", " ", text).strip()
            if text:
                cleaned_html.append(text)
        return "\n".join(cleaned_html)

    return ""


def normalize_live_gmail_message(message: dict[str, Any]) -> dict[str, Any]:
    payload = message.get("payload") or {}
    headers_list = payload.get("headers") or []
    headers: dict[str, str | None] = {
        "from": None,
        "to": None,
        "cc": None,
        "bcc": None,
        "subject": None,
        "date": None,
    }

    for item in headers_list:
        name = (item.get("name") or "").lower()
        if name in headers:
            headers[name] = item.get("value")

    return {
        "id": message.get("id"),
        "threadId": message.get("threadId"),
        "internalDate_ms": int(message.get("internalDate", 0) or 0),
        "labelIds": message.get("labelIds") or [],
        "snippet": message.get("snippet") or "",
        "headers": headers,
        "body": {
            "text": extract_text_from_payload(payload),
            "html": "",
        },
    }


def fetch_live_gmail_messages(max_results: int = 12, query: str | None = None) -> list[dict[str, Any]]:
    creds = load_saved_google_credentials()
    if not creds:
        raise HTTPException(status_code=400, detail="Gmail is not connected. Connect Gmail first.")

    service = build_gmail_service(creds)
    list_kwargs: dict[str, Any] = {
        "userId": "me",
        "maxResults": max_results,
    }
    if query:
        list_kwargs["q"] = query
    else:
        list_kwargs["labelIds"] = ["INBOX"]

    result = service.users().messages().list(**list_kwargs).execute()
    messages = result.get("messages") or []
    normalized: list[dict[str, Any]] = []

    for item in messages:
        full_message = (
            service.users()
            .messages()
            .get(userId="me", id=item["id"], format="full")
            .execute()
        )
        normalized.append(normalize_live_gmail_message(full_message))

    return normalized


def latest_live_gmail_preview(limit: int = 10) -> list[dict[str, Any]]:
    previews: list[dict[str, Any]] = []
    for message in fetch_live_gmail_messages(max_results=limit):
        headers = message.get("headers") or {}
        previews.append(
            {
                "id": message.get("id"),
                "from": headers.get("from") or "",
                "subject": headers.get("subject") or "(no subject)",
                "date": headers.get("date") or "",
                "snippet": clip_text(message.get("snippet") or message_body(message), 220),
            }
        )
    return previews


def get_gmail_profile() -> dict[str, Any] | None:
    try:
        creds = load_saved_google_credentials()
    except Exception:
        return None
    if not creds:
        return None

    try:
        service = build_gmail_service(creds)
        return service.users().getProfile(userId="me").execute()
    except Exception:
        return None


def tokenize(text: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]{3,}", text.lower()))


def build_live_gmail_queries(request: DraftRequest) -> list[str]:
    query_source = " ".join(
        [
            request.scenario,
            request.situation,
            request.thread,
            request.notes,
        ]
    ).lower()
    tokens = [
        token
        for token in re.findall(r"[a-z0-9]{3,}", query_source)
        if token not in SEARCH_STOPWORDS
    ]

    ordered_unique: list[str] = []
    seen: set[str] = set()
    for token in tokens:
        if token in seen:
            continue
        seen.add(token)
        ordered_unique.append(token)

    phrase_matches = re.findall(r"(amazon\s+fresh|whole\s+foods|order\s+number|tracking\s+number)", query_source)

    if not ordered_unique and not phrase_matches:
        return ["newer_than:90d"]

    queries: list[str] = []
    phrase_terms = [f"\"{phrase}\"" for phrase in phrase_matches[:3]]
    search_terms = ordered_unique[:6]

    if phrase_terms:
        for phrase in phrase_terms:
            queries.append(f"newer_than:180d {phrase}")

    if len(search_terms) >= 2:
        queries.append(f"newer_than:180d {search_terms[0]} {search_terms[1]}")

    if search_terms:
        queries.append(f"newer_than:180d {search_terms[0]}")

    if "amazon" in ordered_unique:
        queries.append("newer_than:365d amazon")
    if "fresh" in ordered_unique:
        queries.append("newer_than:365d fresh")
    if "delivery" in ordered_unique:
        queries.append("newer_than:365d delivery")
    if "order" in ordered_unique:
        queries.append("newer_than:365d order")

    deduped_queries: list[str] = []
    seen: set[str] = set()
    for query in queries:
        if query in seen:
            continue
        seen.add(query)
        deduped_queries.append(query)

    return deduped_queries or ["newer_than:180d"]


def clip_text(text: str, max_chars: int = MAX_BODY_CHARS) -> str:
    compact = re.sub(r"\s+", " ", text or "").strip()
    return compact[:max_chars]


def message_body(message: dict[str, Any]) -> str:
    body = message.get("body") or {}
    text = body.get("text") or ""
    if text.strip():
        return text
    html = body.get("html") or ""
    html = re.sub(r"<[^>]+>", " ", html)
    return html


def score_message(message: dict[str, Any], query_tokens: set[str]) -> int:
    headers = message.get("headers") or {}
    searchable_parts = [
        headers.get("from") or "",
        headers.get("to") or "",
        headers.get("subject") or "",
        message.get("snippet") or "",
        clip_text(message_body(message), 4000),
    ]
    haystack = "\n".join(searchable_parts).lower()
    score = 0

    for token in query_tokens:
        if token in haystack:
            score += 1
            if token in (headers.get("subject") or "").lower():
                score += 2
            if token in (headers.get("from") or "").lower():
                score += 2
            if token in (message.get("snippet") or "").lower():
                score += 1

    if "INBOX" in (message.get("labelIds") or []):
        score += 1
    return score


def retrieve_ranked_context(
    request: DraftRequest,
    messages: list[dict[str, Any]],
    source: str,
) -> list[dict[str, Any]]:
    query = "\n".join(
        [
            request.scenario,
            request.tone,
            request.format,
            request.situation,
            request.thread,
            request.notes,
        ]
    )
    query_tokens = tokenize(query)
    if not query_tokens:
        return []

    ranked: list[tuple[int, dict[str, Any]]] = []
    for message in messages:
        score = score_message(message, query_tokens)
        if score > 0:
            ranked.append((score, message))

    ranked.sort(
        key=lambda item: (
            item[0],
            item[1].get("internalDate_ms", 0),
        ),
        reverse=True,
    )

    trimmed: list[dict[str, Any]] = []
    for score, message in ranked[:MAX_CONTEXT_MESSAGES]:
        headers = message.get("headers") or {}
        trimmed.append(
            {
                "score": score,
                "from": headers.get("from") or "",
                "to": headers.get("to") or "",
                "subject": headers.get("subject") or "",
                "date": headers.get("date") or "",
                "snippet": clip_text(message.get("snippet") or "", 240),
                "body_excerpt": clip_text(message_body(message), MAX_BODY_CHARS),
                "source": source,
            }
        )
    return trimmed


def retrieve_archive_context(request: DraftRequest) -> list[dict[str, Any]]:
    return retrieve_ranked_context(request, load_archive(), source="archive")


def retrieve_live_gmail_context(request: DraftRequest) -> list[dict[str, Any]]:
    searched_messages: list[dict[str, Any]] = []
    for query in build_live_gmail_queries(request):
        searched_messages.extend(
            fetch_live_gmail_messages(
                max_results=MAX_LIVE_GMAIL_CANDIDATES,
                query=query,
            )
        )
    recent_messages = fetch_live_gmail_messages(max_results=10)

    merged_messages_by_id: dict[str, dict[str, Any]] = {}
    for message in searched_messages + recent_messages:
        message_id = message.get("id")
        if message_id:
            merged_messages_by_id[message_id] = message

    live_messages = list(merged_messages_by_id.values())
    ranked = retrieve_ranked_context(request, live_messages, source="live_gmail")
    if ranked:
        return ranked

    # If nothing matches strongly, still use the latest few live emails so the
    # feature feels responsive after the user explicitly opts in.
    fallback: list[dict[str, Any]] = []
    for message in live_messages[:3]:
        headers = message.get("headers") or {}
        fallback.append(
            {
                "score": 0,
                "from": headers.get("from") or "",
                "to": headers.get("to") or "",
                "subject": headers.get("subject") or "",
                "date": headers.get("date") or "",
                "snippet": clip_text(message.get("snippet") or "", 240),
                "body_excerpt": clip_text(message_body(message), MAX_BODY_CHARS),
                "source": "live_gmail",
            }
        )
    return fallback


def build_system_prompt(req: DraftRequest) -> str:
    return (
        "You are DraftEase, an expert at writing emotionally intelligent, polished messages "
        "for difficult or high-stakes situations.\n\n"
        f"Write a single ready-to-send {req.format}.\n\n"
        "Rules:\n"
        f"- Match the requested tone precisely: {req.tone}\n"
        "- Be concise, natural, and specific.\n"
        "- Avoid cliches, filler, and corporate jargon.\n"
        "- If the requested format is an email, include a subject line.\n"
        "- If thread context exists, stay consistent with it.\n"
        "- Use retrieved Gmail context only when it is relevant and only when the user opted in.\n"
        "- Do not explain the reasoning. Output only the finished draft.\n"
        "- Keep it under 220 words unless the context clearly requires more."
    )


def build_user_prompt(req: DraftRequest, context_matches: list[dict[str, Any]]) -> str:
    lines = [
        f"Scenario: {req.scenario}",
        f"Tone: {req.tone}",
        f"Format: {req.format}",
        f"Situation: {req.situation}",
    ]

    if req.notes.strip():
        lines.append(f"Notes: {req.notes}")
    if req.thread.strip():
        lines.append(f"Previous thread/context:\n{req.thread}")

    if context_matches:
        lines.append("Relevant Gmail context:")
        for idx, match in enumerate(context_matches, start=1):
            lines.append(
                "\n".join(
                    [
                        f"[Context {idx}]",
                        f"Source: {match['source']}",
                        f"From: {match['from']}",
                        f"To: {match['to']}",
                        f"Subject: {match['subject']}",
                        f"Date: {match['date']}",
                        f"Snippet: {match['snippet']}",
                        f"Body excerpt: {match['body_excerpt']}",
                    ]
                )
            )

    lines.append("Write the draft now.")
    return "\n\n".join(lines)


def call_openai(system_prompt: str, user_prompt: str) -> str:
    api_key = os.getenv("OPENAI_API_KEY") or os.getenv("LITELLM_TOKEN")
    if not api_key:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY or LITELLM_TOKEN is not set")

    model = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
    base_url = normalize_base_url(
        os.getenv("OPENAI_BASE_URL")
        or os.getenv("LITELLM_BASE_URL")
        or "https://api.openai.com"
    )
    api_url = f"{base_url}/v1/responses"
    payload = {
        "model": model,
        "input": [
            {"role": "system", "content": [{"type": "input_text", "text": system_prompt}]},
            {"role": "user", "content": [{"type": "input_text", "text": user_prompt}]},
        ],
    }
    request = urllib.request.Request(
        api_url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise HTTPException(status_code=500, detail=f"OpenAI-compatible request failed: {detail}") from exc
    except urllib.error.URLError as exc:
        raise HTTPException(status_code=500, detail=f"OpenAI-compatible request failed: {exc}") from exc

    text = data.get("output_text")
    if text:
        return text.strip()

    output = data.get("output") or []
    for item in output:
        for content in item.get("content", []):
            text = content.get("text")
            if text:
                return text.strip()

    raise HTTPException(status_code=500, detail="Model response did not include draft text")


def call_anthropic(system_prompt: str, user_prompt: str) -> str:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY is not set")

    model = os.getenv("ANTHROPIC_MODEL", "claude-3-5-sonnet-latest")
    payload = {
        "model": model,
        "max_tokens": 800,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_prompt}],
    }
    request = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            data = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise HTTPException(status_code=500, detail=f"Anthropic request failed: {detail}") from exc
    except urllib.error.URLError as exc:
        raise HTTPException(status_code=500, detail=f"Anthropic request failed: {exc}") from exc

    content = data.get("content") or []
    for item in content:
        text = item.get("text")
        if text:
            return text.strip()

    raise HTTPException(status_code=500, detail="Anthropic response did not include draft text")


def generate_draft(req: DraftRequest, context_matches: list[dict[str, Any]]) -> str:
    system_prompt = build_system_prompt(req)
    user_prompt = build_user_prompt(req, context_matches)
    provider = os.getenv("LLM_PROVIDER", "openai").strip().lower()

    if provider == "anthropic":
        return call_anthropic(system_prompt, user_prompt)
    if provider == "openai":
        return call_openai(system_prompt, user_prompt)
    raise HTTPException(status_code=500, detail=f"Unsupported LLM_PROVIDER: {provider}")


@app.get("/")
def index() -> FileResponse:
    if not HTML_PATH.exists():
        raise HTTPException(status_code=404, detail="Frontend HTML file not found")
    return FileResponse(HTML_PATH)


@app.get("/api/health")
def health() -> dict[str, Any]:
    profile = get_gmail_profile()
    return {
        "ok": True,
        "archive_path": str(ARCHIVE_PATH),
        "archive_exists": ARCHIVE_PATH.exists(),
        "provider": os.getenv("LLM_PROVIDER", "openai").strip().lower(),
        "openai_base_url": normalize_base_url(
            os.getenv("OPENAI_BASE_URL")
            or os.getenv("LITELLM_BASE_URL")
            or "https://api.openai.com"
        ),
        "gmail_connected": bool(profile),
        "gmail_email": profile.get("emailAddress") if profile else None,
    }


@app.get("/api/google/status")
def google_status() -> dict[str, Any]:
    profile = get_gmail_profile()
    return {
        "connected": bool(profile),
        "email": profile.get("emailAddress") if profile else None,
    }


@app.get("/api/google/latest")
def google_latest(limit: int = 10) -> dict[str, Any]:
    if limit < 1 or limit > 25:
        raise HTTPException(status_code=400, detail="limit must be between 1 and 25")

    messages = latest_live_gmail_preview(limit=limit)
    return {"messages": messages, "count": len(messages)}


@app.get("/api/google/auth-url")
def google_auth_url() -> dict[str, str]:
    state = secrets.token_urlsafe(24)
    flow = build_google_flow(state=state)
    authorization_url, _ = flow.authorization_url(
        access_type="offline",
        prompt="consent",
        code_challenge_method="S256",
    )
    app.state.google_oauth_state = state
    app.state.google_oauth_sessions[state] = {
        "code_verifier": flow.code_verifier,
    }
    return {"auth_url": authorization_url}


@app.get("/api/google/callback")
def google_callback(state: str, code: str):
    if state != app.state.google_oauth_state:
        raise HTTPException(status_code=400, detail="Google OAuth state mismatch")

    oauth_session = app.state.google_oauth_sessions.get(state)
    if not oauth_session:
        raise HTTPException(status_code=400, detail="Google OAuth session expired. Please try connecting Gmail again.")

    flow = build_google_flow(state=state)
    flow.code_verifier = oauth_session["code_verifier"]
    flow.fetch_token(code=code)
    save_google_credentials(flow.credentials)
    app.state.google_oauth_state = None
    app.state.google_oauth_sessions.pop(state, None)
    return RedirectResponse(url="/?gmail_connected=1", status_code=302)


@app.post("/api/google/disconnect")
def google_disconnect() -> dict[str, bool]:
    if TOKEN_FILE.exists():
        TOKEN_FILE.unlink()
    app.state.google_oauth_state = None
    app.state.google_oauth_sessions = {}
    return {"ok": True}


@app.post("/api/draft", response_model=DraftResponse)
def draft_email(req: DraftRequest) -> DraftResponse:
    context_matches: list[dict[str, Any]] = []

    if req.use_gmail_archive:
        context_matches.extend(retrieve_archive_context(req))
    if req.use_live_gmail:
        context_matches.extend(retrieve_live_gmail_context(req))

    context_matches.sort(
        key=lambda item: (item.get("score", 0), item.get("date") or ""),
        reverse=True,
    )
    context_matches = context_matches[:MAX_CONTEXT_MESSAGES]

    draft = generate_draft(req, context_matches)
    return DraftResponse(draft=draft, context_matches=context_matches)
