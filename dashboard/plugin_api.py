"""Microsoft Teams inbox API for Hermes Desktop.

Mounted at /api/plugins/hermes-teams-inbox/.

The Azure CLI Graph token is a credential for Microsoft Graph, not
authorization for this HTTP API. Every Graph route requires a per-install
secret (created on first use under the Hermes home, mode 0600). The desktop
client reads that file and sends it. The standalone server binds to
127.0.0.1 unless explicitly opted out. Only the Graph reads the pane uses
are forwarded, and team/channel ids are encoded with ``safe=''`` so a slash
cannot become another path segment.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import secrets
import subprocess
import threading
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Tuple

try:
    from fastapi import APIRouter, Request
    from fastapi.responses import JSONResponse
except Exception:  # pragma: no cover
    class APIRouter:  # type: ignore
        def get(self, *_a, **_k):
            return lambda fn: fn

        def api_route(self, *_a, **_k):
            return lambda fn: fn

    class Request:  # type: ignore
        pass

    JSONResponse = None  # type: ignore

try:
    from hermes_constants import get_hermes_home
except Exception:  # pragma: no cover
    def get_hermes_home() -> Path:  # type: ignore[misc]
        val = (os.environ.get("HERMES_HOME") or "").strip()
        return Path(val) if val else Path.home() / ".hermes"

router = APIRouter()

GRAPH = "https://graph.microsoft.com/v1.0"
PLUGIN_ID = "hermes-teams-inbox"
SECRET_HEADER = "x-hermes-teams-proxy-secret"
SECRET_QUERY = "proxy_secret"
DEFAULT_BIND_HOST = "127.0.0.1"
DEFAULT_PORT = 8765
LOOPBACK_HOSTS = frozenset({"127.0.0.1", "::1", "localhost", "::ffff:127.0.0.1"})

_AZ_FALLBACKS = (
    r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    r"C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    "az.cmd",
    "az",
)

_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:@-]{0,255}$")
_TEAMS_WEB_HOSTS = (
    "teams.microsoft.com",
    "teams.microsoft.us",
    "gov.teams.microsoft.us",
    "teams.live.com",
    "teams.cloud.microsoft",
)

_ME_SELECT = "displayName,mail,userPrincipalName,id"
_TEAMS_SELECT = "id,displayName,description,webUrl"
_CHANNELS_SELECT = "id,displayName,webUrl,membershipType"


class GraphPathError(ValueError):
    """Raised when a Graph target is not on the allowlist or an id is unsafe."""


def proxy_secret_path() -> Path:
    return get_hermes_home() / "plugins" / PLUGIN_ID / "proxy.secret"


def _chmod(path: Path, mode: int) -> None:
    try:
        os.chmod(path, mode)
    except OSError:
        pass


def _read_secret_file(path: Path) -> str:
    try:
        raw = path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return ""
    return raw.strip()


def ensure_proxy_secret() -> str:
    """Return the per-install secret, creating it on first use (mode 0600)."""
    path = proxy_secret_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    _chmod(path.parent, 0o700)
    existing = _read_secret_file(path)
    if existing:
        _chmod(path, 0o600)
        return existing
    if path.is_symlink() or path.is_file():
        # Empty file or a symlink. Replace the directory entry; do not follow it.
        path.unlink()
    token = secrets.token_urlsafe(32)
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    if hasattr(os, "O_CLOEXEC"):
        flags |= os.O_CLOEXEC
    try:
        fd = os.open(path, flags, 0o600)
    except FileExistsError:
        existing = _read_secret_file(path)
        if existing:
            _chmod(path, 0o600)
            return existing
        raise
    try:
        os.write(fd, (token + "\n").encode("utf-8"))
    finally:
        os.close(fd)
    _chmod(path, 0o600)
    return token


def _secret_matches(provided: str, expected: str) -> bool:
    if not provided or not expected:
        return False
    return hmac.compare_digest(
        hashlib.sha256(provided.encode("utf-8")).digest(),
        hashlib.sha256(expected.encode("utf-8")).digest(),
    )


def remote_clients_allowed() -> bool:
    return (os.environ.get("HERMES_TEAMS_PROXY_ALLOW_REMOTE") or "").strip() == "1"


def resolve_bind_host(raw: Optional[str] = None) -> str:
    """Standalone proxy bind address. Defaults to loopback and refuses other hosts."""
    host = (raw if raw is not None else (os.environ.get("HERMES_TEAMS_PROXY_HOST") or "")).strip()
    if not host:
        host = DEFAULT_BIND_HOST
    if host == "localhost":
        host = DEFAULT_BIND_HOST
    if host not in {DEFAULT_BIND_HOST, "::1"} and not remote_clients_allowed():
        raise RuntimeError("proxy_bind_must_be_loopback")
    return host


def _client_is_loopback(client_host: str) -> bool:
    return (client_host or "").strip().lower() in LOOPBACK_HOSTS


def _az_cmd() -> Optional[str]:
    """Resolve the Azure CLI binary.

    ``AZ_CMD`` is honored only when it is an absolute path to a regular file.
    A relative value is ignored so the environment cannot point this process
    at an arbitrary executable name on ``PATH``.
    """
    override = (os.environ.get("AZ_CMD") or "").strip()
    if override:
        candidate = Path(override)
        if candidate.is_absolute() and candidate.is_file():
            return str(candidate)
    for cand in _AZ_FALLBACKS:
        if not cand:
            continue
        path = Path(cand)
        if path.is_file():
            return str(path)
        if cand in ("az", "az.cmd"):
            return cand
    return None


def graph_token() -> Tuple[Optional[str], Optional[str]]:
    """Return (access_token, error). Never log the token."""
    az = _az_cmd()
    if not az:
        return None, "azure_cli_missing"
    try:
        proc = subprocess.run(
            [az, "account", "get-access-token", "--resource", "https://graph.microsoft.com", "-o", "json"],
            capture_output=True,
            text=True,
            timeout=45,
            shell=False,
        )
    except FileNotFoundError:
        return None, "azure_cli_missing"
    except subprocess.TimeoutExpired:
        return None, "azure_cli_timeout"
    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()[:240]
        return None, err or "azure_cli_failed"
    try:
        payload = json.loads(proc.stdout or "{}")
    except json.JSONDecodeError:
        return None, "azure_cli_bad_json"
    token = payload.get("accessToken")
    if not isinstance(token, str) or not token:
        return None, "azure_cli_no_token"
    return token, None


def _validate_graph_id(value: str, *, field: str, min_len: int) -> str:
    if not isinstance(value, str):
        raise GraphPathError(f"invalid_{field}")
    text = value.strip()
    if len(text) < min_len or len(text) > 256:
        raise GraphPathError(f"invalid_{field}")
    if text in {".", ".."} or ".." in text:
        raise GraphPathError(f"invalid_{field}")
    if any(ch in text for ch in "/\\?#%\x00"):
        raise GraphPathError(f"invalid_{field}")
    if not _ID_RE.fullmatch(text):
        raise GraphPathError(f"invalid_{field}")
    return text


def graph_target_path(
    kind: str,
    team_id: Optional[str] = None,
    channel_id: Optional[str] = None,
) -> str:
    """Build a Graph path. Identifiers are encoded with ``safe=''``.

    A slash inside ``team_id`` or ``channel_id`` stays inside one segment
    (``%2F``). It is never a separator, so the target cannot move to another
    Graph resource.
    """
    if kind == "me":
        return "/me"
    if kind == "joined_teams":
        return "/me/joinedTeams"
    if kind == "channels":
        segment = urllib.parse.quote(team_id or "", safe="")
        return f"/teams/{segment}/channels"
    if kind == "messages":
        team = urllib.parse.quote(team_id or "", safe="")
        channel = urllib.parse.quote(channel_id or "", safe="")
        return f"/teams/{team}/channels/{channel}/messages"
    raise GraphPathError("disallowed_path")


def _expected_query(kind: str) -> Dict[str, str]:
    if kind == "me":
        return {"$select": _ME_SELECT}
    if kind == "joined_teams":
        return {"$select": _TEAMS_SELECT}
    if kind == "channels":
        return {"$select": _CHANNELS_SELECT}
    if kind == "messages":
        return {"$top": "15"}
    raise GraphPathError("disallowed_path")


def approved_graph_request(
    kind: str,
    team_id: Optional[str] = None,
    channel_id: Optional[str] = None,
) -> Tuple[str, Dict[str, str]]:
    """Return the only Graph request the app is allowed to make for ``kind``."""
    if kind == "channels":
        team_id = _validate_graph_id(team_id or "", field="team_id", min_len=8)
    elif kind == "messages":
        team_id = _validate_graph_id(team_id or "", field="team_id", min_len=8)
        channel_id = _validate_graph_id(channel_id or "", field="channel_id", min_len=4)
    elif kind not in {"me", "joined_teams"}:
        raise GraphPathError("disallowed_path")
    path = graph_target_path(kind, team_id=team_id, channel_id=channel_id)
    _assert_path_shape(kind, path)
    return path, _expected_query(kind)


def _assert_path_shape(kind: str, path: str) -> None:
    parts = path.split("/")
    if kind == "me" and parts == ["", "me"]:
        return
    if kind == "joined_teams" and parts == ["", "me", "joinedTeams"]:
        return
    if kind == "channels" and len(parts) == 4 and parts[0] == "" and parts[1] == "teams" and parts[3] == "channels":
        if parts[2] and "/" not in parts[2]:
            return
    if (
        kind == "messages"
        and len(parts) == 6
        and parts[0] == ""
        and parts[1] == "teams"
        and parts[3] == "channels"
        and parts[5] == "messages"
        and parts[2]
        and parts[4]
        and "/" not in parts[2]
        and "/" not in parts[4]
    ):
        return
    raise GraphPathError("disallowed_path")


def _classify_graph_path(path: str) -> Tuple[str, Dict[str, str]]:
    """Map a concrete Graph path back to an allowlisted request, or reject it."""
    if path == "/me":
        return approved_graph_request("me")
    if path == "/me/joinedTeams":
        return approved_graph_request("joined_teams")
    parts = path.split("/")
    if len(parts) == 4 and parts[1] == "teams" and parts[3] == "channels":
        team_id = urllib.parse.unquote(parts[2])
        approved = approved_graph_request("channels", team_id=team_id)
        if approved[0] != path:
            raise GraphPathError("disallowed_path")
        return approved
    if len(parts) == 6 and parts[1] == "teams" and parts[3] == "channels" and parts[5] == "messages":
        team_id = urllib.parse.unquote(parts[2])
        channel_id = urllib.parse.unquote(parts[4])
        approved = approved_graph_request("messages", team_id=team_id, channel_id=channel_id)
        if approved[0] != path:
            raise GraphPathError("disallowed_path")
        return approved
    raise GraphPathError("disallowed_path")


def graph_get(path: str, token: str, query: Optional[Dict[str, str]] = None) -> Tuple[int, Dict[str, Any]]:
    """GET one allowlisted Graph path. Anything else is rejected with no network I/O."""
    try:
        approved_path, approved_query = _classify_graph_path(path)
    except GraphPathError as exc:
        code = "invalid_team_id" if str(exc) == "invalid_team_id" else (
            "invalid_channel_id" if str(exc) == "invalid_channel_id" else "disallowed_path"
        )
        status = 400 if code.startswith("invalid_") else 403
        return status, {"ok": False, "error": code}
    if dict(query or {}) != approved_query:
        return 403, {"ok": False, "error": "disallowed_path"}
    url = GRAPH + approved_path
    url += "?" + urllib.parse.urlencode(approved_query)
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            body = json.loads(resp.read().decode("utf-8") or "{}")
            return resp.status, body if isinstance(body, dict) else {"value": body}
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            parsed = {"error": {"message": raw[:400]}}
        return e.code, parsed if isinstance(parsed, dict) else {"error": {"message": raw[:400]}}
    except Exception as e:  # noqa: BLE001
        return 599, {"error": {"message": str(e)}}


def missing_scope(payload: Dict[str, Any]) -> bool:
    msg = str(((payload.get("error") or {}) if isinstance(payload.get("error"), dict) else {}).get("message") or "")
    return "Missing scope permissions" in msg or "Insufficient privileges" in msg


def safe_web_url(url: Any) -> str:
    """Keep https Teams links. Drop every other scheme and host."""
    if not isinstance(url, str):
        return ""
    text = url.strip()
    if not text:
        return ""
    parsed = urllib.parse.urlparse(text)
    if parsed.scheme != "https" or parsed.username or parsed.password:
        return ""
    host = (parsed.hostname or "").lower().rstrip(".")
    if not host:
        return ""
    allowed = any(host == name or host.endswith("." + name) for name in _TEAMS_WEB_HOSTS)
    if not allowed:
        return ""
    return text


def teams_platform_enabled() -> bool:
    cfg = get_hermes_home() / "config.yaml"
    try:
        text = cfg.read_text(encoding="utf-8")
    except OSError:
        return False
    # Cheap structural read: platforms.teams.enabled without a yaml dep requirement.
    in_platforms = False
    in_teams = False
    for raw in text.splitlines():
        line = raw.split("#", 1)[0]
        if line.strip() == "platforms:":
            in_platforms = True
            in_teams = False
            continue
        if in_platforms and line.startswith("  teams:"):
            in_teams = True
            continue
        if in_platforms and in_teams:
            if line.startswith("    enabled:"):
                return "true" in line.lower()
            if line.startswith("  ") and not line.startswith("    "):
                in_teams = False
        if in_platforms and line and not line.startswith(" ") and not line.startswith("\t"):
            in_platforms = False
    return False


def _graph_error_message(payload: Mapping[str, Any], code: int) -> str:
    err = payload.get("error")
    if isinstance(err, dict):
        message = err.get("message")
        if isinstance(message, str) and message:
            return message
    return f"graph_{code}"


def _gate(secret: str, client_host: str) -> Optional[Tuple[int, Dict[str, Any]]]:
    if not remote_clients_allowed() and not _client_is_loopback(client_host):
        return 403, {"ok": False, "error": "loopback_only"}
    # Create the secret on a local request so the desktop client can read the
    # file after a rejected call. The secret is never returned here.
    try:
        expected = ensure_proxy_secret()
    except OSError:
        return 401, {"ok": False, "error": "unauthorized"}
    if not _secret_matches(secret, expected):
        return 401, {"ok": False, "error": "unauthorized"}
    return None


def _no_token(err: Optional[str], extra: Dict[str, Any]) -> Tuple[int, Dict[str, Any]]:
    body: Dict[str, Any] = {"ok": False, "error": err or "no_token"}
    body.update(extra)
    return 200, body


def handle_status() -> Tuple[int, Dict[str, Any]]:
    token, err = graph_token()
    if err or not token:
        return _no_token(
            err,
            {
                "source": "azure-cli",
                "hint": "Run `az login` (Azure CLI is already the Graph credential source).",
                "teams_gateway_enabled": teams_platform_enabled(),
            },
        )
    code, me = graph_get("/me", token, _expected_query("me"))
    if code != 200:
        return 200, {
            "ok": False,
            "source": "azure-cli",
            "error": _graph_error_message(me, code),
            "teams_gateway_enabled": teams_platform_enabled(),
        }
    return 200, {
        "ok": True,
        "source": "azure-cli",
        "me": {
            "id": me.get("id"),
            "displayName": me.get("displayName"),
            "mail": me.get("mail"),
            "userPrincipalName": me.get("userPrincipalName"),
        },
        "teams_gateway_enabled": teams_platform_enabled(),
    }


def handle_teams() -> Tuple[int, Dict[str, Any]]:
    token, err = graph_token()
    if err or not token:
        return _no_token(err, {"teams": []})
    code, data = graph_get("/me/joinedTeams", token, _expected_query("joined_teams"))
    if code != 200:
        return 200, {
            "ok": False,
            "error": _graph_error_message(data, code),
            "needs_scope": missing_scope(data),
            "teams": [],
        }
    teams: List[Dict[str, Any]] = []
    for row in data.get("value") or []:
        if not isinstance(row, dict):
            continue
        teams.append(
            {
                "id": row.get("id"),
                "displayName": row.get("displayName") or "(unnamed team)",
                "description": row.get("description") or "",
                "webUrl": safe_web_url(row.get("webUrl")),
            }
        )
    return 200, {"ok": True, "teams": teams}


def handle_channels(team_id: str) -> Tuple[int, Dict[str, Any]]:
    try:
        path, query = approved_graph_request("channels", team_id=team_id)
    except GraphPathError as exc:
        return 400, {"ok": False, "error": str(exc), "channels": []}
    token, err = graph_token()
    if err or not token:
        return _no_token(err, {"channels": []})
    code, data = graph_get(path, token, query)
    if code != 200:
        return 200, {
            "ok": False,
            "error": _graph_error_message(data, code),
            "needs_scope": missing_scope(data),
            "channels": [],
        }
    channels = []
    for row in data.get("value") or []:
        if not isinstance(row, dict):
            continue
        channels.append(
            {
                "id": row.get("id"),
                "displayName": row.get("displayName") or "(channel)",
                "webUrl": safe_web_url(row.get("webUrl")),
                "membershipType": row.get("membershipType") or "",
            }
        )
    return 200, {"ok": True, "channels": channels}


def handle_messages(team_id: str, channel_id: str) -> Tuple[int, Dict[str, Any]]:
    try:
        path, query = approved_graph_request("messages", team_id=team_id, channel_id=channel_id)
    except GraphPathError as exc:
        return 400, {"ok": False, "error": str(exc), "messages": []}
    token, err = graph_token()
    if err or not token:
        return _no_token(err, {"messages": []})
    code, data = graph_get(path, token, query)
    if code != 200:
        return 200, {
            "ok": False,
            "error": _graph_error_message(data, code),
            "needs_scope": missing_scope(data),
            "messages": [],
            "scope": "ChannelMessage.Read.All",
        }
    messages = []
    for row in data.get("value") or []:
        if not isinstance(row, dict):
            continue
        from_user = ((row.get("from") or {}).get("user") or {})
        body = row.get("body") or {}
        messages.append(
            {
                "id": row.get("id"),
                "createdDateTime": row.get("createdDateTime"),
                "from": from_user.get("displayName") or "unknown",
                "preview": str(body.get("content") or "")[:280],
            }
        )
    return 200, {"ok": True, "messages": messages}


def dispatch(
    method: str,
    path: str,
    query: Mapping[str, str],
    secret: str,
    client_host: str,
) -> Tuple[int, Dict[str, Any]]:
    """Shared router for the FastAPI mount and the standalone localhost server."""
    gate = _gate(secret, client_host)
    if gate:
        return gate
    if method.upper() not in {"GET", "POST"}:
        return 405, {"ok": False, "error": "method_not_allowed"}
    normalized = path.split("?", 1)[0]
    if len(normalized) > 1:
        normalized = normalized.rstrip("/")
    if not normalized.startswith("/"):
        normalized = "/" + normalized
    if normalized == "/status":
        return handle_status()
    if normalized == "/teams":
        return handle_teams()
    if normalized == "/channels":
        return handle_channels(query.get("team_id") or "")
    if normalized == "/messages":
        return handle_messages(query.get("team_id") or "", query.get("channel_id") or "")
    return 403, {"ok": False, "error": "disallowed_path"}


def _request_secret(headers: Mapping[str, str]) -> str:
    """Read the secret from a header. Never from the query string.

    Uvicorn access logs and the desktop API failure log record the request
    line, including the query. A query parameter would persist the secret.
    """
    header = ""
    for key, value in headers.items():
        if key.lower() == SECRET_HEADER and str(value).strip():
            header = str(value).strip()
            break
    if header:
        return header
    auth = ""
    for key, value in headers.items():
        if key.lower() == "authorization":
            auth = str(value).strip()
            break
    if auth.lower().startswith("bearer "):
        return auth[7:].strip()
    return ""


def _json_response(code: int, body: Dict[str, Any]):
    if JSONResponse is None:  # pragma: no cover
        return body
    return JSONResponse(status_code=code, content=body)


async def _from_request(request: Request, route_path: str) -> Tuple[int, Dict[str, Any]]:
    # ``route_path`` is the plugin route (``/status``), not the mounted URL.
    # Hermes prefixes the router with ``/api/plugins/hermes-teams-inbox``.
    query = {str(k): str(v) for k, v in request.query_params.items()}
    headers = {str(k).lower(): str(v) for k, v in request.headers.items()}
    secret = _request_secret(headers)
    if not secret and request.method.upper() == "POST":
        try:
            payload = await request.json()
        except Exception:  # noqa: BLE001 — empty or non-JSON body is just "no secret"
            payload = None
        if isinstance(payload, dict):
            value = payload.get(SECRET_QUERY)
            if isinstance(value, str):
                secret = value.strip()
    client = request.client
    client_host = client.host if client is not None else ""
    return dispatch(request.method, route_path, query, secret, client_host)


@router.api_route("/status", methods=["GET", "POST"])
async def status(request: Request):
    code, body = await _from_request(request, "/status")
    return _json_response(code, body)


@router.api_route("/teams", methods=["GET", "POST"])
async def list_teams(request: Request):
    code, body = await _from_request(request, "/teams")
    return _json_response(code, body)


@router.api_route("/channels", methods=["GET", "POST"])
async def list_channels(request: Request):
    code, body = await _from_request(request, "/channels")
    return _json_response(code, body)


@router.api_route("/messages", methods=["GET", "POST"])
async def list_messages(request: Request):
    code, body = await _from_request(request, "/messages")
    return _json_response(code, body)


@router.api_route("/{full_path:path}", methods=["GET", "POST", "PUT", "PATCH", "DELETE", "HEAD"])
async def reject_other_paths(request: Request, full_path: str):
    """No generic Graph forwarder. Unknown plugin paths are refused."""
    route_path = "/" + full_path if full_path else "/"
    code, body = await _from_request(request, route_path)
    return _json_response(code, body)


class _ProxyHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def do_GET(self) -> None:  # noqa: N802
        self._handle("GET")

    def do_POST(self) -> None:  # noqa: N802
        self._handle("POST")

    def _handle(self, method: str) -> None:
        parsed = urllib.parse.urlparse(self.path)
        raw_query = urllib.parse.parse_qs(parsed.query, keep_blank_values=False)
        query = {key: values[-1] for key, values in raw_query.items() if values}
        secret = _request_secret(self.headers)
        if not secret and method == "POST":
            secret = _secret_from_body(self)
        client_host = self.client_address[0] if self.client_address else ""
        code, body = dispatch(method, parsed.path, query, secret, client_host)
        payload = json.dumps(body).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, fmt: str, *args: Any) -> None:
        # Request targets may carry the proxy secret. Do not log them.
        return


def _secret_from_body(handler: BaseHTTPRequestHandler) -> str:
    raw_length = handler.headers.get("Content-Length") or "0"
    try:
        length = int(raw_length)
    except ValueError:
        return ""
    if length <= 0 or length > 4096:
        return ""
    try:
        raw = handler.rfile.read(length)
        payload = json.loads(raw.decode("utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return ""
    if not isinstance(payload, dict):
        return ""
    value = payload.get(SECRET_QUERY)
    return value.strip() if isinstance(value, str) else ""


def make_server(port: int = 0, host: Optional[str] = None) -> ThreadingHTTPServer:
    """Bind the proxy. ``host`` defaults to loopback and other addresses are refused."""
    bind = resolve_bind_host(host)
    server = ThreadingHTTPServer((bind, port), _ProxyHandler)
    return server


def serve(port: Optional[int] = None, host: Optional[str] = None) -> None:
    """Run the proxy until interrupted. Listens on 127.0.0.1 by default."""
    chosen_port = port if port is not None else int((os.environ.get("HERMES_TEAMS_PROXY_PORT") or str(DEFAULT_PORT)).strip())
    server = make_server(chosen_port, host)
    bound_host, bound_port = server.server_address[:2]
    print(f"hermes-teams-inbox proxy listening on http://{bound_host}:{bound_port}", flush=True)
    thread = threading.Thread(target=server.serve_forever, name="hermes-teams-proxy", daemon=True)
    thread.start()
    try:
        thread.join()
    except KeyboardInterrupt:
        pass
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    serve()
