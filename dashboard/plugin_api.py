"""Microsoft Teams inbox API for Hermes Desktop.

Mounted at /api/plugins/hermes-teams-inbox/.

Auth: Azure CLI's current Graph token (`az account get-access-token
--resource https://graph.microsoft.com`). No passwords, no secrets in git.
Joined teams + channels work with a standard `az login`. Chat/channel
messages need extra Graph scopes; those endpoints return needs_scope=true
instead of fabricating data.
"""
from __future__ import annotations

import json
import os
import subprocess
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

try:
    from fastapi import APIRouter, Query
except Exception:  # pragma: no cover
    class APIRouter:  # type: ignore
        def get(self, *_a, **_k):
            return lambda fn: fn

    def Query(default=None, **_k):  # type: ignore
        return default

try:
    from hermes_constants import get_hermes_home
except Exception:  # pragma: no cover
    def get_hermes_home() -> Path:  # type: ignore[misc]
        val = (os.environ.get("HERMES_HOME") or "").strip()
        return Path(val) if val else Path.home() / ".hermes"

router = APIRouter()

GRAPH = "https://graph.microsoft.com/v1.0"
AZ_CANDIDATES = (
    os.environ.get("AZ_CMD") or "",
    r"C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    r"C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd",
    "az.cmd",
    "az",
)


def _az_cmd() -> Optional[str]:
    for cand in AZ_CANDIDATES:
        if not cand:
            continue
        p = Path(cand)
        if p.is_file():
            return str(p)
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


def graph_get(path: str, token: str, query: Optional[Dict[str, str]] = None) -> Tuple[int, Dict[str, Any]]:
    url = GRAPH + path
    if query:
        url += "?" + urllib.parse.urlencode(query)
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


@router.get("/status")
def status() -> Dict[str, Any]:
    token, err = graph_token()
    if err or not token:
        return {
            "ok": False,
            "source": "azure-cli",
            "error": err or "no_token",
            "hint": "Run `az login` (Azure CLI is already the Graph credential source).",
            "teams_gateway_enabled": teams_platform_enabled(),
        }
    code, me = graph_get("/me", token, {"$select": "displayName,mail,userPrincipalName,id"})
    if code != 200:
        return {
            "ok": False,
            "source": "azure-cli",
            "error": ((me.get("error") or {}) if isinstance(me.get("error"), dict) else {}).get("message") or f"graph_{code}",
            "teams_gateway_enabled": teams_platform_enabled(),
        }
    return {
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


@router.get("/teams")
def list_teams() -> Dict[str, Any]:
    token, err = graph_token()
    if err or not token:
        return {"ok": False, "error": err or "no_token", "teams": []}
    code, data = graph_get("/me/joinedTeams", token, {"$select": "id,displayName,description,webUrl"})
    if code != 200:
        return {
            "ok": False,
            "error": ((data.get("error") or {}) if isinstance(data.get("error"), dict) else {}).get("message") or f"graph_{code}",
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
                "webUrl": row.get("webUrl") or "",
            }
        )
    return {"ok": True, "teams": teams}


@router.get("/channels")
def list_channels(team_id: str = Query(..., min_length=8)) -> Dict[str, Any]:
    token, err = graph_token()
    if err or not token:
        return {"ok": False, "error": err or "no_token", "channels": []}
    code, data = graph_get(f"/teams/{urllib.parse.quote(team_id)}/channels", token, {"$select": "id,displayName,webUrl,membershipType"})
    if code != 200:
        return {
            "ok": False,
            "error": ((data.get("error") or {}) if isinstance(data.get("error"), dict) else {}).get("message") or f"graph_{code}",
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
                "webUrl": row.get("webUrl") or "",
                "membershipType": row.get("membershipType") or "",
            }
        )
    return {"ok": True, "channels": channels}


@router.get("/messages")
def list_messages(
    team_id: str = Query(..., min_length=8),
    channel_id: str = Query(..., min_length=4),
) -> Dict[str, Any]:
    token, err = graph_token()
    if err or not token:
        return {"ok": False, "error": err or "no_token", "messages": []}
    path = f"/teams/{urllib.parse.quote(team_id)}/channels/{urllib.parse.quote(channel_id, safe='')}/messages"
    code, data = graph_get(path, token, {"$top": "15"})
    if code != 200:
        return {
            "ok": False,
            "error": ((data.get("error") or {}) if isinstance(data.get("error"), dict) else {}).get("message") or f"graph_{code}",
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
    return {"ok": True, "messages": messages}
