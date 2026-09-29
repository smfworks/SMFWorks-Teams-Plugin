"""Proxy auth, Graph allowlist, and path-encoding tests. No live tokens."""
from __future__ import annotations

import json
import threading
import urllib.error
import urllib.parse
import urllib.request

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient

from dashboard import plugin_api as api


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setenv("HERMES_HOME", str(tmp_path))
    monkeypatch.setattr(api, "get_hermes_home", lambda: tmp_path)
    monkeypatch.delenv("HERMES_TEAMS_PROXY_ALLOW_REMOTE", raising=False)
    monkeypatch.delenv("HERMES_TEAMS_PROXY_HOST", raising=False)
    monkeypatch.delenv("AZ_CMD", raising=False)
    return tmp_path


@pytest.fixture
def client(home):
    del home
    app = FastAPI()
    app.include_router(api.router)
    return TestClient(app, client=("127.0.0.1", 50000))


def _secret() -> str:
    return api.ensure_proxy_secret()


def _block_network(monkeypatch):
    def explode(*_args, **_kwargs):
        raise AssertionError("Graph must not be contacted")

    monkeypatch.setattr(api.urllib.request, "urlopen", explode)


def test_unauthenticated_requests_are_rejected(client, monkeypatch):
    def boom():
        raise AssertionError("Azure CLI must not run")

    monkeypatch.setattr(api, "graph_token", boom)
    _block_network(monkeypatch)
    for path in ("/status", "/teams", "/channels", "/messages"):
        response = client.get(path)
        assert response.status_code == 401
        body = response.json()
        assert body["ok"] is False
        assert body["error"] == "unauthorized"
        assert "accessToken" not in response.text
        assert "proxy.secret" not in response.text

    wrong = client.post("/status", json={"proxy_secret": "not-the-secret"})
    assert wrong.status_code == 401
    assert wrong.json()["error"] == "unauthorized"


def test_disallowed_paths_are_rejected(client, monkeypatch):
    _block_network(monkeypatch)
    def boom():
        raise AssertionError("Azure CLI must not run")

    monkeypatch.setattr(api, "graph_token", boom)
    secret = _secret()
    for path in (
        "/me/drive",
        "/users",
        "/applications",
        "/teams/abcd1234/channels/extra",
        "/me/joinedTeams/../applications",
    ):
        code, body = api.graph_get(path, "token")
        assert code == 403
        assert body["error"] == "disallowed_path"

    extra = api.graph_get("/me", "token", {"$select": api._ME_SELECT, "$filter": "true"})
    assert extra[0] == 403
    assert extra[1]["error"] == "disallowed_path"

    response = client.get("/applications", headers={api.SECRET_HEADER: secret})
    assert response.status_code == 403
    assert response.json()["error"] == "disallowed_path"


def test_team_id_slash_cannot_change_target_path(client, monkeypatch):
    _block_network(monkeypatch)
    monkeypatch.setattr(api, "graph_token", lambda: ("token", None))
    team_id = "abcd1234/../applications"
    path = api.graph_target_path("channels", team_id=team_id)
    encoded = urllib.parse.quote(team_id, safe="")
    assert path == f"/teams/{encoded}/channels"
    assert path.split("/") == ["", "teams", encoded, "channels"]
    assert "%2F" in path
    assert path.split("/")[2] == encoded
    assert "applications" not in path.split("/")

    code, body = api.graph_get(path, "token", {"$select": api._CHANNELS_SELECT})
    assert code == 400
    assert body["error"] == "invalid_team_id"

    secret = _secret()
    response = client.post(
        "/channels",
        params={"team_id": team_id},
        json={"proxy_secret": secret},
    )
    assert response.status_code == 400
    assert response.json()["error"] == "invalid_team_id"
    assert "graph.microsoft.com" not in response.text


def test_channel_id_slash_stays_inside_one_segment():
    team_id = "12345678-1234-1234-1234-123456789abc"
    channel_id = "19:abc@thread.tacv2/../me"
    path = api.graph_target_path("messages", team_id=team_id, channel_id=channel_id)
    encoded = urllib.parse.quote(channel_id, safe="")
    assert path.split("/") == ["", "teams", team_id, "channels", encoded, "messages"]
    with pytest.raises(api.GraphPathError):
        api.approved_graph_request("messages", team_id=team_id, channel_id=channel_id)


def test_secret_file_is_owner_only_and_stable(home):
    first = api.ensure_proxy_secret()
    second = api.ensure_proxy_secret()
    path = api.proxy_secret_path()
    assert first == second
    assert path.read_text(encoding="utf-8").strip() == first
    assert path.stat().st_mode & 0o777 == 0o600
    assert path.parent.stat().st_mode & 0o777 == 0o700
    assert path.is_relative_to(home)


def test_desktop_post_body_and_header_are_accepted(client, monkeypatch):
    monkeypatch.setattr(api, "graph_token", lambda: (None, "azure_cli_missing"))
    secret = _secret()
    posted = client.post("/status", json={"proxy_secret": secret})
    assert posted.status_code == 200
    assert posted.json()["ok"] is False
    assert "az login" in posted.json()["hint"]
    header = client.get("/teams", headers={api.SECRET_HEADER: secret})
    assert header.status_code == 200
    assert header.json()["teams"] == []
    query = client.get("/status", params={"proxy_secret": secret})
    assert query.status_code == 401
    assert query.json()["error"] == "unauthorized"
    bearer = client.get("/status", headers={"Authorization": f"Bearer {secret}"})
    assert bearer.status_code == 200


def test_remote_client_is_rejected_by_default(home, monkeypatch):
    del home
    app = FastAPI()
    app.include_router(api.router)
    remote = TestClient(app, client=("203.0.113.8", 443))
    def boom():
        raise AssertionError("Azure CLI must not run")

    monkeypatch.setattr(api, "graph_token", boom)
    secret = _secret()
    response = remote.get("/status", headers={api.SECRET_HEADER: secret})
    assert response.status_code == 403
    assert response.json()["error"] == "loopback_only"


def test_proxy_listens_on_localhost_by_default(home, monkeypatch):
    del home
    assert api.DEFAULT_BIND_HOST == "127.0.0.1"
    assert api.resolve_bind_host() == "127.0.0.1"
    assert api.resolve_bind_host("localhost") == "127.0.0.1"
    with pytest.raises(RuntimeError, match="loopback"):
        api.resolve_bind_host("0.0.0.0")
    monkeypatch.setenv("HERMES_TEAMS_PROXY_ALLOW_REMOTE", "1")
    assert api.resolve_bind_host("0.0.0.0") == "0.0.0.0"
    monkeypatch.delenv("HERMES_TEAMS_PROXY_ALLOW_REMOTE")
    monkeypatch.setattr(api, "graph_token", lambda: (None, "azure_cli_missing"))
    server = api.make_server(0)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        host, port = server.server_address[:2]
        assert host == "127.0.0.1"
        request = urllib.request.Request(f"http://127.0.0.1:{port}/status")
        with pytest.raises(urllib.error.HTTPError) as caught:
            urllib.request.urlopen(request, timeout=5)
        assert caught.value.code == 401
        payload = json.loads(caught.value.read().decode("utf-8"))
        assert payload["error"] == "unauthorized"
        secret = api.ensure_proxy_secret()
        leaked = urllib.request.Request(
            f"http://127.0.0.1:{port}/status?proxy_secret={urllib.parse.quote(secret)}"
        )
        with pytest.raises(urllib.error.HTTPError) as leaked_error:
            urllib.request.urlopen(leaked, timeout=5)
        assert leaked_error.value.code == 401
        authed = urllib.request.Request(
            f"http://127.0.0.1:{port}/status",
            data=json.dumps({"proxy_secret": secret}).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(authed, timeout=5) as resp:
            assert resp.status == 200
            body = json.loads(resp.read().decode("utf-8"))
            assert body["ok"] is False
            assert "az login" in body["hint"]
    finally:
        server.shutdown()
        server.server_close()


def test_relative_az_cmd_is_ignored(monkeypatch):
    monkeypatch.setenv("AZ_CMD", "evil/az")
    monkeypatch.setattr(api, "_AZ_FALLBACKS", ())
    assert api._az_cmd() is None
    monkeypatch.setenv("AZ_CMD", "/tmp/missing-azure-cli")
    assert api._az_cmd() is None


def test_web_url_allowlist():
    assert api.safe_web_url("https://teams.microsoft.com/l/team/1").startswith("https://teams.microsoft.com/")
    assert api.safe_web_url("https://gov.teams.microsoft.us/l/channel/1").startswith("https://")
    assert api.safe_web_url("https://teams.microsoft.com.evil.com/l/team/1") == ""
    assert api.safe_web_url("http://teams.microsoft.com/l/team/1") == ""
    assert api.safe_web_url("javascript:alert(1)") == ""
    assert api.safe_web_url("https://user:pass@teams.microsoft.com/l/team/1") == ""


def test_channels_response_shape_and_web_url_filter(client, monkeypatch):
    monkeypatch.setattr(api, "graph_token", lambda: ("token", None))

    def fake_graph_get(path, token, query=None):
        del token, query
        assert path == "/teams/12345678-1234-1234-1234-123456789abc/channels"
        return 200, {
            "value": [
                {
                    "id": "19:abc@thread.tacv2",
                    "displayName": "General",
                    "webUrl": "https://teams.microsoft.com/l/channel/1",
                    "membershipType": "standard",
                },
                {
                    "id": "19:bad",
                    "displayName": "Nope",
                    "webUrl": "javascript:alert(1)",
                    "membershipType": "private",
                },
            ]
        }

    monkeypatch.setattr(api, "graph_get", fake_graph_get)
    response = client.post(
        "/channels",
        params={"team_id": "12345678-1234-1234-1234-123456789abc"},
        json={"proxy_secret": _secret()},
    )
    assert response.status_code == 200
    body = response.json()
    assert body["ok"] is True
    assert body["channels"][0]["displayName"] == "General"
    assert body["channels"][0]["webUrl"] == "https://teams.microsoft.com/l/channel/1"
    assert body["channels"][1]["webUrl"] == ""


def test_allowlisted_channels_url_is_one_segment(monkeypatch):
    captured = {}

    class _Response:
        status = 200

        def read(self):
            return b'{"value":[]}'

        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

    def fake_urlopen(req, timeout=30):
        del timeout
        captured["url"] = req.full_url
        return _Response()

    monkeypatch.setattr(api.urllib.request, "urlopen", fake_urlopen)
    team_id = "12345678-1234-1234-1234-123456789abc"
    path, query = api.approved_graph_request("channels", team_id=team_id)
    code, body = api.graph_get(path, "token", query)
    assert code == 200
    assert body == {"value": []}
    assert captured["url"].startswith("https://graph.microsoft.com/v1.0/teams/")
    assert captured["url"].split("?", 1)[0].split("/")[-2:] == [team_id, "channels"]
