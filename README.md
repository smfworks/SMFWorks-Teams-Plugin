# SMFWorks Teams Plugin

Hermes Desktop pane for Microsoft Teams. Plugin id: `hermes-teams-inbox`.

There is no official Hermes Desktop plugin for Microsoft Teams. Hermes already ships a gateway adapter (`hermes-teams`, a bot people message) and a meeting-summary pipeline. This repo is the desktop pane.

## What it does

- Sidebar **Teams** page in Hermes Desktop
- Lists joined teams and channels from Microsoft Graph
- Hosts the Teams web client in the pane, so chat, files, and calls stay inside Hermes
- **Pop out** still opens the same channel in the Teams app
- Uses the machine's Azure CLI session (`az account get-access-token --resource https://graph.microsoft.com`). No passwords in chat, no secrets in git
- Talks to Graph only through a **localhost proxy**. Every request needs a per-install secret the desktop client sends. See [Proxy auth](#proxy-auth)

Channel message bodies need `ChannelMessage.Read.All`. A normal `az login` does not have that scope, so the pane hosts the web client instead of inventing an empty inbox.

The Graph channel link (`/l/...`) is a client picker. Its buttons open the desktop app and do nothing inside the pane. The pane loads `https://teams.cloud.microsoft/v2/` and follows sign-in in the same view.

## Install

Desktop UI is app-global. The Python API loads only when the plugin is in the **active profile's** `plugins.enabled`. A plugin enabled only on the default home returns `404 {"detail":"Plugin not found"}` from a named profile. Routes mount at backend start, so enable it, then restart that backend.

```text
HERMES_HOME/desktop-plugins/hermes-teams-inbox/plugin.js
HERMES_HOME/plugins/hermes-teams-inbox/          (Python API)
```

The folder name must equal the plugin id.

Windows, from this checkout:

```text
copy desktop\plugin.js  %LOCALAPPDATA%\hermes\desktop-plugins\hermes-teams-inbox\plugin.js
xcopy /E /I .           %LOCALAPPDATA%\hermes\plugins\hermes-teams-inbox\
```

If the window's backend is a named profile, copy the Python tree into that profile's `plugins\hermes-teams-inbox\` as well, or rely on the default-root scan if that backend already reads it.

Then:

1. `hermes plugins enable hermes-teams-inbox` with `HERMES_HOME` set to the profile that owns the window
2. Restart that backend
3. Command palette: **Reload desktop plugins**
4. Sidebar → **Teams**

`install.sh` does the Unix copy and enable. Requires Azure CLI signed in (`az login`) with Graph access to `/me` and `/me/joinedTeams`.

## Proxy auth

The plugin API is not open just because Azure CLI is logged in. On first use it writes a secret to:

```text
<HERMES_HOME>/plugins/hermes-teams-inbox/proxy.secret
```

The file is mode `0600` and the directory is `0700`. Hermes Desktop reads that file (same Hermes home as the backend) and sends it in a POST JSON body (`proxy_secret`), not on the URL. You do not copy it by hand for a normal local install. Other callers send `X-Hermes-Teams-Proxy-Secret` or `Authorization: Bearer`. A `proxy_secret` query parameter is ignored.

A standalone server, if you run one, listens on **127.0.0.1:8765** only:

```text
python -m dashboard.plugin_api
```

Run that from the plugin directory (the repo root).

`HERMES_TEAMS_PROXY_PORT` changes the port. Binding any other address, or accepting non-loopback clients when Hermes mounts the routes, requires `HERMES_TEAMS_PROXY_ALLOW_REMOTE=1`. Leave that unset.

Only `/me`, joined teams, a team's channels, and a channel's messages are forwarded. `team_id` / `channel_id` are fully URL-encoded, and values containing `/` are rejected. **Pop out** only follows `https` links on Microsoft Teams hosts. The in-pane web client follows those same hosts, plus `login.microsoftonline.com` and `login.microsoft.com`, so sign-in can finish in the pane. Other `https` URLs are not loaded into the persistent partition.

Tests (no live token): `pip install -r requirements-dev.txt && python -m pytest -q`.

## Not this plugin

- The Hermes Teams gateway bot (`platforms.teams`)
- The meeting transcript pipeline
