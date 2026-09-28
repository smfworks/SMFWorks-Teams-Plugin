# SMFWorks Teams Plugin

Hermes Desktop pane for Microsoft Teams. Plugin id: `hermes-teams-inbox`.

There is no official Hermes Desktop plugin for Microsoft Teams. Hermes already ships a gateway adapter (`hermes-teams`, a bot people message) and a meeting-summary pipeline. This repo is the desktop pane.

## What it does

- Sidebar **Teams** page in Hermes Desktop
- Lists joined teams and channels from Microsoft Graph
- Hosts the Teams web client in the pane, so chat, files, and calls stay inside Hermes
- **Pop out** still opens the same channel in the Teams app
- Uses the machine's Azure CLI session (`az account get-access-token --resource https://graph.microsoft.com`). No passwords in chat, no secrets in git

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

## Not this plugin

- The Hermes Teams gateway bot (`platforms.teams`)
- The meeting transcript pipeline
