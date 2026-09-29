# Security

This plugin reads Microsoft Graph with the machine Azure CLI token (`az account get-access-token --resource https://graph.microsoft.com`). That token is a Graph credential. It is not authorization for this HTTP API.

## Trust boundary

- The standalone proxy binds to **127.0.0.1** (or `::1`) unless `HERMES_TEAMS_PROXY_ALLOW_REMOTE=1` is set. Do not publish it on a LAN or the public internet.
- When Hermes mounts the plugin routes, requests from non-loopback clients are rejected unless that same opt-in is set.
- Every Graph route requires a **per-install secret**. It is created on first local use at `<HERMES_HOME>/plugins/hermes-teams-inbox/proxy.secret` with mode `0600` (the directory is `0700`). On Windows, those mode bits are not a user ACL; the profile directory is what keeps other accounts out. The desktop client reads that file and sends it in a POST JSON body (`proxy_secret`). The API never returns the secret or the Azure CLI token, and it ignores `proxy_secret` on the query string so the value is not written into access logs.
- Mode `0600` stops other OS accounts. It does not stop other processes running as the same user. Any of those processes can read `proxy.secret`. Other Hermes Desktop plugins share the renderer file bridge (`readFileText`) and can read that path too. The secret is a boundary against other users, remote callers, and web pages, not against same-user code or sibling plugins.
- `AZ_CMD`, if set, must be an absolute path to the Azure CLI binary. Relative values are ignored.

## Graph surface

Only these reads are forwarded:

| App route | Graph |
| --- | --- |
| `/status` | `GET /me` (`displayName`, `mail`, `userPrincipalName`, `id`) |
| `/teams` | `GET /me/joinedTeams` |
| `/channels` | `GET /teams/{team-id}/channels` |
| `/messages` | `GET /teams/{team-id}/channels/{channel-id}/messages` (`$top=15`) |

`team_id` and `channel_id` are encoded with `safe=''`. A slash cannot open another Graph path. Ids that contain `/`, `\`, `..`, or `%` are rejected before any network call.

`/me` and joined teams work with a normal `az login`. Channel message bodies need `ChannelMessage.Read.All` on that token; without it the API returns `needs_scope` and does not invent messages.

Team and channel `webUrl` values are returned only for `https` hosts under `teams.microsoft.com`, `teams.microsoft.us`, `gov.teams.microsoft.us`, `teams.live.com`, and `teams.cloud.microsoft`. The desktop pane uses the same allowlist before opening a link.

## Desktop pane

This repo hosts the Teams web client in a persistent partition (`persist:hermes-teams`). In-pane navigation and new-window targets must be `https` on the Teams hosts above, or exactly `login.microsoftonline.com` or `login.microsoft.com`, so sign-in can finish in the same view. Other schemes and hosts are dropped and do not load. Pop-out uses the Teams host allowlist only. The proxy secret is read from `proxy.secret` and sent in a POST JSON body. It is not stored in plugin storage and it is not placed on the URL.

## Follow-up

A dedicated app registration with only the Graph scopes above would replace the Azure CLI user token. That is not what this plugin does today.
