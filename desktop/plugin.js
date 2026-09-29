/**
 * Hermes Desktop — Microsoft Teams (SMF Works).
 * Folder name must equal id (`hermes-teams-inbox`).
 * Roster: localhost Graph proxy (Azure CLI token + per-install secret, POST body).
 * Conversation: Teams web client in-pane.
 * Channel message reads need ChannelMessage.Read.All, which `az` does not have.
 */
import {
  Badge,
  Button,
  Codicon,
  EmptyState,
  ErrorState,
  GlyphSpinner,
  PALETTE_AREA,
  ROUTES_AREA,
  SIDEBAR_NAV_AREA,
  ScrollArea,
  Separator,
  Tip,
  cn,
  haptic,
  host,
  usePluginI18n,
  useQuery
} from '@hermes/plugin-sdk'
import { useEffect, useRef, useState } from 'react'
import { jsx, jsxs } from 'react/jsx-runtime'

const TEAMS_HOME = 'https://teams.cloud.microsoft/v2/'
const TEAMS_WEB_HOSTS = [
  'teams.microsoft.com',
  'teams.microsoft.us',
  'gov.teams.microsoft.us',
  'teams.live.com',
  'teams.cloud.microsoft'
]
const LOGIN_HOSTS = new Set(['login.microsoftonline.com', 'login.microsoft.com'])
const CHROME_UA =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'

const ID = 'hermes-teams-inbox'
const ROUTE = '/teams'

let pluginCtx = null
let secretCache = ''

function isTeamsHost(host) {
  return TEAMS_WEB_HOSTS.some((name) => host === name || host.endsWith('.' + name))
}

function isAllowedTeamsUrl(url) {
  try {
    const parsed = new URL(url)
    if (parsed.protocol !== 'https:') return false
    if (parsed.username || parsed.password) return false
    const host = parsed.hostname.toLowerCase().replace(/\.$/, '')
    return isTeamsHost(host)
  } catch {
    return false
  }
}

function isAllowedWebviewUrl(url) {
  try {
    const parsed = new URL(url)
    if (parsed.protocol !== 'https:') return false
    if (parsed.username || parsed.password) return false
    const host = parsed.hostname.toLowerCase().replace(/\.$/, '')
    return isTeamsHost(host) || LOGIN_HOSTS.has(host)
  } catch {
    return false
  }
}

function openUrl(url) {
  if (!isAllowedTeamsUrl(url)) return
  const opener = pluginCtx?.os?.openExternal
  if (typeof opener === 'function') {
    void opener(url)
    return
  }
  window.open(url, '_blank', 'noopener,noreferrer')
}

function isLauncherPath(pathname) {
  return pathname === '/l' || pathname.startsWith('/l/') || pathname.startsWith('/dl/')
}

function webClientUrl(url) {
  const admitted = teamsUrl(url)
  if (!admitted) return null
  try {
    const parsed = new URL(admitted)
    // Only the client picker is a dead end. Every other Teams path, including
    // the ones the web app uses while it boots, has to be left alone.
    if (isLauncherPath(parsed.pathname)) return TEAMS_HOME
    return admitted
  } catch {
    return null
  }
}

function followUrl(url) {
  if (!isAllowedWebviewUrl(url)) return null
  try {
    const parsed = new URL(url)
    const host = parsed.hostname.toLowerCase().replace(/\.$/, '')
    // Picker buttons call window.open. Follow Teams and sign-in URLs in this
    // pane. Anything else stays out of the persistent partition.
    if (isTeamsHost(host) && isLauncherPath(parsed.pathname)) return TEAMS_HOME
    return parsed.toString()
  } catch {
    return null
  }
}

function teamsUrl(url) {
  if (!isAllowedWebviewUrl(url)) return null
  try {
    return new URL(url).toString()
  } catch {
    return null
  }
}

const DISMISS_APP_GATE = `(() => {
  const nodes = Array.from(document.querySelectorAll('button, a'))
  const web = nodes.find((el) => (el.innerText || el.textContent || '').toLowerCase().indexOf('use the web app instead') !== -1)
  if (!web) return 'none'
  web.click()
  return 'clicked'
})()`

function dismissAppGate(webview) {
  try {
    void webview.executeJavaScript(DISMISS_APP_GATE, false)
  } catch {
    // Guest not ready yet.
  }
}

function joinPath(root, ...parts) {
  const slash = String(root).includes('\\') ? '\\' : '/'
  const pieces = [String(root).replace(/[\\/]+$/, '')]
  for (const part of parts) {
    pieces.push(String(part).replace(/^[\\/]+|[\\/]+$/g, ''))
  }
  return pieces.join(slash)
}

async function resolvePluginRoot(bridge) {
  if (typeof bridge.agentPluginsRoot === 'function') {
    try {
      const root = await bridge.agentPluginsRoot()
      if (root) return String(root)
    } catch {
      // Typed on the desktop bridge, but current preload builds do not bind it.
    }
  }
  // logsRoot is the implemented twin: same profile home, final segment "logs".
  // Agent plugins live in the sibling "plugins" directory.
  if (typeof bridge.logsRoot !== 'function') return ''
  try {
    const logs = String((await bridge.logsRoot()) || '')
    const parts = logs.split(/[/\\]/)
    if (!parts.length || parts[parts.length - 1].toLowerCase() !== 'logs') return ''
    const slash = logs.includes('\\') ? '\\' : '/'
    parts[parts.length - 1] = 'plugins'
    return parts.join(slash)
  } catch {
    return ''
  }
}

async function readInstalledSecret() {
  const bridge = typeof window !== 'undefined' ? window.hermesDesktop : null
  if (!bridge || typeof bridge.readFileText !== 'function') return ''
  const root = await resolvePluginRoot(bridge)
  if (!root) return ''
  try {
    const result = await bridge.readFileText(joinPath(root, ID, 'proxy.secret'))
    const text = String(result?.text || '').trim()
    if (!text || result?.truncated) return ''
    return text.split(/\s+/)[0]
  } catch {
    return ''
  }
}

async function ensureSecret(force) {
  if (secretCache && !force) return secretCache
  const secret = await readInstalledSecret()
  secretCache = secret || ''
  return secretCache
}

async function rest(path) {
  if (!pluginCtx?.rest) {
    const err = new Error('backend-off')
    err.backend = false
    throw err
  }
  let secret = await ensureSecret(false)
  if (!secret) {
    // A rejected local call creates the secret file. It is not in the response.
    let probeError = null
    try {
      await pluginCtx.rest(path)
    } catch (err) {
      probeError = err
    }
    secret = await ensureSecret(true)
    if (!secret) {
      const probeMessage = String(probeError?.message || '')
      if (probeError && !/401|unauthorized/i.test(probeMessage)) {
        throw probeError
      }
      const err = new Error('proxy-secret-unavailable')
      throw err
    }
  }
  // pluginCtx.rest accepts method and body, not custom headers. The host
  // JSON-encodes the body, so the secret is not part of the request line.
  const send = (value) =>
    pluginCtx.rest(path, {
      method: 'POST',
      body: { proxy_secret: value }
    })
  try {
    return await send(secret)
  } catch (err) {
    secretCache = ''
    const fresh = await readInstalledSecret()
    if (fresh && fresh !== secret) {
      secretCache = fresh
      return send(fresh)
    }
    throw err
  }
}

function TeamsEmbed({ url }) {
  const hostRef = useRef(null)
  const viewRef = useRef(null)
  const readyRef = useRef(false)
  const urlRef = useRef(url)
  urlRef.current = url

  useEffect(() => {
    const node = hostRef.current
    if (!node) return undefined
    readyRef.current = false
    const webview = document.createElement('webview')
    webview.setAttribute('partition', 'persist:hermes-teams')
    webview.setAttribute('allowpopups', '')
    webview.setAttribute('webpreferences', 'contextIsolation=yes,nodeIntegration=no,sandbox=no')
    webview.setAttribute('useragent', CHROME_UA)
    webview.setAttribute('src', webClientUrl(urlRef.current) || TEAMS_HOME)
    webview.style.border = '0'
    webview.style.display = 'flex'
    webview.style.width = '100%'
    webview.style.height = '100%'
    node.replaceChildren(webview)
    viewRef.current = webview

    const fit = () => {
      const rect = node.getBoundingClientRect()
      webview.style.width = Math.max(1, Math.floor(rect.width)) + 'px'
      webview.style.height = Math.max(1, Math.floor(rect.height)) + 'px'
    }
    fit()
    const observer = new ResizeObserver(fit)
    observer.observe(node)

    const go = (next) => {
      const target = webClientUrl(next) || TEAMS_HOME
      if (!readyRef.current) {
        webview.setAttribute('src', target)
        return
      }
      try {
        if (webview.getURL() === target) return
      } catch {
        // Guest not attached yet. src is already the web client.
      }
      try {
        webview.loadURL(target)
      } catch {
        webview.setAttribute('src', target)
      }
    }

    const onWillNavigate = (event) => {
      const next = event.url
      if (!next) return
      const target = webClientUrl(next)
      if (!target) {
        event.preventDefault()
        return
      }
      if (target === next) return
      event.preventDefault()
      go(target)
    }

    // Picker buttons call window.open. Follow that URL in this pane.
    // Bouncing an auth popup back to /v2/ leaves the logo on screen forever.
    const onNewWindow = (event) => {
      event.preventDefault()
      const next = followUrl(event.url)
      if (!next) return
      try {
        webview.loadURL(next)
      } catch {
        webview.setAttribute('src', next)
      }
    }

    const onReady = () => {
      readyRef.current = true
      try {
        webview.focus()
      } catch {
        // Focus is best-effort. The guest still receives clicks once attached.
      }
      try {
        webview.setBackgroundThrottling(false)
      } catch {
        // Older Electron builds omit this. Boot still proceeds.
      }
      dismissAppGate(webview)
      let current = ''
      try {
        current = webview.getURL()
      } catch {
        current = ''
      }
      try {
        if (current && isLauncherPath(new URL(current).pathname)) go(TEAMS_HOME)
      } catch {
        // A non-URL guest location is not a launcher. Leave it.
      }
    }

    const onStop = () => dismissAppGate(webview)

    webview.addEventListener('dom-ready', onReady)
    webview.addEventListener('did-stop-loading', onStop)
    webview.addEventListener('will-navigate', onWillNavigate)
    webview.addEventListener('new-window', onNewWindow)
    const dismissTimer = setInterval(() => dismissAppGate(webview), 1500)
    const dismissStop = setTimeout(() => clearInterval(dismissTimer), 20000)
    return () => {
      clearInterval(dismissTimer)
      clearTimeout(dismissStop)
      observer.disconnect()
      webview.removeEventListener('dom-ready', onReady)
      webview.removeEventListener('did-stop-loading', onStop)
      webview.removeEventListener('will-navigate', onWillNavigate)
      webview.removeEventListener('new-window', onNewWindow)
      readyRef.current = false
      viewRef.current = null
      node.replaceChildren()
    }
  }, [])

  useEffect(() => {
    const webview = viewRef.current
    const target = webClientUrl(url) || TEAMS_HOME
    if (!webview) return
    if (!readyRef.current) {
      webview.setAttribute('src', target)
      return
    }
    try {
      if (webview.getURL() === target) return
      webview.loadURL(target)
    } catch {
      webview.setAttribute('src', target)
    }
  }, [url])

  return jsx('div', { ref: hostRef, className: 'relative min-h-0 min-w-0 flex-1' })
}

function TeamsPage() {
  const t = usePluginI18n(ID)
  const [teamId, setTeamId] = useState(null)
  const [channelId, setChannelId] = useState(null)
  const [embedUrl, setEmbedUrl] = useState(TEAMS_HOME)
  const status = useQuery({
    queryKey: [ID, 'status'],
    queryFn: () => rest('/status'),
    staleTime: 30 * 1000
  })
  const teams = useQuery({
    queryKey: [ID, 'teams'],
    queryFn: () => rest('/teams'),
    enabled: Boolean(status.data?.ok),
    staleTime: 30 * 1000
  })
  const channels = useQuery({
    queryKey: [ID, 'channels', teamId],
    queryFn: () => rest(`/channels?team_id=${encodeURIComponent(teamId)}`),
    enabled: Boolean(teamId),
    staleTime: 30 * 1000
  })

  const list = teams.data?.teams || []
  const chList = channels.data?.channels || []

  useEffect(() => {
    if (teamId || !list.length) return
    setTeamId(list[0].id)
  }, [teamId, list])

  useEffect(() => {
    if (!chList.length) return
    const current = chList.find((row) => row.id === channelId)
    const next = current || chList[0]
    if (next.id !== channelId) setChannelId(next.id)
  }, [teamId, channelId, chList])

  if (status.isLoading) {
    return jsx('div', {
      className: 'flex h-full items-center justify-center',
      children: jsx(GlyphSpinner, {})
    })
  }

  if (status.isError || status.data?.ok === false) {
    const message = String(status.error?.message || '')
    const backendOff = message === 'backend-off' || /Plugin not found/.test(message)
    const secretMissing = message === 'proxy-secret-unavailable'
    return jsx(ErrorState, {
      title: backendOff ? t('backendError') : t('authError'),
      description: backendOff
        ? t('backendHint')
        : secretMissing
          ? t('proxySecretHint')
          : status.data?.hint || status.data?.error || message,
      action: jsx(Button, {
        size: 'sm',
        onClick: () => {
          haptic('tap')
          void status.refetch()
        },
        children: t('retry')
      })
    })
  }

  const me = status.data?.me || {}
  const selected = list.find((row) => row.id === teamId) || null
  const selectedChannel = chList.find((row) => row.id === channelId) || null
  const popTarget = teamsUrl(selectedChannel?.webUrl) || teamsUrl(selected?.webUrl) || embedUrl

  return jsxs('div', {
    className: 'flex h-full min-h-0 flex-col',
    children: [
      jsxs('div', {
        className: 'flex items-center gap-2 px-3 py-2',
        children: [
          jsx(Codicon, { name: 'organization', className: 'text-(--ui-text-tertiary)' }),
          jsxs('div', {
            className: 'min-w-0 flex-1',
            children: [
              jsx('div', { className: 'truncate text-sm font-medium', children: t('paneTitle') }),
              jsx('div', {
                className: 'truncate text-[0.6875rem] text-(--ui-text-tertiary)',
                children: me.displayName || me.userPrincipalName || t('signedOut')
              })
            ]
          }),
          status.data?.teams_gateway_enabled
            ? jsx(Badge, { children: t('gatewayOn') })
            : jsx(Badge, { children: t('gatewayOff') }),
          jsx(Tip, {
            label: t('popTip'),
            children: jsx(Button, {
              size: 'sm',
              variant: 'ghost',
              onClick: () => {
                haptic('tap')
                openUrl(popTarget)
              },
              children: t('popOut')
            })
          }),
          jsx(Tip, {
            label: t('refreshTip'),
            children: jsx(Button, {
              size: 'sm',
              variant: 'ghost',
              onClick: () => {
                haptic('tap')
                void status.refetch()
                void teams.refetch()
                if (teamId) void channels.refetch()
              },
              children: t('refresh')
            })
          })
        ]
      }),
      jsx(Separator, {}),
      jsxs('div', {
        className: 'flex min-h-0 flex-1',
        children: [
          jsx(ScrollArea, {
            className: 'w-[180px] shrink-0 border-r border-(--ui-stroke-secondary)',
            children: jsx('div', {
              className: 'flex flex-col p-1',
              children: list.length
                ? list.map((row) =>
                    jsx(
                      'button',
                      {
                        type: 'button',
                        className: cn(
                          'flex w-full items-center gap-2 rounded px-2 py-1.5 text-left text-sm',
                          row.id === teamId
                            ? 'bg-(--chrome-action-hover) text-foreground'
                            : 'text-(--ui-text-secondary) hover:bg-(--chrome-action-hover)'
                        ),
                        onClick: () => {
                          haptic('tap')
                          setTeamId(row.id)
                          setChannelId(null)
                        },
                        children: jsx('span', { className: 'truncate', children: row.displayName })
                      },
                      row.id
                    )
                  )
                : jsx('div', {
                    className: 'p-3 text-xs text-(--ui-text-tertiary)',
                    children: teams.isLoading ? t('loading') : t('noTeams')
                  })
            })
          }),
          jsx(ScrollArea, {
            className: 'w-[200px] shrink-0 border-r border-(--ui-stroke-secondary)',
            children: jsx('div', {
              className: 'flex flex-col p-1',
              children: !selected
                ? jsx('div', {
                    className: 'p-3 text-xs text-(--ui-text-tertiary)',
                    children: t('pickTeam')
                  })
                : channels.isLoading
                  ? jsx('div', {
                      className: 'flex justify-center p-4',
                      children: jsx(GlyphSpinner, {})
                    })
                  : chList.length
                    ? chList.map((ch) =>
                        jsx(
                          'button',
                          {
                            type: 'button',
                            className: cn(
                              'flex w-full items-center gap-2 rounded px-2 py-1.5 text-left text-sm',
                              ch.id === channelId
                                ? 'bg-(--chrome-action-hover) text-foreground'
                                : 'text-(--ui-text-secondary) hover:bg-(--chrome-action-hover)'
                            ),
                            onClick: () => {
                              haptic('tap')
                              setChannelId(ch.id)
                            },
                            children: [
                              jsx(Codicon, {
                                name: 'comment-discussion',
                                className: 'text-(--ui-text-quaternary)'
                              }),
                              jsx('span', { className: 'truncate', children: ch.displayName })
                            ]
                          },
                          ch.id
                        )
                      )
                    : jsx('div', {
                        className: 'p-3 text-xs text-(--ui-text-tertiary)',
                        children: channels.data?.error || t('noChannels')
                      })
            })
          }),
          jsx(TeamsEmbed, { url: embedUrl })
        ]
      })
    ]
  })
}

export default {
  id: ID,
  name: 'Microsoft Teams',
  defaultEnabled: true,
  register(ctx) {
    pluginCtx = ctx
    ctx.i18n.register({
      en: {
        paneTitle: 'Teams',
        signedOut: 'Not signed in',
        gatewayOn: 'gateway on',
        gatewayOff: 'gateway off',
        refresh: 'Refresh',
        refreshTip: 'Reload the team roster',
        retry: 'Retry',
        loading: 'Loading…',
        noTeams: 'No joined teams',
        pickTeam: 'Pick a team',
        popOut: 'Pop out',
        popTip: 'Open this channel in the Teams app',
        noChannels: 'No channels',
        backendError: 'Teams backend is off',
        backendHint: 'This profile has not enabled hermes-teams-inbox. Enable it in that profile and restart the backend.',
        authError: 'Graph sign-in needed',
        proxySecretHint:
          'The Teams proxy secret was not readable. It is created on first backend start at <HERMES_HOME>/plugins/hermes-teams-inbox/proxy.secret (mode 0600). Reload Desktop after az login.',
        chipTip: 'Microsoft Teams inbox'
      }
    })

    ctx.registerMany([
      {
        id: 'page',
        area: ROUTES_AREA,
        data: { path: ROUTE },
        render: () => jsx(TeamsPage, {})
      },
      {
        id: 'nav',
        area: SIDEBAR_NAV_AREA,
        data: { path: ROUTE, label: 'Teams', codicon: 'organization' }
      },
      {
        id: 'open',
        area: PALETTE_AREA,
        data: {
          id: 'hermes-teams-inbox.open',
          label: 'Open Microsoft Teams',
          keywords: ['teams', 'microsoft', 'graph', 'm365'],
          run: () => host.navigate(ROUTE)
        }
      }
    ])
  }
}
