#!/bin/bash
# Install SMFWorks Teams Plugin into Hermes Desktop.
#
# Desktop UI loads from $HOME/.hermes/desktop-plugins (app-global).
# The Python API mounts only if the plugin is enabled on the profile
# that owns the window, and only after that backend restarts.
set -euo pipefail

REPO="${REPO:-https://github.com/smfworks/SMFWorks-Teams-Plugin.git}"
NAME=hermes-teams-inbox
APP_HOME="${HOME}/.hermes"
HERMES_BIN="${HERMES_BIN:-hermes}"

if ! command -v "$HERMES_BIN" >/dev/null 2>&1; then
  echo "hermes not on PATH" >&2
  exit 1
fi

homes=("$APP_HOME")
if [[ -n ${HERMES_HOME:-} && $HERMES_HOME != "$APP_HOME" ]]; then
  homes+=("$HERMES_HOME")
fi
shopt -s nullglob
for p in "$APP_HOME"/profiles/*/plugins; do
  homes+=("$(dirname "$(dirname "$p")")")
done
shopt -u nullglob
mapfile -t homes < <(printf '%s\n' "${homes[@]}" | awk 'NF && !seen[$0]++')

install_into() {
  local home=$1
  mkdir -p "$home/plugins"
  echo "==> HERMES_HOME=$home"
  if [[ -L $home/plugins/$NAME ]]; then
    echo "    symlink $home/plugins/$NAME -> $(readlink "$home/plugins/$NAME")"
  elif [[ -d $home/plugins/$NAME/.git ]]; then
    git -C "$home/plugins/$NAME" fetch --depth 1 origin main
    git -C "$home/plugins/$NAME" merge --ff-only FETCH_HEAD
  else
    env -u HERMES_PROFILE HERMES_HOME="$home" "$HERMES_BIN" plugins install "$REPO" --enable --no-deps || \
      git clone --depth 1 "$REPO" "$home/plugins/$NAME"
  fi
  env -u HERMES_PROFILE HERMES_HOME="$home" "$HERMES_BIN" plugins enable "$NAME" --no-allow-tool-override
}

SELF="$(cd "$(dirname "$0")" && pwd)"

src=""
for home in "${homes[@]}"; do
  install_into "$home"
  if [[ -f $home/plugins/$NAME/desktop/plugin.js ]]; then
    src="$home/plugins/$NAME/desktop/plugin.js"
  elif [[ -L $home/plugins/$NAME ]]; then
    candidate="$(readlink -f "$home/plugins/$NAME")/desktop/plugin.js"
    [[ -f $candidate ]] && src=$candidate
  fi
done

if [[ -f $SELF/desktop/plugin.js ]]; then
  src="$SELF/desktop/plugin.js"
fi

if [[ -z $src ]]; then
  echo "desktop/plugin.js not found" >&2
  exit 1
fi

mkdir -p "$APP_HOME/desktop-plugins/$NAME"
cp -f "$src" "$APP_HOME/desktop-plugins/$NAME/plugin.js"
echo "==> copied JS -> $APP_HOME/desktop-plugins/$NAME/plugin.js"
echo
echo "DONE. Restart the Hermes backend for the active profile so the"
echo "plugin API mounts, then: command palette → Reload desktop plugins."
echo "Sidebar → Teams. Requires: az login"
