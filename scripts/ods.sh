#!/usr/bin/env bash
# Thin wrapper so `ods` can live on $PATH via symlink (e.g. ~/bin/ods).
# Loads non-secret repo config and invokes build/cli.js.
set -euo pipefail

SCRIPT_PATH="$(python3 -c 'import os,sys; print(os.path.realpath(sys.argv[1]))' "${BASH_SOURCE[0]}")"
SCRIPT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

if [[ ! -f "$REPO_ROOT/build/cli.js" ]]; then
  echo "Missing build/cli.js. Run: npm run build" >&2
  exit 1
fi

# shellcheck disable=SC1091
source "$SCRIPT_DIR/onepassword-graph-env.sh"
load_graph_client_secret

exec node "$SCRIPT_DIR/exec-with-env.mjs" node "$REPO_ROOT/build/cli.js" "$@"
