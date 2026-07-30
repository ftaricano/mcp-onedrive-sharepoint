#!/bin/bash -p
# Thin wrapper so `ods` can live on $PATH via symlink (e.g. ~/bin/ods).
# Resolves Graph configuration from 1Password and invokes build/cli.js.
set -euo pipefail

SCRIPT_PATH="$(/usr/bin/python3 -I -c 'import os,sys; print(os.path.realpath(sys.argv[1]))' "${BASH_SOURCE[0]}")"
SCRIPT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

if [[ ! -f "$REPO_ROOT/build/cli.js" ]]; then
  echo "Missing build/cli.js. Run: npm run build" >&2
  exit 1
fi

exec "$SCRIPT_DIR/with-onepassword-graph-env.sh" node "$REPO_ROOT/build/cli.js" "$@"
