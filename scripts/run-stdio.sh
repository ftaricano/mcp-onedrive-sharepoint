#!/usr/bin/env bash
set -euo pipefail

SCRIPT_PATH="$(python3 -c 'import os,sys; print(os.path.realpath(sys.argv[1]))' "${BASH_SOURCE[0]}")"
SCRIPT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

if [[ ! -f "$REPO_ROOT/build/index.js" ]]; then
  echo "Missing build/index.js. Run: npm run build" >&2
  exit 1
fi

exec node "$SCRIPT_DIR/exec-with-env.mjs" node "$REPO_ROOT/build/index.js"
