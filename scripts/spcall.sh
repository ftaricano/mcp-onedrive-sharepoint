#!/usr/bin/env bash
set -euo pipefail

SCRIPT_PATH="$(python3 -c 'import os,sys; print(os.path.realpath(sys.argv[1]))' "${BASH_SOURCE[0]}")"
SCRIPT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

if [[ ! -f "$REPO_ROOT/build/index.js" ]]; then
  echo "Missing build/index.js. Run: npm run build" >&2
  exit 1
fi

if [[ $# -eq 0 ]]; then
  cat >&2 <<'USAGE'
Usage:
  spcall <tool> [arg=value ...] [--output json]

Examples:
  spcall health_check
  spcall list_files driveId=b!abc123 path=/Shared%20Documents
  spcall search_files query="quarterly report"
USAGE
  exit 64
fi

args=("$@")
default_output=1

for arg in "${args[@]}"; do
  if [[ "$arg" == "--output" || "$arg" == --output=* ]]; then
    default_output=0
    break
  fi
done

cmd=(
  npx -y mcporter call
  --stdio "$REPO_ROOT/scripts/run-stdio.sh"
  --cwd "$REPO_ROOT"
  --name sharepoint
)

cmd+=("${args[@]}")

if [[ $default_output -eq 1 ]]; then
  cmd+=(--output json)
fi

cleanup() {
  pkill -f "$REPO_ROOT/build/index.js" >/dev/null 2>&1 || true
  pkill -f "$REPO_ROOT/scripts/exec-with-env.mjs node $REPO_ROOT/build/index.js" >/dev/null 2>&1 || true
}

trap cleanup EXIT INT TERM

node "$SCRIPT_DIR/exec-with-env.mjs" "${cmd[@]}"
