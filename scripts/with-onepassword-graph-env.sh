#!/bin/bash -p
set -euo pipefail

SCRIPT_PATH="$(/usr/bin/python3 -I -c 'import os,sys; print(os.path.realpath(sys.argv[1]))' "${BASH_SOURCE[0]}")"
SCRIPT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"

if [[ $# -eq 0 ]]; then
  echo "Usage: with-onepassword-graph-env.sh <command> [args...]" >&2
  exit 64
fi

# shellcheck disable=SC1091
source "$SCRIPT_DIR/onepassword-graph-env.sh"
load_graph_environment

exec "$@"
