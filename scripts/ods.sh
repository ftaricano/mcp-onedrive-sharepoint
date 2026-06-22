#!/usr/bin/env bash
# Thin wrapper so `ods` can live on $PATH via symlink (e.g. ~/bin/ods).
# Loads the repo-local .env and invokes build/cli.js.
set -euo pipefail

SCRIPT_PATH="$(python3 -c 'import os,sys; print(os.path.realpath(sys.argv[1]))' "${BASH_SOURCE[0]}")"
SCRIPT_DIR="$(cd "$(dirname "$SCRIPT_PATH")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

if [[ ! -f "$REPO_ROOT/build/cli.js" ]]; then
  echo "Missing build/cli.js. Run: npm run build" >&2
  exit 1
fi

if [[ -f "$REPO_ROOT/.env" ]]; then
  set -o allexport
  # shellcheck disable=SC1091
  source "$REPO_ROOT/.env"
  set +o allexport
fi

# Política "sempre Keychain" (Ferd 2026-06-09): o SP_CLIENT_SECRET vem do Keychain
# com PRECEDÊNCIA sobre o .env — a rotação pós-leak (2026-06-05) atualiza o Keychain,
# não o .env, então confiar no .env usa um secret revogado (AADSTS7000215). Com o
# secret presente o cli.js usa client_credentials (app-only, sem device code).
# Fallback: se o Keychain não tiver o item, mantém o que veio do .env.
__ods_kc_secret="$(security find-generic-password -s 'cpz::SP_CLIENT_SECRET' -w "$HOME/Library/Keychains/login.keychain-db" 2>/dev/null || true)"
if [[ -n "${__ods_kc_secret:-}" ]]; then
  export SP_CLIENT_SECRET="$__ods_kc_secret"
  export MICROSOFT_GRAPH_CLIENT_SECRET="$__ods_kc_secret"
fi
unset __ods_kc_secret

exec node "$REPO_ROOT/build/cli.js" "$@"
