#!/usr/bin/env bash

load_graph_client_secret() {
  local hub secret
  if [[ -n "${MICROSOFT_GRAPH_CLIENT_SECRET:-${SP_CLIENT_SECRET:-}}" ]]; then
    secret="${MICROSOFT_GRAPH_CLIENT_SECRET:-${SP_CLIENT_SECRET}}"
  else
    hub="${JARVIS_HUB:-${HOME}/jarvis-hub}"
    secret="$(
      PYTHONPATH="${hub}/scripts/infra${PYTHONPATH:+:${PYTHONPATH}}" \
        python3 -c 'from cpz_keychain import get; print(get("SP_CLIENT_SECRET"), end="")'
    )"
  fi
  if [[ -z "${secret}" ]]; then
    echo "Microsoft Graph client secret ausente no ambiente efemero e no 1Password." >&2
    return 1
  fi
  export SP_CLIENT_SECRET="${secret}"
  export MICROSOFT_GRAPH_CLIENT_SECRET="${secret}"
  unset secret
}
