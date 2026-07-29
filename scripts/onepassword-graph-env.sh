#!/usr/bin/env bash

load_graph_environment() {
  local hub output
  hub="${JARVIS_HUB:-${HOME}/jarvis-hub}"

  if ! output="$(
    PYTHONPATH="${hub}/scripts/infra${PYTHONPATH:+:${PYTHONPATH}}" \
      python3 -c '
import json
from cpz_keychain import get_item

items = {
    "MICROSOFT_GRAPH_CLIENT_ID": "SP_CLIENT_ID",
    "MICROSOFT_GRAPH_CLIENT_SECRET": "SP_CLIENT_SECRET",
    "MICROSOFT_GRAPH_TENANT_ID": "SP_TENANT_ID",
}
values = {name: get_item(item) for name, item in items.items()}
missing = [name for name, value in values.items() if not value]
if missing:
    raise SystemExit("1Password did not resolve " + ", ".join(missing))
print(json.dumps(values))
' 2>&1
  )"; then
    echo "1Password não resolveu as credenciais Microsoft Graph. Verifique os itens cpz::SP_* com o owner." >&2
    return 1
  fi

  eval "$(python3 -c '
import json
import shlex
import sys

for key, value in json.loads(sys.stdin.read()).items():
    print(f"export {key}={shlex.quote(value)}")
' <<<"$output")"
  export SP_CLIENT_SECRET="$MICROSOFT_GRAPH_CLIENT_SECRET"
  unset output
}
