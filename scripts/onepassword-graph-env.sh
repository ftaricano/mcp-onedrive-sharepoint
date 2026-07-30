#!/bin/bash -p
# Lib sourceable — resolve as credenciais Microsoft Graph via 1Password
# (política 1P-only, JAR-424). Usa `cpz_keychain.get(key)` — API presente na
# main do hub scripts — com interpretador fixo e modo isolado (-I), no mesmo
# padrão de infra/op-secrets/op_get.sh.

load_graph_environment() {
  local helper_dir exports
  helper_dir="${JARVIS_HUB:-${HOME}/jarvis-hub}/scripts/infra"

  if ! exports="$(
    /usr/bin/python3 -I -c '
import shlex
import sys

sys.path.insert(0, sys.argv[1])
import os
from cpz_keychain import get

items = {
    "MICROSOFT_GRAPH_CLIENT_ID": "SP_CLIENT_ID",
    "MICROSOFT_GRAPH_CLIENT_SECRET": "SP_CLIENT_SECRET",
    "MICROSOFT_GRAPH_TENANT_ID": "SP_TENANT_ID",
}

# CPZ_KEYCHAIN_DEBUG imprime em stdout; nada além dos exports pode chegar ao eval.
real_stdout = sys.stdout
sys.stdout = sys.stderr
values = {}
for name, key in items.items():
    from_env = os.environ.get(name)
    if from_env:
        print(
            f"[onepassword-graph-env] {name} veio do ambiente do processo, "
            "não do 1Password (reservado ao launcher e aos testes)",
            file=sys.stderr,
        )
        values[name] = from_env
    else:
        values[name] = get(key)
sys.stdout = real_stdout

missing = [name for name, value in values.items() if not value]
if missing:
    raise SystemExit("1Password did not resolve " + ", ".join(missing))
for name, value in values.items():
    print(f"export {name}={shlex.quote(value)}")
' "$helper_dir"
  )"; then
    echo "1Password não resolveu as credenciais Microsoft Graph (itens cpz::SP_*). Diagnóstico no stderr acima." >&2
    return 1
  fi

  eval "$exports"
  export SP_CLIENT_SECRET="$MICROSOFT_GRAPH_CLIENT_SECRET"
  unset exports
}
