# AGENTS.md -- mcp-onedrive-sharepoint

As regras operacionais deste repo sao canonicas em [CLAUDE.md](CLAUDE.md) (fonte unica para Claude/Codex/Hermes). Leia-o antes de tocar em codigo.

TL;DR das invariantes:
- 1Password-only -- `cpz::SP_CLIENT_ID`, `cpz::SP_CLIENT_SECRET` e `cpz::SP_TENANT_ID` sao resolvidos em runtime; nao adicionar `.env`, Keychain, arquivo ou token delegado como fallback
- Perfil `core` e o default -- ferramentas destrutivas ficam atras de `MCP_TOOL_PROFILE=full`; nao mudar sem intencao clara
- Uso on-demand, nao permanente -- preferir execucao one-shot via `spcall`/`mcporter --stdio`; nao manter o MCP bound permanentemente
- `client-credentials` exige tenant UUID real (nao `common`) e Application permissions com admin consent no Azure AD

Validar: `npm run ci` (build + lint + tests)
