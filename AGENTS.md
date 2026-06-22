# AGENTS.md -- mcp-onedrive-sharepoint

As regras operacionais deste repo sao canonicas em [CLAUDE.md](CLAUDE.md) (fonte unica para Claude/Codex/Hermes). Leia-o antes de tocar em codigo.

TL;DR das invariantes:
- Credenciais nunca em git -- `.env`, `config/sites.local.json`, `tokens.json` e Keychain sao gitignored; `siteId`/`driveId`/URLs de tenant ficam fora de commits e README
- Perfil `core` e o default -- ferramentas destrutivas ficam atras de `MCP_TOOL_PROFILE=full`; nao mudar sem intencao clara
- Uso on-demand, nao permanente -- preferir execucao one-shot via `spcall`/`mcporter --stdio`; nao manter o MCP bound permanentemente
- `client-credentials` exige tenant UUID real (nao `common`) e Application permissions com admin consent no Azure AD
- Token MSAL (device-code) expira em 90 dias de inatividade -- automatizacoes devem usar client-credentials

Validar: `npm run ci` (build + lint + tests)
