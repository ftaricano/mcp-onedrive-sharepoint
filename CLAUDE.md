# CLAUDE.md -- mcp-onedrive-sharepoint

MCP server e CLI (`ods`) para acesso a OneDrive e SharePoint via Microsoft Graph. Serve automacoes do ecossistema do Ferd e pode ser usado em modo stdio (MCP) ou como CLI avulso em scripts shell.

## O que e

Servidor MCP + CLI unificado para operacoes em OneDrive e SharePoint usando a Microsoft Graph API. Suporta dois modos de autenticacao (device-code interativo e client-credentials para automacoes nao-interativas) e dois perfis de ferramentas (`core` para uso diario, `full` para automacoes destrutivas/avancadas). Consumido via Claude Code MCP ou via `ods` / `spcall.sh` em scripts locais.

## Stack & estrutura

Node.js 18+ + TypeScript 5.3 + MSAL Node + MCP SDK 1.29 + keytar (Keychain); testes com `node --test` nativo (sem Jest/Vitest).

```
mcp-onedrive-sharepoint/
├── src/
│   ├── index.ts              # entry MCP server
│   ├── cli.ts / cli/         # entry CLI (ods)
│   ├── auth/                 # MSAL device-code + client-credentials + setup-auth
│   ├── graph/                # Graph HTTP client + error handler
│   ├── sharepoint/           # site resolver
│   ├── config/               # carregamento de env + sites registry
│   ├── core/                 # bootstrap de ferramentas
│   ├── tools/
│   │   ├── files/            # list, download, upload, move, delete, share, copy, search, metadata
│   │   ├── sharepoint/       # sites, lists, list items
│   │   ├── advanced/         # analytics, collaboration, excel, sync
│   │   ├── registry.ts       # registro de ferramentas por perfil
│   │   └── utils/            # path-helper
│   ├── utils/                # local-path + helpers
│   └── tests/                # testes (.test.ts -> build/tests/*.test.js)
├── scripts/
│   ├── run-stdio.sh          # inicia MCP stdio carregando .env local
│   ├── spcall.sh             # chamada ad-hoc via mcporter contra servidor local
│   ├── ods.sh                # wrapper shell do CLI ods
│   └── exec-with-env.mjs     # loader de .env para scripts
├── config/
│   ├── sites.example.json    # template do registry de sites
│   └── sites.local.json      # (gitignored) aliases de sites com siteId/driveId reais
├── .env.example              # template de variaveis de ambiente
├── .env                      # (gitignored) credenciais locais reais
└── tsconfig.json
```

## Como rodar / validar

```bash
# Setup inicial
npm install
cp .env.example .env   # preencher CLIENT_ID, TENANT_ID e opcionalmente CLIENT_SECRET

# Build
npm run build

# Auth (device-code -- so se nao usar client-credentials)
npm run setup-auth

# Fumar o servidor em modo stdio
./scripts/spcall.sh health_check
./scripts/spcall.sh list_drives

# CLI direto
ods list
ods health_check

# Suite completa (o que o CI roda)
npm run ci              # build + lint + tests

# So testes
npm test

# So lint
npm run lint
```

## Invariantes / regras criticas

- **Credenciais nunca em git**: `.env`, `tokens.json`, `credentials.json`, `config/sites.local.json` e entradas do Keychain (service `mcp-onedrive-sharepoint`) sao sempre gitignored. Nenhum `siteId`, `driveId`, URL de tenant ou secret vai para commits ou README.
- **Perfil `core` e o default**: nao mudar `MCP_TOOL_PROFILE` para `full` em config persistente sem intencao clara. Ferramentas destrutivas (`delete_item`, `manage_permissions`) ficam atras do perfil `full`.
- **Uso on-demand, nao permanente**: nao manter este MCP bound/loaded permanentemente no Hermes ou Claude Code. Preferir execucao one-shot via `spcall` / `mcporter --stdio` para que o processo encerre apos a chamada e nao acumule processos zumbi.
- **`batch_operations` e experimental**: nao ativar `MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH=true` em uso normal; e um escape-hatch de Graph batch crua para debug/admin.
- **`client-credentials` exige tenant UUID especifico**: `MICROSOFT_GRAPH_TENANT_ID=common` nao funciona com esse flow; deve ser um UUID real do tenant. Permissoes do tipo Application (nao Delegated) com admin consent no Azure AD.
- **Token MSAL expira em 90 dias de inatividade** (device-code): se o servidor nao for usado por varios dias, re-executar `npm run setup-auth`. Para automacoes, preferir client-credentials (sem expiry).
- **`npm run ci` e o gate de validacao**: qualquer mudanca de codigo deve passar `build + lint + tests` antes de ser considerada pronta.

## Gotchas

- Testes sao compilados antes de rodar (`npm test` faz `build` primeiro, entao `node --test build/tests/*.test.js`). Editar `.test.ts` sem buildar nao reflete nos testes executados.
- `MICROSOFT_GRAPH_TENANT_ID=common` funciona para device-code mas quebra silenciosamente client-credentials com erro `AADSTS700016`.
- Site aliases exigem `config/sites.local.json` (gitignored). Se o arquivo nao existir, ferramentas com `site=<alias>` falham; ferramentas com `siteId`/`driveId` explicitos continuam funcionando.
- O MCP stdio deve ser iniciado via `./scripts/run-stdio.sh` (nao `node build/index.js` diretamente) para garantir que o `.env` local seja carregado.
- `keytar` requer acesso ao Keychain do macOS; em ambientes de CI/sem Keychain, usar client-credentials via variaveis de ambiente diretamente.

## Documentacao canonica

- Skill: n/a (sem skill dedicada no hub ainda)
- Tracking: ver time JAR no Linear para issues relacionadas a automacoes OneDrive/SharePoint
