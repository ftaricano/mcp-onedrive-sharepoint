# CPZ operational SharePoint usage

This repository includes lightweight wrappers for day-to-day CPZ SharePoint operations:

- `./scripts/run-stdio.sh`: launches the MCP stdio server from this repo with the repo-local `.env` parsed safely
- `./scripts/spcall.sh`: runs ad-hoc `mcporter call` requests against the local MCP server
- `npm run stdio`: same as `./scripts/run-stdio.sh`
- `npm run spcall -- <tool> ...`: same as `./scripts/spcall.sh <tool> ...`

The wrappers do not `source` `.env` in the shell. They parse the file with `dotenv`, merge it into the process environment, and keep explicit environment variables higher priority.

## Canonical site aliases

Use these aliases consistently in prompts, wrapper examples, and future skills/docs.

### Financeiro / Operacional

Recommended aliases:

- `financeiro-operacional`
- `financeiro`
- `operacional`
- `finops`

Canonical references:

- `siteId`: `cpzseg.sharepoint.com,ba154fbe-85ac-4cb2-869a-427e4d0d251d,f096663c-dc94-48a4-9e2d-982a3c61e01b`
- `siteUrl`: `https://cpzseg.sharepoint.com/sites/FinanceiroOperacional`
- `driveId`: `b!vk8VuqyFskyGmkJ-TQ0lHTxmlvCU3KRIni2YKjxh4Bs2mBj0spMxRbYZtGTl-cmn`

### Socios

Recommended aliases:

- `socios`
- `socios2`

Canonical references:

- `siteId`: `cpzseg.sharepoint.com,259c3398-2f44-4c48-8558-39515feb5b4e,f67178fe-fd20-42dc-bd79-d68f3d25a4f3`
- `siteUrl`: `https://cpzseg.sharepoint.com/sites/socios2`
- `driveId`: `b!mDOcJUQvSEyFWDlRX-tbTv54cfYg_dxCvXnWjz0lpPOpFZGqnAuOS5orKS0nDH1R`

## Important location aliases

These are practical human aliases for recurring operational paths.

### Financeiro / Operacional hot paths

Base drive: `b!vk8VuqyFskyGmkJ-TQ0lHTxmlvCU3KRIni2YKjxh4Bs2mBj0spMxRbYZtGTl-cmn`

- `fin.comissoes-finais` -> `/comissoes/planilhas-finais`
- `fin.comissoes-master` -> `/comissoes/master`
- `fin.producao-finais` -> `/producao/planilhas-finais`
- `fin.comissoes-base` -> `/comissoes/base`

### Socios hot paths

Base drive: `b!mDOcJUQvSEyFWDlRX-tbTv54cfYg_dxCvXnWjz0lpPOpFZGqnAuOS5orKS0nDH1R`

- `socios.competenza-logos-png` -> `/COMPETENZA/marketing/logos/PNG`
  - Fundo Transparente
- `socios.competenza-logos-jpeg` -> `/COMPETENZA/marketing/logos/JPEG`
  - Fundo branco
- `socios.competenza-logos-eps` -> `/COMPETENZA/marketing/logos/EPS`
  - Arquivo em vetor, para designers e gráficas
- `socios.competenza-templates` -> `/COMPETENZA/marketing/material-grafico/templates`
- `socios.interpar-dashboard` -> `/INTERPAR/financeiro/dashboard-financeiro`
- `socios.interpar-dados-base` -> `/INTERPAR/financeiro/02_dados_base`

## Process lifecycle rule

This MCP is operationally on-demand only.

- Do not keep it registered as a permanently loaded MCP in Hermes or Claude Code.
- Prefer `spcall` or ad-hoc `mcporter call --stdio ...`.
- Each call should start the MCP process, use it for that request, and let it exit immediately afterward.
- `spcall` also runs a cleanup trap after each invocation to kill leftover repo-local MCP processes if any child fails to exit cleanly.
- The goal is zero long-lived idle MCP processes and no zombie leftovers after routine usage.

In practice, `spcall` is the canonical path because it launches an ephemeral stdio MCP process through `mcporter`, returns the result, and exits.

## Wrapper usage

### Start the MCP server over stdio

```bash
npm run build
./scripts/run-stdio.sh
```

Equivalent:

```bash
npm run stdio
```

This is the wrapper to use in MCP client configs because it always launches from the repo root and loads the repo-local `.env` safely.

### Ad-hoc tool calls with `spcall`

```bash
npm run build
./scripts/spcall.sh health_check
./scripts/spcall.sh list_drives
./scripts/spcall.sh list_files driveId=b!vk8VuqyFskyGmkJ-TQ0lHTxmlvCU3KRIni2YKjxh4Bs2mBj0spMxRbYZtGTl-cmn path=/comissoes/master
```

Equivalent npm form:

```bash
npm run spcall -- health_check
npm run spcall -- list_files driveId=b!mDOcJUQvSEyFWDlRX-tbTv54cfYg_dxCvXnWjz0lpPOpFZGqnAuOS5orKS0nDH1R path=/COMPETENZA/marketing/logos/PNG
```

Behavior notes:

- defaults to `--output json` unless you pass your own `--output ...`
- uses `mcporter call --stdio` so the MCP process is ephemeral
- passes `--cwd` as the repo root so local build artifacts and `.env` resolution stay consistent

## CPZ operational examples

### Financeiro / Operacional examples

```bash
./scripts/spcall.sh list_files \
  driveId=b!vk8VuqyFskyGmkJ-TQ0lHTxmlvCU3KRIni2YKjxh4Bs2mBj0spMxRbYZtGTl-cmn \
  path=/comissoes/planilhas-finais

./scripts/spcall.sh list_files \
  driveId=b!vk8VuqyFskyGmkJ-TQ0lHTxmlvCU3KRIni2YKjxh4Bs2mBj0spMxRbYZtGTl-cmn \
  path=/producao/planilhas-finais
```

### Socios examples

```bash
./scripts/spcall.sh list_files \
  driveId=b!mDOcJUQvSEyFWDlRX-tbTv54cfYg_dxCvXnWjz0lpPOpFZGqnAuOS5orKS0nDH1R \
  path=/COMPETENZA/marketing/logos/PNG

./scripts/spcall.sh list_files \
  driveId=b!mDOcJUQvSEyFWDlRX-tbTv54cfYg_dxCvXnWjz0lpPOpFZGqnAuOS5orKS0nDH1R \
  path=/INTERPAR/financeiro/dashboard-financeiro
```

### Site-based examples

```bash
./scripts/spcall.sh list_drives siteId=cpzseg.sharepoint.com,ba154fbe-85ac-4cb2-869a-427e4d0d251d,f096663c-dc94-48a4-9e2d-982a3c61e01b

./scripts/spcall.sh list_site_lists siteId=cpzseg.sharepoint.com,259c3398-2f44-4c48-8558-39515feb5b4e,f67178fe-fd20-42dc-bd79-d68f3d25a4f3
```

## Claude / Hermes / native MCP integration snippet

Use the stdio wrapper as the MCP command so the repo-local environment is loaded automatically.

```json
{
  "mcpServers": {
    "cpz-sharepoint": {
      "command": "/Users/jarvis/jarvis-hub/repos/tools/mcp-onedrive-sharepoint/scripts/run-stdio.sh"
    }
  }
}
```

For skills or agent-side ad-hoc execution, prefer `spcall` when you want a one-off tool call without registering a persistent MCP server.
