# MCP OneDrive/SharePoint Server

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%E2%89%A518-brightgreen.svg)](https://nodejs.org)
[![MCP](https://img.shields.io/badge/MCP-compatible-8A2BE2.svg)](https://modelcontextprotocol.io)
[![TypeScript](https://img.shields.io/badge/typescript-%5E5.3-3178c6.svg)](https://www.typescriptlang.org)

MCP server for Microsoft Graph focused on OneDrive, SharePoint and related document workflows. Delegated device-code auth, 33 tools, also usable as a standalone `ods` CLI for shell scripting.

Onboarding commands on a clean clone:

- `npm run build`
- `npm run lint`
- `npm test`
- `npm run ci`
- `npm run setup-auth`

## What is implemented

The server exposes 33 MCP tools grouped into:

- Files: `list_files`, `download_file`, `upload_file`, `create_folder`, `move_item`, `delete_item`, `search_files`, `get_file_metadata`, `share_item`, `copy_item`
- SharePoint: `discover_sites`, `list_site_lists`, `get_list_schema`, `list_items`, `get_list_item`, `create_list_item`, `update_list_item`, `delete_list_item`
- Utilities: `health_check`, `get_user_profile`, `list_drives`, `global_search`, `batch_operations`
- Advanced: `advanced_share`, `manage_permissions`, `check_user_access`, `sync_folder`, `batch_file_operations`, `storage_analytics`, `version_management`, `excel_operations`, `excel_analysis`

## Recent foundation improvements

This version includes real structural improvements instead of documentation-only changes:

- fixed `npm run setup-auth` to call the real TypeScript auth setup flow
- added working ESLint configuration
- added executable tests with Node's built-in test runner
- aligned README and `.env.example` with the real environment variables and scripts
- introduced reusable Graph helpers for:
  - consistent MCP JSON/error envelopes
  - pagination extraction from Microsoft Graph responses
  - resolving OneDrive/SharePoint resources by `driveId`, `siteId`, `itemId` and path
- updated key listing/search flows to return pagination metadata and support `driveId`

## Requirements

- Node.js 18+
- A Microsoft Entra ID / Azure app registration for delegated device-code authentication

## Installation

```bash
git clone https://github.com/ftaricano/mcp-onedrive-sharepoint.git
cd mcp-onedrive-sharepoint
npm install
cp .env.example .env
```

## Operational wrappers

Important operational rule:
- use this MCP on demand
- do not keep it permanently bound/loaded in Hermes or Claude Code when not needed
- prefer one-shot `spcall` / `mcporter --stdio` execution so the process exits right after the call and does not accumulate zombie or idle MCP processes
- the `spcall` wrapper includes post-call cleanup for stray repo-local MCP child processes


This repo includes lightweight wrappers for local operational use:

- `./scripts/run-stdio.sh`: start the MCP stdio server with the repo-local `.env` loaded safely
- `./scripts/spcall.sh`: run ad-hoc `mcporter` calls against the local MCP server
- `npm run stdio`: same as `./scripts/run-stdio.sh`
- `npm run spcall -- <tool> ...`: same as `./scripts/spcall.sh <tool> ...`

Quick examples:

```bash
npm run build
./scripts/spcall.sh health_check
./scripts/spcall.sh list_drives
./scripts/spcall.sh list_files driveId=b!abc123 path=/Shared%20Documents
```

Tenant-specific site aliases and drive ids are loaded from a local file — see [Site registry](#site-registry) below.

## CLI (`ods`)

Every MCP tool is also exposed as a plain subcommand through the `ods` CLI. It shares the same auth, config and handlers as the MCP server, so anything the MCP does is one-shot runnable from a terminal or a shell script.

```bash
npm run build
ods list                                  # list all tools with descriptions
ods schema list_files                     # print JSON schema for a tool
ods auth                                  # interactive device-code login
ods <tool> --key=value [--key value]      # invoke a tool with CLI flags
ods <tool> --json '{"k":"v"}'             # pass the full payload as JSON
```

During development you can skip the build with `npm run cli -- <tool> ...`.

### Examples

```bash
ods health_check
ods list_files --site=primary --path=/
ods list_files --driveId=b!abc --path=/Shared%20Documents --limit=50
ods upload_file --json '{"driveId":"b!abc","path":"/x.txt","content":"hello"}'
```

### Rules for flags

- `--key=value` and `--key value` are both accepted.
- `true` / `false` / `null` and numeric strings are coerced automatically; anything else stays a string.
- Bare flags (no value, or followed by another flag) become `true`.
- `--json '<payload>'` takes a JSON object; individual `--key=value` flags layered on top override fields from the payload. Use this for tools with nested objects/arrays (e.g. advanced Excel tools).
- Output is the raw tool payload (usually pretty-printed JSON). If the handler returns an error envelope, the process exits with code `2`.

## Configuration

The server reads the following environment variables:

```bash
MICROSOFT_GRAPH_CLIENT_ID=your_app_client_id
MICROSOFT_GRAPH_TENANT_ID=common
MICROSOFT_GRAPH_SCOPES=Files.ReadWrite.All,Sites.ReadWrite.All,Directory.Read.All,User.Read,offline_access
MICROSOFT_GRAPH_BASE_URL=https://graph.microsoft.com/v1.0
MICROSOFT_GRAPH_TIMEOUT=30000
MICROSOFT_GRAPH_MAX_RETRIES=3
MICROSOFT_GRAPH_CACHE_ENABLED=true
MICROSOFT_GRAPH_CACHE_TTL=3600
```

Notes:

- use `MICROSOFT_GRAPH_TENANT_ID=common` for multi-tenant/device-code onboarding
- use a specific tenant id if you want tenant-scoped sign-in
- delegated scopes are what the current auth flow uses

## Authentication setup

Run:

```bash
npm run setup-auth
```

The script:

- reads `MICROSOFT_GRAPH_CLIENT_ID` / `MICROSOFT_GRAPH_TENANT_ID` from `.env` when present
- prompts for missing values
- starts Microsoft device-code login
- stores the token through the existing auth layer

## Development commands

```bash
npm run build
npm run lint
npm test
npm run ci
npm start
npm run stdio
npm run spcall -- health_check
```

`npm run ci` is the local verification entrypoint and is also what GitHub Actions runs on every PR/push.

## MCP behavior notes

### Root site inclusion

`discover_sites.includePersonalSite=true` currently attempts to append the tenant root SharePoint site (`/sites/root`) when it is available to the authenticated user.
It does not discover or synthesize a personal OneDrive site.

### Pagination

The following tools now expose consistent pagination metadata in their JSON payloads:

- `list_files`
- `search_files`
- `discover_sites`
- `list_site_lists`
- `list_items`

When Microsoft Graph returns `@odata.nextLink`, the response includes:

- `pagination.returned`
- `pagination.limit`
- `pagination.totalCount` when available
- `pagination.nextPageToken`
- `pagination.hasMore`

Pass `pageToken` back to the same tool to continue paging.

### Drive/site targeting

Core file listing/search/download flows now accept:

- `siteId` for a SharePoint site's default drive
- `driveId` for a specific document library or drive
- path-based addressing where supported

This is the current foundation for moving beyond a `/me/drive`-only model.

## Site registry

The resolver can target named SharePoint sites by alias (e.g. `site=primary`). The registry is loaded from an external JSON file so no tenant-specific ids are committed:

- Copy `config/sites.example.json` to `config/sites.local.json` (gitignored) and fill in your values.
- Or set `MCP_SITES_CONFIG_PATH` to point at a different JSON file.
- If the file is missing, the registry stays empty and the tools only accept explicit `siteId`, `siteUrl`, or `driveId`.

Each site entry looks like:

```json
{
  "key": "primary",
  "name": "Primary",
  "siteId": "yourtenant.sharepoint.com,<guid>,<guid>",
  "siteUrl": "https://yourtenant.sharepoint.com/sites/Primary",
  "driveId": "b!<drive-id>",
  "aliases": ["primary", "/sites/Primary"]
}
```

### MCP stdio snippet

Use the wrapper as the MCP command so the repo-local `.env` is loaded automatically:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "/absolute/path/to/mcp-onedrive-sharepoint/scripts/run-stdio.sh"
    }
  }
}
```

## Example tool inputs

### List files from a specific drive

```json
{
  "driveId": "b!abc123",
  "path": "/Shared Documents",
  "limit": 50
}
```

### Continue a paginated file listing

```json
{
  "driveId": "b!abc123",
  "pageToken": "https://graph.microsoft.com/v1.0/drives/b!abc123/root/children?$skiptoken=..."
}
```

### Search files in a site drive

```json
{
  "siteId": "contoso.sharepoint.com,123,456",
  "query": "quarterly report",
  "limit": 25
}
```

### List SharePoint list items with pagination

```json
{
  "siteId": "contoso.sharepoint.com,123,456",
  "listId": "9c6b8b70-0000-0000-0000-111111111111",
  "orderBy": "Created desc",
  "limit": 100
}
```

## Troubleshooting

- `invalid_grant` / `AADSTS` on first run: token store is empty or expired. Run `npm run setup-auth` again.
- `403 Forbidden` on SharePoint lists/drives: the signed-in user lacks permission to the target site. Check with the site owner.
- `404` on a `driveId` or `siteId`: the identifier is stale or the resource was deleted. Use `list_drives` / `discover_sites` to re-discover.
- Build fails on a clean clone: make sure Node.js is 18+ and run `npm install` before `npm run build`.

## Security

This server handles Microsoft Graph OAuth tokens and delegated access to corporate file storage. Treat it accordingly:

- `.env`, `tokens.json`, `credentials.json` and the OS keychain entries are **never** committed — see [.gitignore](.gitignore).
- Report security issues privately via [GitHub security advisories](https://github.com/ftaricano/mcp-onedrive-sharepoint/security/advisories/new) — do not open a public issue.
- If a token leaks, revoke it from [Azure AD → Enterprise Applications → your app → Users & groups](https://portal.azure.com) and re-run `npm run setup-auth`.

## Contributing

Issues and PRs welcome. Before opening a PR:

- `npm run ci` passes (build + lint + tests)
- one focused change per PR
- no credentials, tenant-specific ids, or internal paths in commits or README

## License

[MIT](LICENSE) © Fernando Taricano

## Current limitations

- authentication depends on real Microsoft Graph credentials and an interactive device-code login
- only the most critical listing/search flows use the pagination/resource helpers; other tools still use direct endpoint construction
