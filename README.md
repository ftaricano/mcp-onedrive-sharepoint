# MCP OneDrive/SharePoint Server

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%E2%89%A518-brightgreen.svg)](https://nodejs.org)
[![MCP](https://img.shields.io/badge/MCP-compatible-8A2BE2.svg)](https://modelcontextprotocol.io)
[![TypeScript](https://img.shields.io/badge/typescript-%5E5.3-3178c6.svg)](https://www.typescriptlang.org)

MCP server and CLI for Microsoft Graph focused on OneDrive, SharePoint and related document workflows. It uses 1Password-provisioned client credentials only, starts with a safe 10-tool core profile, and can opt into advanced tools for trusted automation.

Onboarding commands on a clean clone:

- `npm run build`
- `npm run lint`
- `npm test`
- `npm run ci`

## Tool profiles

The server defaults to `MCP_TOOL_PROFILE=core`, a smaller public surface intended for day-to-day document workflows:

- `health_check`, `list_drives`
- `discover_sites`, `resolve_site`
- `list_files`, `search_files`, `get_file_metadata`
- `download_file`, `upload_file`, `create_folder`

Set `MCP_TOOL_PROFILE=full` to expose advanced and destructive tools for trusted environments:

- Files: `list_files`, `download_file`, `upload_file`, `create_folder`, `move_item`, `delete_item`, `search_files`, `get_file_metadata`, `share_item`, `copy_item`
- SharePoint: `discover_sites`, `resolve_site`, `list_site_lists`, `get_list_schema`, `list_items`, `get_list_item`, `create_list_item`, `update_list_item`, `delete_list_item`
- Utilities: `health_check`, `get_user_profile`, `list_drives`, `global_search`
- Advanced: `advanced_share`, `manage_permissions`, `check_user_access`, `sync_folder`, `batch_file_operations`, `storage_analytics`, `version_management`, `excel_operations`, `excel_analysis`

`batch_operations` is intentionally not part of either profile by default because it is a raw Microsoft Graph escape hatch. Enable it only for admin/debug workflows with `MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH=true`.

You can also remove individual tools with `MCP_DISABLED_TOOLS=delete_item,manage_permissions`.

## Why this repo

- one MCP server for both OneDrive and SharePoint document libraries
- matching `ods` CLI for shell scripting and one-shot automation
- 1Password-only client credentials for interactive and unattended use
- site aliases loaded from a local registry so tenant IDs stay out of git
- pagination/resource helpers for `driveId`, `siteId`, `itemId` and path targeting

## Requirements

- Node.js 18+
- A Microsoft Entra ID / Azure AD confidential app registration with Application permissions (`Files.ReadWrite.All`, `Sites.ReadWrite.All`) and admin consent. The 1Password owner must provision `cpz::SP_CLIENT_ID`, `cpz::SP_CLIENT_SECRET`, and `cpz::SP_TENANT_ID`; the tenant must be a specific UUID, not `common`.

## Installation

```bash
git clone https://github.com/ftaricano/mcp-onedrive-sharepoint.git
cd mcp-onedrive-sharepoint
npm install
```

## Operational wrappers

Important operational rule:

- use this MCP on demand
- do not keep it permanently bound/loaded in Hermes or Claude Code when not needed
- prefer one-shot `spcall` / `mcporter --stdio` execution so the process exits right after the call and does not accumulate zombie or idle MCP processes
- the `spcall` wrapper includes post-call cleanup for stray repo-local MCP child processes

This repo includes lightweight wrappers for local operational use:

- `./scripts/run-stdio.sh`: start the MCP stdio server after resolving required values from 1Password
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
# `npm install` does NOT put `ods` on your PATH. Link it once, e.g.:
#   npm link            # or: ln -s "$PWD/scripts/ods.sh" ~/bin/ods
ods list                                  # list all tools with descriptions
ods schema list_files                     # print JSON schema for a tool
ods auth                                  # exits: delegated token persistence is intentionally disabled
ods <tool> --key=value [--key value]      # invoke a tool with CLI flags
ods <tool> --json '{"k":"v"}'             # pass the full payload as JSON
```

During development, rebuild before `npm run cli -- <tool> ...`; the command uses
the same packaged 1Password launcher as the installed `ods` bin.

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
# These values are injected only by scripts/with-onepassword-graph-env.sh:
# MICROSOFT_GRAPH_CLIENT_ID
# MICROSOFT_GRAPH_TENANT_ID (specific UUID)
# MICROSOFT_GRAPH_CLIENT_SECRET
MICROSOFT_GRAPH_SCOPES=Files.ReadWrite.All,Sites.ReadWrite.All,Directory.Read.All,User.Read,offline_access
MICROSOFT_GRAPH_BASE_URL=https://graph.microsoft.com/v1.0
MICROSOFT_GRAPH_TIMEOUT=30000
MICROSOFT_GRAPH_MAX_RETRIES=3
MICROSOFT_GRAPH_CACHE_ENABLED=true
MICROSOFT_GRAPH_CACHE_TTL=3600
MCP_TOOL_PROFILE=core
MCP_DISABLED_TOOLS=
MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH=false
```

Notes:

- Graph credentials are resolved from 1Password for every supported npm command and packaged bin
- process-local credential variables are reserved for the launcher and tests; `.env` is never loaded
- set `MCP_LOCAL_FILE_ROOT` to constrain local upload/download/sync file access; if unset, local paths are constrained to the process working directory

## Authentication modes

### 1Password client credentials

Every supported launcher resolves `cpz::SP_CLIENT_ID`, `cpz::SP_CLIENT_SECRET`, and
`cpz::SP_TENANT_ID` through the canonical 1Password helper and injects the values only
into its child process. There is no `.env`, Keychain, file cache, or delegated token
fallback. `npm run setup-auth` and `ods auth` fail intentionally because the service
account cannot persist delegated tokens; request owner-mediated provisioning instead.

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

Use the wrapper as the MCP command so Graph credentials are resolved from 1Password:

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

- `403 Forbidden` on SharePoint lists/drives: the app registration lacks permission to the target site. Check application permissions and admin consent with the owner.
- `404` on a `driveId` or `siteId`: the identifier is stale or the resource was deleted. Use `list_drives` / `discover_sites` to re-discover.
- Build fails on a clean clone: make sure Node.js is 18+ and run `npm install` before `npm run build`.
- `AADSTS700016` or `401`: ensure the 1Password owner has provisioned a specific tenant UUID (not `common`) and Application permissions have admin consent in Azure AD.
- `AADSTS7000215` (invalid client secret): rotate the secret in the app registration and have the 1Password owner update `cpz::SP_CLIENT_SECRET`.

## Security

This server handles Microsoft Graph client credentials and access to corporate file storage. Treat it accordingly:

- `.env`, `tokens.json`, `credentials.json`, and secret-store exports are **never** committed — see [.gitignore](.gitignore).
- tenant-specific `siteId`, `driveId`, SharePoint URLs and internal operational paths should stay in local/private docs, not in this public repo.
- Report security issues privately via [GitHub security advisories](https://github.com/ftaricano/mcp-onedrive-sharepoint/security/advisories/new) — do not open a public issue.
- If a client secret leaks, revoke it in Azure AD and ask the 1Password owner to rotate `cpz::SP_CLIENT_SECRET`.

## Contributing

Issues and PRs welcome. Before opening a PR:

- `npm run ci` passes (build + lint + tests)
- one focused change per PR
- no credentials, tenant-specific ids, or internal paths in commits or README

## License

[MIT](LICENSE) © Fernando Taricano

## Current limitations

- client credentials require Application permissions, admin consent, and owner-mediated provisioning in 1Password
- advanced/destructive tools require `MCP_TOOL_PROFILE=full`
- raw Graph batch calls require `MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH=true`
