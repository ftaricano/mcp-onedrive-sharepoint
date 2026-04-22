# MCP OneDrive/SharePoint Server

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%E2%89%A518-brightgreen.svg)](https://nodejs.org)
[![MCP](https://img.shields.io/badge/MCP-compatible-8A2BE2.svg)](https://modelcontextprotocol.io)
[![TypeScript](https://img.shields.io/badge/typescript-%5E5.3-3178c6.svg)](https://www.typescriptlang.org)

MCP server for Microsoft Graph focused on OneDrive, SharePoint and related document workflows. Delegated device-code auth, 32 tools, also usable as a standalone `ods` CLI for shell scripting.

Onboarding commands on a clean clone:

- `npm run build`
- `npm run lint`
- `npm test`
- `npm run ci`
- `npm run setup-auth`

## What is implemented

The server exposes 32 MCP tools grouped into:
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
