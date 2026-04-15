# MCP OneDrive/SharePoint Server

MCP server for Microsoft Graph focused on OneDrive, SharePoint and related document workflows.

This repository now has working onboarding and quality commands on a clean clone:
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
git clone <repository-url>
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

## Quality status

Validated locally after these changes:
- `npm run build` ✅
- `npm run lint` ✅
- `npm test` ✅

## Current limitations

- authentication still depends on real Microsoft Graph credentials and an interactive device-code login
- only the most critical listing/search foundations were migrated to the new pagination/resource helpers in this pass; other tools still use older direct endpoint construction
- advanced tools exist, but this PR focuses on onboarding, contract clarity and foundation hardening rather than a full architectural rewrite
