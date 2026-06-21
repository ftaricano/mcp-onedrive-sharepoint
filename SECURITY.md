# Security Policy

This project handles Microsoft Graph OAuth tokens and access to OneDrive and SharePoint content.

## Reporting a Vulnerability

Please report security issues privately through GitHub Security Advisories:

https://github.com/ftaricano/mcp-onedrive-sharepoint/security/advisories/new

Do not open a public issue for credential leaks, authorization bypasses, token handling bugs, or tenant data exposure.

## Credentials and Tenant Data

Never commit:

- `.env` or any `.env.*` file other than `.env.example`
- Microsoft Graph client secrets
- OAuth tokens, MSAL cache files, or Keychain exports
- tenant-specific `siteId`, `driveId`, internal SharePoint URLs, or private operational paths
- populated `config/sites.local.json`

Use `config/sites.example.json` as the public template and keep real site registries outside git.

## Tool Surface

The default `MCP_TOOL_PROFILE=core` exposes a smaller set of common tools. Set `MCP_TOOL_PROFILE=full` only for trusted environments that need advanced, destructive, or mutating operations.

The raw Graph `batch_operations` tool is disabled unless `MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH=true` is set.

## If a Token Leaks

Revoke the token in Microsoft Entra ID / Azure AD, rotate any affected client secret, remove the local cache, and re-run authentication.
