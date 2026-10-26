# Changelog

All notable changes to this project are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- `upload_file` / `UploadOptions.retryPolicy`: optional `default|single-attempt`
  policy shared by CLI and MCP. `single-attempt` accepts files smaller than 4 MiB
  in an existing destination folder, sends at most one simple PUT with no
  redirects, and returns errors without retry, folder creation or upload sessions.
  Invalid policies and sizes at or above 4 MiB fail before any Graph write.
  Errors may follow an applied write; no rollback, CAS or exactly-once guarantee.
  Omitted/default policy retains existing behavior, including large uploads.

### Changed

- Upgrade MSAL Node to 7.0.0 and require Node.js 20 or newer.
- Documentation: Quickstart with plain environment variables, a tools table,
  English `AGENTS.md`, `CONTRIBUTING.md`, and issue and pull request templates.
- Credential error messages, the `ods auth` / `setup-auth` notices and the
  `health_check` `authMethod` field no longer refer to a specific secret manager;
  they name the `MICROSOFT_GRAPH_*` environment variables instead.

## [1.0.1] - 2026-09-25

### Security

- Validate the destination of every authenticated Microsoft Graph request.
  The client now sends a request (and fetches or attaches the access token)
  only when the effective URL is `https://graph.microsoft.com/v1.0/...` or
  `https://graph.microsoft.com/beta/...` on the default port, without
  userinfo. Absolute `pageToken` values in `list_files`, `search_files`,
  `discover_sites`, `list_site_lists` and `list_items`, and `@odata.nextLink`
  values followed by the client, that point anywhere else are rejected with a
  validation error. Graph nextLinks and relative tokens keep working as before.

[Unreleased]: https://github.com/ftaricano/mcp-onedrive-sharepoint/compare/v1.0.1...HEAD
[1.0.1]: https://github.com/ftaricano/mcp-onedrive-sharepoint/releases/tag/v1.0.1
