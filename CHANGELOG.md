# Changelog

All notable changes to this project are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Security

- Mask signed download URLs, authentication query values and authorization headers
  in tool responses, CLI output and diagnostics; retain URL host and path only.
  Signed URLs in output are diagnostic values and must not be reused for requests.
- Apply the same token field pattern to objects and embedded JSON diagnostics,
  including `refresh_token` and `id_token`, while preserving Graph error codes.
- Redact sensitive fields inside escaped JSON string literals in diagnostics,
  errors and audit logs, including multiple serialization layers. Bound excessive
  nesting with a masked value rather than returning uninspected text.
- Preserve the association between authorization header names and quoted values
  while masking diagnostics, without consuming adjacent JSON fields.
- Unify sensitive field masking across objects, assignments, quoted values,
  headers and escaped diagnostics, including passwords and client secrets.
  Mask standalone JWT-shaped values and all SharePoint query strings, including
  unknown signing parameter names, while retaining URL host/path. Ordinary Graph
  error codes are preserved via an explicit public diagnostic-code allowlist;
  unknown code values and all code query/equals assignments are masked.

- Authorization, Proxy-Authorization and WWW-Authenticate output preserves only
  the scheme plus `***`, masking every parameter through end of line or the full
  quoted/JSON value. This includes Digest, AWS4-HMAC-SHA256 and unknown schemes;
  external JSON metadata remains separate from the masked header value.
- Standalone authentication schemes share one credential grammar for token68,
  quoted values and parameter lists, including unknown scheme names and escaped
  diagnostics. The scheme stays visible while its credential becomes `***`.

### Added

- Upload and download results expose `file.eTag` and `file.cTag` (null when absent).
  Downloads read metadata first and send `If-Match` when an eTag is available.
  This conditional request applies to download only. `upload_file` exposes the
  resulting tags but does not accept an expected eTag or enforce conditional upload;
  conditional upload remains a separate follow-up.
  Version tags identify the remote item; they are not a local content checksum.

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
