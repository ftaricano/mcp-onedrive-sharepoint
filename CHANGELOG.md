# Changelog

All notable changes to this project are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/).

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
