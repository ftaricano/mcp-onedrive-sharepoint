# Contributing

Thanks for taking the time to contribute. Bug reports, fixes and documentation
improvements are all welcome.

## Before you start

- For anything bigger than a small fix, open an issue first so we can agree on the
  approach.
- Security problems go through [SECURITY.md](SECURITY.md), not public issues.

## Development setup

The commands, layout and conventions are in [AGENTS.md](AGENTS.md). In short:

```bash
npm ci
npm run ci    # build + lint + tests
```

The unit tests do not call Microsoft Graph and need no credentials.

## Pull requests

1. Fork the repository and create a branch from `main`.
2. Keep the change focused; add or update tests for behavior changes.
3. Add a line under `## [Unreleased]` in [CHANGELOG.md](CHANGELOG.md) for user-visible
   changes.
4. Use [Conventional Commits](https://www.conventionalcommits.org/) for commit messages
   (`feat:`, `fix:`, `docs:`, `chore:`).
5. Make sure CI is green. Examples and fixtures must use fictional data (Acme,
   `contoso.sharepoint.com`), never real names, addresses, ids or credentials.

By contributing you agree that your contribution is licensed under the project's
[MIT License](LICENSE).
