#!/usr/bin/env tsx

import { fileURLToPath } from "node:url";

export async function setupAuthentication(): Promise<void> {
  throw new Error(
    "setup-auth is disabled: this tool does not persist delegated tokens. Set MICROSOFT_GRAPH_TENANT_ID, MICROSOFT_GRAPH_CLIENT_ID and MICROSOFT_GRAPH_CLIENT_SECRET for an app registration with client credentials (see README > Configuration).",
  );
}

const isDirectRun = process.argv[1] === fileURLToPath(import.meta.url);

if (isDirectRun) {
  setupAuthentication().catch((error) => {
    process.stderr.write(
      `Authentication setup unavailable: ${error instanceof Error ? error.message : String(error)}\n`,
    );
    process.exit(1);
  });
}
