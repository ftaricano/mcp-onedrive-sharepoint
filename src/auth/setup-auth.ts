#!/usr/bin/env tsx

import { fileURLToPath } from "node:url";

export async function setupAuthentication(): Promise<void> {
  throw new Error(
    "setup-auth is disabled: delegated tokens cannot be persisted outside 1Password. Ask a 1Password owner to provision cpz::SP_CLIENT_ID, cpz::SP_CLIENT_SECRET, and cpz::SP_TENANT_ID.",
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
