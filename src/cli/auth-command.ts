export async function runAuthSetup(): Promise<void> {
  throw new Error(
    "ods auth is disabled: it would need to persist a delegated token. Ask a 1Password owner to provision cpz::SP_CLIENT_ID, cpz::SP_CLIENT_SECRET, and cpz::SP_TENANT_ID.",
  );
}
