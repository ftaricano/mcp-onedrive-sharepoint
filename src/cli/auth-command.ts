export async function runAuthSetup(): Promise<void> {
  throw new Error(
    "ods auth is disabled: it would need to persist a delegated token. Set MICROSOFT_GRAPH_TENANT_ID, MICROSOFT_GRAPH_CLIENT_ID and MICROSOFT_GRAPH_CLIENT_SECRET for an app registration with client credentials (see README > Configuration).",
  );
}
