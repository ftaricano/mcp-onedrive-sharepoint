import * as dotenv from "dotenv";
import * as readline from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";
import { initializeAuth } from "../auth/microsoft-graph-auth.js";

dotenv.config();

async function promptForMissingValue(
  rl: readline.Interface,
  label: string,
  currentValue?: string,
  fallback?: string,
): Promise<string> {
  if (currentValue?.trim()) return currentValue.trim();
  const suffix = fallback ? ` (${fallback})` : "";
  const answer = await rl.question(`${label}${suffix}: `);
  return (answer.trim() || fallback || "").trim();
}

export async function runAuthSetup(): Promise<void> {
  const rl = readline.createInterface({ input, output });
  try {
    const clientId = await promptForMissingValue(
      rl,
      "Azure App Client ID",
      process.env.MICROSOFT_GRAPH_CLIENT_ID,
    );
    if (!clientId) throw new Error("Client ID is required");

    const tenantId = await promptForMissingValue(
      rl,
      "Tenant ID",
      process.env.MICROSOFT_GRAPH_TENANT_ID,
      "common",
    );

    const auth = initializeAuth({ clientId, tenantId });

    process.stderr.write(
      "\nStarting Microsoft Graph device-code authentication...\n",
    );
    const tokenInfo = await auth.authenticate();

    process.stderr.write("\nAuthentication successful.\n");
    process.stderr.write(`User: ${tokenInfo.account.username}\n`);
    if (tokenInfo.account.name) {
      process.stderr.write(`Name: ${tokenInfo.account.name}\n`);
    }
    if (tokenInfo.account.tenantId) {
      process.stderr.write(`Tenant: ${tokenInfo.account.tenantId}\n`);
    }
    process.stderr.write(
      `Token expires: ${tokenInfo.expiresOn.toLocaleString()}\n`,
    );
  } finally {
    rl.close();
  }
}
