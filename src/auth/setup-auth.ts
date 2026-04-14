#!/usr/bin/env tsx

/**
 * Authentication setup script for OneDrive/SharePoint MCP server.
 */

import * as dotenv from 'dotenv';
import * as readline from 'node:readline/promises';
import { stdin as input, stdout as output } from 'node:process';
import { fileURLToPath } from 'node:url';
import { initializeAuth } from './microsoft-graph-auth.js';

dotenv.config();

async function promptForMissingValue(
  rl: readline.Interface,
  label: string,
  currentValue?: string,
  fallback?: string
): Promise<string> {
  if (currentValue?.trim()) {
    return currentValue.trim();
  }

  const suffix = fallback ? ` (${fallback})` : '';
  const answer = await rl.question(`${label}${suffix}: `);
  return (answer.trim() || fallback || '').trim();
}

async function setupAuthentication(): Promise<void> {
  console.log('OneDrive/SharePoint MCP Server - Authentication Setup');
  console.log('======================================================');

  const rl = readline.createInterface({ input, output });

  try {
    const clientId = await promptForMissingValue(
      rl,
      'Azure App Client ID',
      process.env.MICROSOFT_GRAPH_CLIENT_ID
    );

    if (!clientId) {
      throw new Error('Client ID is required');
    }

    const tenantId = await promptForMissingValue(
      rl,
      'Tenant ID',
      process.env.MICROSOFT_GRAPH_TENANT_ID,
      'common'
    );

    const auth = initializeAuth({ clientId, tenantId });

    console.log('\nStarting Microsoft Graph device-code authentication...');
    const tokenInfo = await auth.authenticate();

    console.log('\nAuthentication successful.');
    console.log(`User: ${tokenInfo.account.username}`);
    if (tokenInfo.account.name) {
      console.log(`Name: ${tokenInfo.account.name}`);
    }
    if (tokenInfo.account.tenantId) {
      console.log(`Tenant: ${tokenInfo.account.tenantId}`);
    }
    console.log(`Token expires: ${tokenInfo.expiresOn.toLocaleString()}`);
    console.log('\nNext steps: npm run build && npm start');
  } finally {
    rl.close();
  }
}

async function main(): Promise<void> {
  if (!process.env.MICROSOFT_GRAPH_CLIENT_ID) {
    console.log('Tip: you can set MICROSOFT_GRAPH_CLIENT_ID in .env to avoid retyping it.\n');
  }

  await setupAuthentication();
}

const isDirectRun = process.argv[1] === fileURLToPath(import.meta.url);

if (isDirectRun) {
  main().catch((error) => {
    console.error('\nAuthentication setup failed:');
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  });
}
