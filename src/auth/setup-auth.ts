#!/usr/bin/env tsx

/**
 * Authentication setup script for OneDrive/SharePoint MCP server
 * Run this script to authenticate with Microsoft Graph
 */

import { initializeAuth } from './microsoft-graph-auth';
import * as readline from 'readline';

async function setupAuthentication() {
  console.log('🔧 OneDrive/SharePoint MCP Server - Authentication Setup');
  console.log('='.repeat(60));

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  const question = (query: string): Promise<string> => {
    return new Promise(resolve => rl.question(query, resolve));
  };

  try {
    // Get client ID
    const clientId = process.env.MICROSOFT_GRAPH_CLIENT_ID || 
      await question('Enter your Azure App Client ID: ');

    if (!clientId.trim()) {
      throw new Error('Client ID is required');
    }

    // Get tenant ID (optional)
    const defaultTenant = process.env.MICROSOFT_GRAPH_TENANT_ID || 'common';
    const tenantInput = await question(`Enter your Tenant ID (press Enter for '${defaultTenant}'): `);
    const tenantId = tenantInput.trim() || defaultTenant;

    // Initialize auth
    const auth = initializeAuth({
      clientId: clientId.trim(),
      tenantId: tenantId.trim()
    });

    console.log('\n🔑 Starting authentication process...');
    
    // Authenticate
    const tokenInfo = await auth.authenticate();
    
    console.log('\n✅ Authentication successful!');
    console.log(`👤 User: ${tokenInfo.account.username}`);
    if (tokenInfo.account.name) {
      console.log(`📝 Name: ${tokenInfo.account.name}`);
    }
    if (tokenInfo.account.tenantId) {
      console.log(`🏢 Tenant: ${tokenInfo.account.tenantId}`);
    }
    console.log(`⏰ Token expires: ${tokenInfo.expiresOn.toLocaleString()}`);

    console.log('\n🎉 Setup complete! You can now use the MCP server.');
    console.log('\nNext steps:');
    console.log('1. Add this server to your Claude Code configuration');
    console.log('2. Run: npm run build && npm start');

  } catch (error) {
    console.error('\n❌ Authentication setup failed:');
    console.error(error instanceof Error ? error.message : 'Unknown error');
    process.exit(1);
  } finally {
    rl.close();
  }
}

// Handle environment variables
function checkEnvironment() {
  const requiredEnvVars = ['MICROSOFT_GRAPH_CLIENT_ID'];
  const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);

  if (missingVars.length > 0) {
    console.log('⚠️  Missing environment variables:');
    missingVars.forEach(varName => {
      console.log(`   - ${varName}`);
    });
    console.log('\nYou can either:');
    console.log('1. Set environment variables in .env file');
    console.log('2. Provide them during this setup process');
    console.log('');
  }
}

// Main execution
async function main() {
  checkEnvironment();
  await setupAuthentication();
}

if (require.main === module) {
  main().catch(error => {
    console.error('Setup failed:', error);
    process.exit(1);
  });
}