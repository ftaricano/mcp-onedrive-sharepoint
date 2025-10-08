/**
 * Configuration management for MCP OneDrive/SharePoint Server
 */

import * as dotenv from 'dotenv';
import { readFileSync, existsSync } from 'fs';
import { join } from 'path';

// Load environment variables
dotenv.config();

export interface AuthConfig {
  clientId: string;
  tenantId: string;
  scopes: string[];
}

export interface ServerConfig {
  auth: AuthConfig;
  graph: {
    baseUrl: string;
    timeout: number;
    maxRetries: number;
  };
  cache: {
    enabled: boolean;
    ttl: number;
  };
}

const DEFAULT_SCOPES = [
  'Files.ReadWrite.All',
  'Sites.ReadWrite.All', 
  'Directory.Read.All',
  'User.Read',
  'offline_access'
];

export function loadConfig(): ServerConfig {
  const config: ServerConfig = {
    auth: {
      clientId: process.env.MICROSOFT_GRAPH_CLIENT_ID || '',
      tenantId: process.env.MICROSOFT_GRAPH_TENANT_ID || 'common',
      scopes: process.env.MICROSOFT_GRAPH_SCOPES?.split(',') || DEFAULT_SCOPES
    },
    graph: {
      baseUrl: process.env.MICROSOFT_GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0',
      timeout: parseInt(process.env.MICROSOFT_GRAPH_TIMEOUT || '30000'),
      maxRetries: parseInt(process.env.MICROSOFT_GRAPH_MAX_RETRIES || '3')
    },
    cache: {
      enabled: process.env.MICROSOFT_GRAPH_CACHE_ENABLED !== 'false',
      ttl: parseInt(process.env.MICROSOFT_GRAPH_CACHE_TTL || '3600')
    }
  };

  // Validate required configuration
  if (!config.auth.clientId) {
    throw new Error(
      'Missing MICROSOFT_GRAPH_CLIENT_ID environment variable. ' +
      'Please run the setup-auth script or set the environment variable.'
    );
  }

  return config;
}

export function validateConfig(config: ServerConfig): void {
  if (!config.auth.clientId) {
    throw new Error('Client ID is required for Microsoft Graph authentication');
  }

  if (!Array.isArray(config.auth.scopes) || config.auth.scopes.length === 0) {
    throw new Error('At least one scope is required for Microsoft Graph authentication');
  }

  // Validate required scopes
  const requiredScopes = ['Files.ReadWrite.All', 'Sites.ReadWrite.All'];
  const missingScopes = requiredScopes.filter(scope => !config.auth.scopes.includes(scope));
  
  if (missingScopes.length > 0) {
    throw new Error(`Missing required scopes: ${missingScopes.join(', ')}`);
  }
}

// Export configuration constants
export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
export const DEFAULT_TIMEOUT = 30000;
export const DEFAULT_MAX_RETRIES = 3;