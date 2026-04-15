/**
 * Configuration management for MCP OneDrive/SharePoint Server
 */

import * as dotenv from 'dotenv';
import { DEFAULT_SCOPES } from './scopes.js';

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

export function loadConfig(): ServerConfig {
  const config: ServerConfig = {
    auth: {
      clientId: process.env.MICROSOFT_GRAPH_CLIENT_ID || '',
      tenantId: process.env.MICROSOFT_GRAPH_TENANT_ID || 'common',
      scopes: process.env.MICROSOFT_GRAPH_SCOPES?.split(',').map((scope) => scope.trim()).filter(Boolean) || [...DEFAULT_SCOPES]
    },
    graph: {
      baseUrl: process.env.MICROSOFT_GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0',
      timeout: parseInt(process.env.MICROSOFT_GRAPH_TIMEOUT || '30000', 10),
      maxRetries: parseInt(process.env.MICROSOFT_GRAPH_MAX_RETRIES || '3', 10)
    },
    cache: {
      enabled: process.env.MICROSOFT_GRAPH_CACHE_ENABLED !== 'false',
      ttl: parseInt(process.env.MICROSOFT_GRAPH_CACHE_TTL || '3600', 10)
    }
  };

  validateConfig(config);
  return config;
}

export function validateConfig(config: ServerConfig): void {
  if (!config.auth.clientId) {
    throw new Error(
      'Missing MICROSOFT_GRAPH_CLIENT_ID environment variable. Run npm run setup-auth or set it in your environment/.env file.'
    );
  }

  if (!Array.isArray(config.auth.scopes) || config.auth.scopes.length === 0) {
    throw new Error('At least one Microsoft Graph scope is required');
  }

  const requiredScopes = ['Files.ReadWrite.All', 'Sites.ReadWrite.All'];
  const missingScopes = requiredScopes.filter((scope) => !config.auth.scopes.includes(scope));

  if (missingScopes.length > 0) {
    throw new Error(`Missing required scopes: ${missingScopes.join(', ')}`);
  }
}

export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
export const DEFAULT_TIMEOUT = 30000;
export const DEFAULT_MAX_RETRIES = 3;
