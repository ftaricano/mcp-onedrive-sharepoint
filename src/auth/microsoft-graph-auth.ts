/**
 * Microsoft Graph authentication using device code flow
 * Optimized for CLI/MCP environments with secure token storage
 */

import { PublicClientApplication, AuthenticationResult, DeviceCodeRequest } from '@azure/msal-node';
import * as keytar from 'keytar';
import { DEFAULT_SCOPES } from '../config/scopes.js';

export interface AuthConfig {
  clientId: string;
  tenantId?: string;
  scopes?: string[];
}

export interface TokenInfo {
  accessToken: string;
  expiresOn: Date;
  account: {
    username: string;
    name?: string;
    tenantId?: string;
  };
}

export class MicrosoftGraphAuth {
  private pca: PublicClientApplication;
  private config: AuthConfig;
  private readonly serviceKeyName = 'mcp-onedrive-sharepoint';

  constructor(config: AuthConfig) {
    this.config = {
      tenantId: 'common',
      scopes: [...DEFAULT_SCOPES],
      ...config
    };

    this.pca = new PublicClientApplication({
      auth: {
        clientId: this.config.clientId,
        authority: `https://login.microsoftonline.com/${this.config.tenantId}`
      }
    });
  }

  /**
   * Authenticate using device code flow
   * Perfect for CLI applications - shows code to user for browser authentication
   */
  async authenticate(): Promise<TokenInfo> {
    try {
      // First, try to get cached token
      const cachedToken = await this.getCachedToken();
      if (cachedToken && this.isTokenValid(cachedToken)) {
        return cachedToken;
      }

      // If no valid cached token, start device code flow
      console.log('Starting Microsoft Graph authentication...');
      
      const deviceCodeRequest: DeviceCodeRequest = {
        scopes: this.config.scopes!,
        deviceCodeCallback: (response) => {
          console.log('\n=== Microsoft Graph Authentication ===');
          console.log(`Please visit: ${response.verificationUri}`);
          console.log(`Enter code: ${response.userCode}`);
          console.log('Waiting for authentication...\n');
        }
      };

      const result = await this.pca.acquireTokenByDeviceCode(deviceCodeRequest);
      
      if (!result) {
        throw new Error('Authentication failed - no result returned');
      }

      const tokenInfo = this.extractTokenInfo(result);
      
      // Cache the token securely
      await this.cacheToken(tokenInfo);
      
      console.log(`✅ Successfully authenticated as: ${tokenInfo.account.username}`);
      return tokenInfo;

    } catch (error) {
      console.error('Authentication failed:', error);
      throw new Error(`Microsoft Graph authentication failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Get a valid access token, refreshing if necessary
   */
  async getAccessToken(): Promise<string> {
    const cachedToken = await this.getCachedToken();
    
    if (cachedToken && this.isTokenValid(cachedToken)) {
      return cachedToken.accessToken;
    }

    // Try to refresh token silently
    try {
      const accounts = await this.pca.getTokenCache().getAllAccounts();
      
      if (accounts.length > 0) {
        const silentRequest = {
          scopes: this.config.scopes!,
          account: accounts[0]
        };

        const result = await this.pca.acquireTokenSilent(silentRequest);
        if (result) {
          const tokenInfo = this.extractTokenInfo(result);
          await this.cacheToken(tokenInfo);
          return tokenInfo.accessToken;
        }
      }
    } catch (error) {
      console.log('Silent token refresh failed, re-authentication required');
    }

    // If silent refresh fails, require re-authentication
    const tokenInfo = await this.authenticate();
    return tokenInfo.accessToken;
  }

  /**
   * Check if user is currently authenticated
   */
  async isAuthenticated(): Promise<boolean> {
    try {
      const cachedToken = await this.getCachedToken();
      return cachedToken ? this.isTokenValid(cachedToken) : false;
    } catch {
      return false;
    }
  }

  /**
   * Sign out and clear cached tokens
   */
  async signOut(): Promise<void> {
    try {
      // Clear keychain
      await keytar.deletePassword(this.serviceKeyName, 'access_token');
      
      // Clear MSAL cache
      const accounts = await this.pca.getTokenCache().getAllAccounts();
      for (const account of accounts) {
        await this.pca.getTokenCache().removeAccount(account);
      }
      
      console.log('✅ Successfully signed out');
    } catch (error) {
      console.error('Error during sign out:', error);
    }
  }

  /**
   * Get current user information
   */
  async getCurrentUser(): Promise<TokenInfo['account'] | null> {
    try {
      const cachedToken = await this.getCachedToken();
      return cachedToken?.account || null;
    } catch {
      return null;
    }
  }

  // Private helper methods

  private extractTokenInfo(result: AuthenticationResult): TokenInfo {
    if (!result.accessToken || !result.expiresOn || !result.account) {
      throw new Error('Invalid authentication result');
    }

    return {
      accessToken: result.accessToken,
      expiresOn: result.expiresOn,
      account: {
        username: result.account.username,
        name: result.account.name || undefined,
        tenantId: result.account.tenantId || undefined
      }
    };
  }

  private async cacheToken(tokenInfo: TokenInfo): Promise<void> {
    try {
      const tokenData = JSON.stringify(tokenInfo);
      await keytar.setPassword(this.serviceKeyName, 'access_token', tokenData);
    } catch (error) {
      console.warn('Failed to cache token securely:', error);
    }
  }

  private async getCachedToken(): Promise<TokenInfo | null> {
    try {
      const tokenData = await keytar.getPassword(this.serviceKeyName, 'access_token');
      if (!tokenData) return null;

      const tokenInfo = JSON.parse(tokenData) as TokenInfo;
      
      // Ensure expiresOn is a Date object
      tokenInfo.expiresOn = new Date(tokenInfo.expiresOn);
      
      return tokenInfo;
    } catch (error) {
      console.warn('Failed to retrieve cached token:', error);
      return null;
    }
  }

  private isTokenValid(tokenInfo: TokenInfo): boolean {
    const now = new Date();
    const expiry = new Date(tokenInfo.expiresOn);
    
    // Add 5 minute buffer for token expiry
    const bufferTime = 5 * 60 * 1000;
    return expiry.getTime() - now.getTime() > bufferTime;
  }
}

// Singleton instance for the MCP server
let authInstance: MicrosoftGraphAuth | null = null;

export function initializeAuth(config: AuthConfig): MicrosoftGraphAuth {
  authInstance = new MicrosoftGraphAuth(config);
  return authInstance;
}

export function getAuthInstance(): MicrosoftGraphAuth {
  if (!authInstance) {
    throw new Error('Authentication not initialized. Call initializeAuth() first.');
  }
  return authInstance;
}

export function __setAuthInstanceForTests(auth: MicrosoftGraphAuth | null): void {
  authInstance = auth;
}