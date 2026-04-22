/**
 * Microsoft Graph authentication using device code flow
 * Optimized for CLI/MCP environments with secure token storage
 */

import {
  AccountInfo,
  AuthenticationResult,
  DeviceCodeRequest,
  ICachePlugin,
  PublicClientApplication,
} from "@azure/msal-node";
import keytar from "keytar";
import { DEFAULT_SCOPES } from "../config/scopes.js";

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

interface SecureStore {
  getPassword(service: string, account: string): Promise<string | null>;
  setPassword(service: string, account: string, password: string): Promise<void>;
  deletePassword(service: string, account: string): Promise<boolean>;
}

interface MicrosoftGraphAuthDependencies {
  keychain?: SecureStore;
  pca?: PublicClientApplication;
}

export function createKeychainMsalCachePlugin(
  keychain: SecureStore,
  serviceKeyName: string,
  cacheAccount: string,
): ICachePlugin {
  return {
    beforeCacheAccess: async (tokenCacheContext) => {
      const cacheSnapshot = await keychain.getPassword(
        serviceKeyName,
        cacheAccount,
      );

      if (cacheSnapshot) {
        tokenCacheContext.tokenCache.deserialize(cacheSnapshot);
      }
    },
    afterCacheAccess: async (tokenCacheContext) => {
      if (!tokenCacheContext.cacheHasChanged) {
        return;
      }

      const serializedCache = tokenCacheContext.tokenCache.serialize();
      await keychain.setPassword(serviceKeyName, cacheAccount, serializedCache);
    },
  };
}

export class MicrosoftGraphAuth {
  private pca: PublicClientApplication;
  private config: AuthConfig;
  private keychain: SecureStore;
  private readonly serviceKeyName = "mcp-onedrive-sharepoint";
  private readonly accessTokenCacheAccount = "access_token";
  private readonly msalCacheAccount = "msal_token_cache";
  private inMemoryToken: TokenInfo | null = null;
  private inflightRefresh: Promise<TokenInfo | null> | null = null;

  constructor(
    config: AuthConfig,
    dependencies: MicrosoftGraphAuthDependencies = {},
  ) {
    this.config = {
      tenantId: "common",
      scopes: [...DEFAULT_SCOPES],
      ...config,
    };

    this.keychain = dependencies.keychain ?? keytar;
    this.pca =
      dependencies.pca ??
      new PublicClientApplication({
        auth: {
          clientId: this.config.clientId,
          authority: `https://login.microsoftonline.com/${this.config.tenantId}`,
        },
        cache: {
          cachePlugin: this.createCachePlugin(),
        },
      });
  }

  /**
   * Authenticate using device code flow
   * Perfect for CLI applications - shows code to user for browser authentication
   */
  async authenticate(): Promise<TokenInfo> {
    try {
      if (this.inMemoryToken && this.isTokenValid(this.inMemoryToken)) {
        return this.inMemoryToken;
      }

      const cachedToken = await this.getCachedToken();
      if (cachedToken && this.isTokenValid(cachedToken)) {
        this.inMemoryToken = cachedToken;
        return cachedToken;
      }

      const silentlyRefreshedToken = await this.tryAcquireTokenSilently();
      if (silentlyRefreshedToken) {
        this.inMemoryToken = silentlyRefreshedToken;
        return silentlyRefreshedToken;
      }

      // If no valid cached token, start device code flow
      console.log("Starting Microsoft Graph authentication...");

      const deviceCodeRequest: DeviceCodeRequest = {
        scopes: this.config.scopes!,
        deviceCodeCallback: (response) => {
          console.log("\n=== Microsoft Graph Authentication ===");
          console.log(`Please visit: ${response.verificationUri}`);
          console.log(`Enter code: ${response.userCode}`);
          console.log("Waiting for authentication...\n");
        },
      };

      const result = await this.pca.acquireTokenByDeviceCode(deviceCodeRequest);

      if (!result) {
        throw new Error("Authentication failed - no result returned");
      }

      const tokenInfo = this.extractTokenInfo(result);
      this.inMemoryToken = tokenInfo;
      await this.cacheToken(tokenInfo);

      console.log(
        `✅ Successfully authenticated as: ${tokenInfo.account.username}`,
      );
      return tokenInfo;
    } catch (error) {
      console.error("Authentication failed:", error);
      throw new Error(
        `Microsoft Graph authentication failed: ${error instanceof Error ? error.message : "Unknown error"}`,
      );
    }
  }

  /**
   * Get a valid access token, refreshing if necessary.
   * In-memory cache short-circuits Keychain I/O for warm calls.
   */
  async getAccessToken(): Promise<string> {
    if (this.inMemoryToken && this.isTokenValid(this.inMemoryToken)) {
      return this.inMemoryToken.accessToken;
    }

    if (!this.inflightRefresh) {
      this.inflightRefresh = this.loadOrRefreshToken().finally(() => {
        this.inflightRefresh = null;
      });
    }

    const refreshed = await this.inflightRefresh;
    if (refreshed) {
      return refreshed.accessToken;
    }

    const tokenInfo = await this.authenticate();
    return tokenInfo.accessToken;
  }

  private async loadOrRefreshToken(): Promise<TokenInfo | null> {
    const cachedToken = await this.getCachedToken();
    if (cachedToken && this.isTokenValid(cachedToken)) {
      this.inMemoryToken = cachedToken;
      return cachedToken;
    }

    const silent = await this.tryAcquireTokenSilently();
    if (silent) {
      this.inMemoryToken = silent;
      return silent;
    }

    return null;
  }

  /**
   * Fire-and-forget warm-up: kicks off token load during MCP handshake
   * so the first tool call finds an in-memory token ready.
   */
  prewarm(): void {
    if (this.inMemoryToken && this.isTokenValid(this.inMemoryToken)) return;
    if (this.inflightRefresh) return;
    this.inflightRefresh = this.loadOrRefreshToken()
      .catch(() => null)
      .finally(() => {
        this.inflightRefresh = null;
      });
  }

  /**
   * Check if user is currently authenticated
   */
  async isAuthenticated(): Promise<boolean> {
    try {
      if (this.inMemoryToken && this.isTokenValid(this.inMemoryToken)) {
        return true;
      }

      const cachedToken = await this.getCachedToken();
      if (cachedToken && this.isTokenValid(cachedToken)) {
        this.inMemoryToken = cachedToken;
        return true;
      }

      const silent = await this.tryAcquireTokenSilently();
      if (silent) {
        this.inMemoryToken = silent;
        return true;
      }
      return false;
    } catch {
      return false;
    }
  }

  /**
   * Sign out and clear cached tokens
   */
  async signOut(): Promise<void> {
    try {
      this.inMemoryToken = null;
      await this.keychain.deletePassword(
        this.serviceKeyName,
        this.accessTokenCacheAccount,
      );
      await this.keychain.deletePassword(
        this.serviceKeyName,
        this.msalCacheAccount,
      );

      // Clear MSAL cache
      const accounts = await this.pca.getTokenCache().getAllAccounts();
      for (const account of accounts) {
        await this.pca.getTokenCache().removeAccount(account);
      }

      console.log("✅ Successfully signed out");
    } catch (error) {
      console.error("Error during sign out:", error);
    }
  }

  /**
   * Get current user information
   */
  async getCurrentUser(): Promise<TokenInfo["account"] | null> {
    try {
      if (this.inMemoryToken) return this.inMemoryToken.account;
      const cachedToken = await this.getCachedToken();
      if (cachedToken) this.inMemoryToken = cachedToken;
      return cachedToken?.account || null;
    } catch {
      return null;
    }
  }

  // Private helper methods

  private extractTokenInfo(result: AuthenticationResult): TokenInfo {
    if (!result.accessToken || !result.expiresOn || !result.account) {
      throw new Error("Invalid authentication result");
    }

    return {
      accessToken: result.accessToken,
      expiresOn: result.expiresOn,
      account: {
        username: result.account.username,
        name: result.account.name || undefined,
        tenantId: result.account.tenantId || undefined,
      },
    };
  }

  private async cacheToken(tokenInfo: TokenInfo): Promise<void> {
    try {
      const tokenData = JSON.stringify(tokenInfo);
      await this.keychain.setPassword(
        this.serviceKeyName,
        this.accessTokenCacheAccount,
        tokenData,
      );
    } catch (error) {
      console.warn("Failed to cache token securely:", error);
    }
  }

  private async getCachedToken(): Promise<TokenInfo | null> {
    try {
      const tokenData = await this.keychain.getPassword(
        this.serviceKeyName,
        this.accessTokenCacheAccount,
      );
      if (!tokenData) return null;

      const tokenInfo = JSON.parse(tokenData) as TokenInfo;

      // Ensure expiresOn is a Date object
      tokenInfo.expiresOn = new Date(tokenInfo.expiresOn);

      return tokenInfo;
    } catch (error) {
      console.warn("Failed to retrieve cached token:", error);
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

  private createCachePlugin(): ICachePlugin {
    return createKeychainMsalCachePlugin(
      this.keychain,
      this.serviceKeyName,
      this.msalCacheAccount,
    );
  }

  private async tryAcquireTokenSilently(): Promise<TokenInfo | null> {
    try {
      const account = await this.getCachedAccount();
      if (!account) {
        return null;
      }

      const result = await this.pca.acquireTokenSilent({
        scopes: this.config.scopes!,
        account,
      });

      if (!result) {
        return null;
      }

      const tokenInfo = this.extractTokenInfo(result);
      this.inMemoryToken = tokenInfo;
      await this.cacheToken(tokenInfo);
      return tokenInfo;
    } catch {
      console.log("Silent token refresh failed, re-authentication required");
      return null;
    }
  }

  private async getCachedAccount(): Promise<AccountInfo | null> {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    if (accounts.length === 0) {
      return null;
    }

    const cachedUser = await this.getCurrentUser();
    if (!cachedUser?.username) {
      return accounts[0] ?? null;
    }

    return (
      accounts.find((account) => account.username === cachedUser.username) ??
      accounts[0] ??
      null
    );
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
    throw new Error(
      "Authentication not initialized. Call initializeAuth() first.",
    );
  }
  return authInstance;
}

export function __setAuthInstanceForTests(
  auth: MicrosoftGraphAuth | null,
): void {
  authInstance = auth;
}
