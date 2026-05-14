/**
 * Microsoft Graph authentication using device code flow
 * Optimized for CLI/MCP environments with secure token storage
 */

import {
  AccountInfo,
  AuthenticationResult,
  ClientCredentialRequest,
  ConfidentialClientApplication,
  DeviceCodeRequest,
  ICachePlugin,
  PublicClientApplication,
} from "@azure/msal-node";
import { createHash } from "node:crypto";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { homedir } from "node:os";
import { join } from "node:path";
import { DEFAULT_SCOPES } from "../config/scopes.js";

const lazyRequire = createRequire(import.meta.url);
let cachedKeytar: SecureStore | null = null;
const issuedFallbackWarnings = new Set<string>();
function getDefaultKeychain(): SecureStore {
  if (!cachedKeytar) {
    cachedKeytar = lazyRequire("keytar") as SecureStore;
  }
  return cachedKeytar;
}

export interface AuthConfig {
  clientId: string;
  tenantId?: string;
  scopes?: string[];
  clientSecret?: string;
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
  fallbackStore?: SecureStore;
  pca?: PublicClientApplication;
}

class FileFallbackStore implements SecureStore {
  constructor(private readonly cacheDir = getDefaultFallbackCacheDir()) {}

  async getPassword(service: string, account: string): Promise<string | null> {
    try {
      return await readFile(this.getCachePath(service, account), "utf-8");
    } catch (error) {
      if (isFileNotFound(error)) {
        return null;
      }
      throw error;
    }
  }

  async setPassword(
    service: string,
    account: string,
    password: string,
  ): Promise<void> {
    await mkdir(this.cacheDir, { recursive: true, mode: 0o700 });
    await writeFile(this.getCachePath(service, account), password, {
      encoding: "utf-8",
      mode: 0o600,
    });
  }

  async deletePassword(service: string, account: string): Promise<boolean> {
    try {
      await rm(this.getCachePath(service, account), { force: true });
      return true;
    } catch (error) {
      if (isFileNotFound(error)) {
        return false;
      }
      throw error;
    }
  }

  private getCachePath(service: string, account: string): string {
    const cacheKey = createHash("sha256")
      .update(`${service}:${account}`)
      .digest("hex");
    return join(this.cacheDir, `${cacheKey}.json`);
  }
}

function getDefaultFallbackCacheDir(): string {
  const baseDir =
    process.env.MCP_ONEDRIVE_SHAREPOINT_CACHE_DIR ??
    process.env.XDG_CACHE_HOME ??
    (process.platform === "win32" ? process.env.LOCALAPPDATA : undefined) ??
    join(homedir(), ".cache");
  return join(baseDir, "mcp-onedrive-sharepoint");
}

function isFileNotFound(error: unknown): boolean {
  return (
    error instanceof Error &&
    "code" in error &&
    (error as NodeJS.ErrnoException).code === "ENOENT"
  );
}

function warnOnce(key: string, message: string, error: unknown): void {
  if (issuedFallbackWarnings.has(key)) {
    return;
  }

  issuedFallbackWarnings.add(key);
  const detail = error instanceof Error ? error.message : String(error);
  console.warn(`${message}: ${detail}`);
}

async function readWithFallback(
  keychain: SecureStore,
  fallbackStore: SecureStore,
  serviceKeyName: string,
  account: string,
): Promise<string | null> {
  try {
    const secureValue = await keychain.getPassword(serviceKeyName, account);
    if (secureValue) {
      return secureValue;
    }
  } catch (error) {
    warnOnce(
      `read:${serviceKeyName}:${account}`,
      "Failed to read from system keychain; trying file cache",
      error,
    );
  }

  return fallbackStore.getPassword(serviceKeyName, account);
}

async function writeWithFallback(
  keychain: SecureStore,
  fallbackStore: SecureStore,
  serviceKeyName: string,
  account: string,
  value: string,
): Promise<void> {
  try {
    await keychain.setPassword(serviceKeyName, account, value);
  } catch (error) {
    warnOnce(
      `write:${serviceKeyName}:${account}`,
      "Failed to write to system keychain; using file cache",
      error,
    );
    await fallbackStore.setPassword(serviceKeyName, account, value);
    return;
  }

  try {
    await fallbackStore.deletePassword(serviceKeyName, account);
  } catch (error) {
    warnOnce(
      `delete-stale-file:${serviceKeyName}:${account}`,
      "Failed to clear stale file cache entry",
      error,
    );
  }
}

async function deleteFromAllStores(
  keychain: SecureStore,
  fallbackStore: SecureStore,
  serviceKeyName: string,
  account: string,
): Promise<void> {
  try {
    await keychain.deletePassword(serviceKeyName, account);
  } catch (error) {
    warnOnce(
      `delete-keychain:${serviceKeyName}:${account}`,
      "Failed to clear system keychain entry",
      error,
    );
  }

  try {
    await fallbackStore.deletePassword(serviceKeyName, account);
  } catch (error) {
    warnOnce(
      `delete-file:${serviceKeyName}:${account}`,
      "Failed to clear file cache entry",
      error,
    );
  }
}

export function createKeychainMsalCachePlugin(
  keychain: SecureStore,
  serviceKeyName: string,
  cacheAccount: string,
  fallbackStore: SecureStore = new FileFallbackStore(),
): ICachePlugin {
  return {
    beforeCacheAccess: async (tokenCacheContext) => {
      const cacheSnapshot = await readWithFallback(
        keychain,
        fallbackStore,
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
      await writeWithFallback(
        keychain,
        fallbackStore,
        serviceKeyName,
        cacheAccount,
        serializedCache,
      );
    },
  };
}

export class MicrosoftGraphAuth {
  private pca: PublicClientApplication;
  private config: AuthConfig;
  private keychain: SecureStore;
  private fallbackStore: SecureStore;
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

    this.keychain = dependencies.keychain ?? getDefaultKeychain();
    this.fallbackStore = dependencies.fallbackStore ?? new FileFallbackStore();
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
   * When clientSecret is configured, uses client credentials flow (no user interaction).
   * In-memory cache short-circuits Keychain I/O for warm calls.
   */
  async getAccessToken(): Promise<string> {
    if (this.config.clientSecret) {
      return this.getAccessTokenByClientCredentials();
    }

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

  private async getAccessTokenByClientCredentials(): Promise<string> {
    if (this.inMemoryToken && this.isTokenValid(this.inMemoryToken)) {
      return this.inMemoryToken.accessToken;
    }

    const cca = new ConfidentialClientApplication({
      auth: {
        clientId: this.config.clientId,
        authority: `https://login.microsoftonline.com/${this.config.tenantId}`,
        clientSecret: this.config.clientSecret,
      },
    });

    const request: ClientCredentialRequest = {
      scopes: ["https://graph.microsoft.com/.default"],
    };

    const result = await cca.acquireTokenByClientCredential(request);
    if (!result?.accessToken) {
      throw new Error("Client credentials flow returned no token");
    }

    this.inMemoryToken = {
      accessToken: result.accessToken,
      expiresOn: result.expiresOn ?? new Date(Date.now() + 3600 * 1000),
      account: { username: `app:${this.config.clientId}` },
    };
    return this.inMemoryToken.accessToken;
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
      await deleteFromAllStores(
        this.keychain,
        this.fallbackStore,
        this.serviceKeyName,
        this.accessTokenCacheAccount,
      );
      await deleteFromAllStores(
        this.keychain,
        this.fallbackStore,
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
      await writeWithFallback(
        this.keychain,
        this.fallbackStore,
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
      const tokenData = await readWithFallback(
        this.keychain,
        this.fallbackStore,
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
      this.fallbackStore,
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
