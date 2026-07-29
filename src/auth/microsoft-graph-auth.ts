/**
 * Microsoft Graph client-credential authentication.
 *
 * Credentials arrive only from the 1Password launcher as process-local
 * environment variables. This module keeps access tokens in memory only.
 */

import {
  ClientCredentialRequest,
  ConfidentialClientApplication,
} from "@azure/msal-node";
import { DEFAULT_SCOPES } from "../config/scopes.js";

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

export class MicrosoftGraphAuth {
  private readonly config: Required<Pick<AuthConfig, "clientId" | "tenantId" | "scopes">> &
    Pick<AuthConfig, "clientSecret">;
  private inMemoryToken: TokenInfo | null = null;
  private inflightRefresh: Promise<string> | null = null;

  constructor(config: AuthConfig) {
    this.config = {
      clientId: config.clientId,
      tenantId: config.tenantId ?? "common",
      scopes: config.scopes ?? [...DEFAULT_SCOPES],
      clientSecret: config.clientSecret,
    };
  }

  async authenticate(): Promise<TokenInfo> {
    throw new Error(
      "Delegated device-code authentication is disabled because this tool cannot persist tokens outside 1Password. Ask a 1Password owner to provision client-credential items.",
    );
  }

  async getAccessToken(): Promise<string> {
    if (!this.config.clientSecret) {
      throw new Error(
        "Missing Microsoft Graph client secret from the 1Password launcher. Ask a 1Password owner to provision cpz::SP_CLIENT_SECRET.",
      );
    }

    if (this.inMemoryToken && this.isTokenValid(this.inMemoryToken)) {
      return this.inMemoryToken.accessToken;
    }

    if (!this.inflightRefresh) {
      this.inflightRefresh = this.acquireClientCredentialToken().finally(() => {
        this.inflightRefresh = null;
      });
    }
    return this.inflightRefresh;
  }

  prewarm(): void {
    void this.getAccessToken().catch(() => undefined);
  }

  async isAuthenticated(): Promise<boolean> {
    try {
      await this.getAccessToken();
      return true;
    } catch {
      return false;
    }
  }

  async signOut(): Promise<void> {
    this.inMemoryToken = null;
  }

  async getCurrentUser(): Promise<TokenInfo["account"] | null> {
    return this.inMemoryToken?.account ?? null;
  }

  private async acquireClientCredentialToken(): Promise<string> {
    const client = new ConfidentialClientApplication({
      auth: {
        clientId: this.config.clientId,
        authority: `https://login.microsoftonline.com/${this.config.tenantId}`,
        clientSecret: this.config.clientSecret,
      },
    });
    const request: ClientCredentialRequest = {
      scopes: ["https://graph.microsoft.com/.default"],
    };
    const result = await client.acquireTokenByClientCredential(request);
    if (!result?.accessToken) {
      throw new Error("Client credentials flow returned no token");
    }

    this.inMemoryToken = {
      accessToken: result.accessToken,
      expiresOn: result.expiresOn ?? new Date(Date.now() + 3_600_000),
      account: { username: `app:${this.config.clientId}` },
    };
    return this.inMemoryToken.accessToken;
  }

  private isTokenValid(token: TokenInfo): boolean {
    return token.expiresOn.getTime() - Date.now() > 5 * 60_000;
  }
}

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
