import assert from "node:assert/strict";
import test from "node:test";

import {
  MicrosoftGraphAuth,
  createKeychainMsalCachePlugin,
} from "../auth/microsoft-graph-auth.js";

class InMemoryKeychain {
  private readonly store = new Map<string, string>();

  async getPassword(service: string, account: string): Promise<string | null> {
    return this.store.get(`${service}:${account}`) ?? null;
  }

  async setPassword(
    service: string,
    account: string,
    password: string,
  ): Promise<void> {
    this.store.set(`${service}:${account}`, password);
  }

  async deletePassword(service: string, account: string): Promise<boolean> {
    return this.store.delete(`${service}:${account}`);
  }
}

test("MSAL cache plugin persists serialized cache snapshots across executions", async () => {
  const keychain = new InMemoryKeychain();
  const plugin = createKeychainMsalCachePlugin(
    keychain,
    "mcp-onedrive-sharepoint",
    "msal_token_cache",
  );

  let restoredSnapshot = "";

  await plugin.afterCacheAccess({
    cacheHasChanged: true,
    tokenCache: {
      serialize: () => JSON.stringify({ Account: { abc: { username: "user@example.com" } } }),
      deserialize: () => undefined,
    },
  } as any);

  await plugin.beforeCacheAccess({
    tokenCache: {
      serialize: () => "",
      deserialize: (snapshot: string) => {
        restoredSnapshot = snapshot;
      },
    },
  } as any);

  assert.match(restoredSnapshot, /user@example.com/);
});

test("expired access tokens are silently refreshed when MSAL account state is available", async () => {
  const keychain = new InMemoryKeychain();
  await keychain.setPassword(
    "mcp-onedrive-sharepoint",
    "access_token",
    JSON.stringify({
      accessToken: "expired-token",
      expiresOn: new Date(Date.now() - 60_000).toISOString(),
      account: {
        username: "user@example.com",
        name: "Example User",
        tenantId: "tenant-123",
      },
    }),
  );

  let silentCalls = 0;
  const account = {
    homeAccountId: "home-account-id",
    environment: "login.microsoftonline.com",
    tenantId: "tenant-123",
    username: "user@example.com",
    localAccountId: "local-account-id",
    name: "Example User",
  };

  const fakePca = {
    getTokenCache: () => ({
      getAllAccounts: async () => [account],
      removeAccount: async () => undefined,
    }),
    acquireTokenSilent: async ({ account: requestedAccount }: { account: typeof account }) => {
      silentCalls += 1;
      assert.equal(requestedAccount.username, "user@example.com");

      return {
        accessToken: "fresh-token",
        expiresOn: new Date(Date.now() + 60 * 60 * 1000),
        account,
      };
    },
    acquireTokenByDeviceCode: async () => {
      throw new Error("device code flow should not be required");
    },
  };

  const auth = new MicrosoftGraphAuth(
    { clientId: "client-id", tenantId: "common", scopes: ["User.Read"] },
    { keychain, pca: fakePca as any },
  );

  assert.equal(await auth.isAuthenticated(), true);
  assert.equal(silentCalls, 1);
  assert.equal(await auth.getAccessToken(), "fresh-token");
  assert.equal(silentCalls, 1);

  const updatedCache = await keychain.getPassword(
    "mcp-onedrive-sharepoint",
    "access_token",
  );
  assert.ok(updatedCache);
  assert.match(updatedCache ?? "", /fresh-token/);
});

test("signOut clears both access token and persisted MSAL cache snapshots", async () => {
  const keychain = new InMemoryKeychain();
  await keychain.setPassword(
    "mcp-onedrive-sharepoint",
    "access_token",
    "token-data",
  );
  await keychain.setPassword(
    "mcp-onedrive-sharepoint",
    "msal_token_cache",
    "cache-data",
  );

  const fakePca = {
    getTokenCache: () => ({
      getAllAccounts: async () => [
        {
          username: "user@example.com",
        },
      ],
      removeAccount: async () => undefined,
    }),
    acquireTokenSilent: async () => null,
    acquireTokenByDeviceCode: async () => null,
  };

  const auth = new MicrosoftGraphAuth(
    { clientId: "client-id", tenantId: "common" },
    { keychain, pca: fakePca as any },
  );

  await auth.signOut();

  assert.equal(
    await keychain.getPassword("mcp-onedrive-sharepoint", "access_token"),
    null,
  );
  assert.equal(
    await keychain.getPassword("mcp-onedrive-sharepoint", "msal_token_cache"),
    null,
  );
});
