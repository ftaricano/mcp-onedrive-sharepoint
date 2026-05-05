import assert from "node:assert/strict";
import test from "node:test";
import type { ICachePlugin, PublicClientApplication } from "@azure/msal-node";

import {
  MicrosoftGraphAuth,
  createKeychainMsalCachePlugin,
} from "../auth/microsoft-graph-auth.js";

type BeforeCacheContext = Parameters<ICachePlugin["beforeCacheAccess"]>[0];
type AfterCacheContext = Parameters<ICachePlugin["afterCacheAccess"]>[0];

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

class WriteFailingKeychain extends InMemoryKeychain {
  async setPassword(): Promise<void> {
    throw new Error("keychain write failed");
  }
}

class ReadFailingKeychain extends InMemoryKeychain {
  async getPassword(): Promise<string | null> {
    throw new Error("keychain read failed");
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
  } as unknown as AfterCacheContext);

  await plugin.beforeCacheAccess({
    tokenCache: {
      serialize: () => "",
      deserialize: (snapshot: string) => {
        restoredSnapshot = snapshot;
      },
    },
  } as unknown as BeforeCacheContext);

  assert.match(restoredSnapshot, /user@example.com/);
});

test("MSAL cache plugin falls back to file storage when keychain writes fail", async (t) => {
  t.mock.method(console, "warn", () => undefined);

  const keychain = new WriteFailingKeychain();
  const fallbackStore = new InMemoryKeychain();
  const plugin = createKeychainMsalCachePlugin(
    keychain,
    "mcp-onedrive-sharepoint",
    "msal_token_cache",
    fallbackStore,
  );

  const serializedCache = JSON.stringify({
    Account: { abc: { username: "user@example.com" } },
  });

  await plugin.afterCacheAccess({
    cacheHasChanged: true,
    tokenCache: {
      serialize: () => serializedCache,
      deserialize: () => undefined,
    },
  } as unknown as AfterCacheContext);

  assert.equal(
    await fallbackStore.getPassword(
      "mcp-onedrive-sharepoint",
      "msal_token_cache",
    ),
    serializedCache,
  );
});

test("MSAL cache plugin reads file fallback when keychain reads fail", async (t) => {
  t.mock.method(console, "warn", () => undefined);

  const keychain = new ReadFailingKeychain();
  const fallbackStore = new InMemoryKeychain();
  const serializedCache = JSON.stringify({
    Account: { abc: { username: "fallback@example.com" } },
  });
  await fallbackStore.setPassword(
    "mcp-onedrive-sharepoint",
    "msal_token_cache",
    serializedCache,
  );

  const plugin = createKeychainMsalCachePlugin(
    keychain,
    "mcp-onedrive-sharepoint",
    "msal_token_cache",
    fallbackStore,
  );
  let restoredSnapshot = "";

  await plugin.beforeCacheAccess({
    tokenCache: {
      serialize: () => "",
      deserialize: (snapshot: string) => {
        restoredSnapshot = snapshot;
      },
    },
  } as unknown as BeforeCacheContext);

  assert.match(restoredSnapshot, /fallback@example.com/);
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
    { keychain, pca: fakePca as unknown as PublicClientApplication },
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

test("access token cache reads from file fallback when keychain reads fail", async (t) => {
  t.mock.method(console, "warn", () => undefined);

  const keychain = new ReadFailingKeychain();
  const fallbackStore = new InMemoryKeychain();
  await fallbackStore.setPassword(
    "mcp-onedrive-sharepoint",
    "access_token",
    JSON.stringify({
      accessToken: "fallback-token",
      expiresOn: new Date(Date.now() + 60 * 60 * 1000).toISOString(),
      account: {
        username: "user@example.com",
        name: "Example User",
        tenantId: "tenant-123",
      },
    }),
  );

  const fakePca = {
    getTokenCache: () => ({
      getAllAccounts: async () => [],
      removeAccount: async () => undefined,
    }),
    acquireTokenSilent: async () => {
      throw new Error("silent auth should not be required");
    },
    acquireTokenByDeviceCode: async () => {
      throw new Error("device code flow should not be required");
    },
  };

  const auth = new MicrosoftGraphAuth(
    { clientId: "client-id", tenantId: "common", scopes: ["User.Read"] },
    {
      keychain,
      fallbackStore,
      pca: fakePca as unknown as PublicClientApplication,
    },
  );

  assert.equal(await auth.getAccessToken(), "fallback-token");
});

test("access token cache writes to file fallback when keychain writes fail", async (t) => {
  t.mock.method(console, "warn", () => undefined);

  const keychain = new WriteFailingKeychain();
  const fallbackStore = new InMemoryKeychain();
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
    acquireTokenSilent: async () => ({
      accessToken: "fresh-fallback-token",
      expiresOn: new Date(Date.now() + 60 * 60 * 1000),
      account,
    }),
    acquireTokenByDeviceCode: async () => {
      throw new Error("device code flow should not be required");
    },
  };

  const auth = new MicrosoftGraphAuth(
    { clientId: "client-id", tenantId: "common", scopes: ["User.Read"] },
    {
      keychain,
      fallbackStore,
      pca: fakePca as unknown as PublicClientApplication,
    },
  );

  assert.equal(await auth.getAccessToken(), "fresh-fallback-token");

  const fallbackToken = await fallbackStore.getPassword(
    "mcp-onedrive-sharepoint",
    "access_token",
  );
  assert.ok(fallbackToken);
  assert.match(fallbackToken ?? "", /fresh-fallback-token/);
});

test("signOut clears both access token and persisted MSAL cache snapshots", async () => {
  const keychain = new InMemoryKeychain();
  const fallbackStore = new InMemoryKeychain();
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
  await fallbackStore.setPassword(
    "mcp-onedrive-sharepoint",
    "access_token",
    "fallback-token-data",
  );
  await fallbackStore.setPassword(
    "mcp-onedrive-sharepoint",
    "msal_token_cache",
    "fallback-cache-data",
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
    {
      keychain,
      fallbackStore,
      pca: fakePca as unknown as PublicClientApplication,
    },
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
  assert.equal(
    await fallbackStore.getPassword("mcp-onedrive-sharepoint", "access_token"),
    null,
  );
  assert.equal(
    await fallbackStore.getPassword(
      "mcp-onedrive-sharepoint",
      "msal_token_cache",
    ),
    null,
  );
});
