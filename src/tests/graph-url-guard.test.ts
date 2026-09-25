import test, { afterEach } from "node:test";
import assert from "node:assert/strict";
import type { AxiosInstance, InternalAxiosRequestConfig } from "axios";

import {
  __setAuthInstanceForTests,
  type MicrosoftGraphAuth,
} from "../auth/microsoft-graph-auth.js";
import { GraphApiError } from "../graph/error-handler.js";
import {
  GraphClient,
  __setGraphClientInstanceForTests,
} from "../graph/client.js";
import { assertGraphRequestUrl } from "../graph/url-guard.js";
import { handleListFiles, handleSearchFiles } from "../tools/files/index.js";
import {
  handleDiscoverSites,
  handleListItems,
  handleListSiteLists,
} from "../tools/sharepoint/index.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";
import type { ToolEnvelope } from "./helpers/tool-test-helpers.js";

/**
 * Regression guard: a caller-supplied `pageToken` (or a Graph-returned
 * `@odata.nextLink`) must never make the client send the bearer token to a
 * host other than Microsoft Graph.
 *
 * These tests use the real GraphClient with its real request interceptor and
 * swap only the axios adapter (the transport) for a recorder, so "no request"
 * means nothing reached the network layer, with or without Authorization.
 */

registerGraphClientTestLifecycle();

afterEach(() => {
  __setAuthInstanceForTests(null);
});

type RecordedRequest = { url: string; authorization: unknown };

function createRecordingClient(
  respond: (url: string) => unknown = () => ({ value: [] }),
) {
  const requests: RecordedRequest[] = [];
  let tokenFetches = 0;

  __setAuthInstanceForTests({
    getAccessToken: async () => {
      tokenFetches++;
      return "test-token";
    },
    getCurrentUser: async () => null,
  } as unknown as MicrosoftGraphAuth);

  const client = new GraphClient();
  const instance = (client as unknown as { axios: AxiosInstance }).axios;
  instance.defaults.adapter = async (config: InternalAxiosRequestConfig) => {
    const url = instance.getUri(config);
    requests.push({ url, authorization: config.headers?.Authorization });
    return {
      data: respond(url),
      status: 200,
      statusText: "OK",
      headers: {},
      config,
      request: {},
    };
  };

  __setGraphClientInstanceForTests(client);

  return {
    client,
    requests,
    get tokenFetches() {
      return tokenFetches;
    },
  };
}

const PAGINATED_TOOLS: Array<{
  name: string;
  run: (pageToken: string) => Promise<unknown>;
  legitimateNextLink: string;
}> = [
  {
    name: "list_files",
    run: (pageToken) => handleListFiles({ pageToken }),
    legitimateNextLink:
      "https://graph.microsoft.com/v1.0/me/drive/root/children?$top=100&$skiptoken=page2",
  },
  {
    name: "search_files",
    run: (pageToken) => handleSearchFiles({ query: "report", pageToken }),
    legitimateNextLink:
      "https://graph.microsoft.com/v1.0/me/drive/root/search(q='report')?$skiptoken=page2",
  },
  {
    name: "discover_sites",
    run: (pageToken) => handleDiscoverSites({ pageToken }),
    legitimateNextLink:
      "https://graph.microsoft.com/v1.0/sites?search=*&$skiptoken=page2",
  },
  {
    name: "list_site_lists",
    run: (pageToken) => handleListSiteLists({ siteId: "site-1", pageToken }),
    legitimateNextLink:
      "https://graph.microsoft.com/v1.0/sites/site-1/lists?$skiptoken=page2",
  },
  {
    name: "list_items",
    run: (pageToken) =>
      handleListItems({ siteId: "site-1", listId: "list-1", pageToken }),
    legitimateNextLink:
      "https://graph.microsoft.com/beta/sites/site-1/lists/list-1/items?$skiptoken=page2",
  },
];

const FOREIGN_PAGE_TOKENS: Array<[label: string, pageToken: string]> = [
  ["foreign https host", "https://attacker.example/v1.0/me/drive"],
  ["loopback http host", "http://127.0.0.1:8080/steal?next=1"],
  ["plain http to Graph", "http://graph.microsoft.com/v1.0/me/drive"],
  [
    "userinfo before foreign host",
    "https://graph.microsoft.com@attacker.example/v1.0/me",
  ],
  ["userinfo on Graph host", "https://user:pass@graph.microsoft.com/v1.0/me"],
  [
    "Graph host as subdomain",
    "https://graph.microsoft.com.attacker.example/v1.0/me",
  ],
  ["lookalike host", "https://graph-microsoft.com/v1.0/me"],
  ["sibling subdomain", "https://evil.graph.microsoft.com/v1.0/me"],
  ["non-default port", "https://graph.microsoft.com:8443/v1.0/me"],
  ["protocol-relative", "//attacker.example/v1.0/me"],
];

for (const tool of PAGINATED_TOOLS) {
  for (const [label, pageToken] of FOREIGN_PAGE_TOKENS) {
    test(`${tool.name} refuses a ${label} pageToken without sending any request`, async () => {
      const recorder = createRecordingClient();

      const response = (await tool.run(pageToken)) as ToolEnvelope;

      assert.equal(response.isError, true);
      assert.match(
        response.content[0].text,
        /Refusing to send an authenticated request/,
      );
      assert.deepEqual(recorder.requests, []);
      assert.equal(recorder.tokenFetches, 0);
    });
  }

  test(`${tool.name} follows a legitimate Graph nextLink with the bearer token`, async () => {
    const recorder = createRecordingClient();

    const response = (await tool.run(tool.legitimateNextLink)) as ToolEnvelope;

    assert.equal(response.isError, undefined);
    assert.equal(recorder.requests.length, 1);
    assert.equal(recorder.requests[0].url, tool.legitimateNextLink);
    assert.equal(recorder.requests[0].authorization, "Bearer test-token");
  });
}

test("a relative pageToken keeps resolving against the Graph base URL", async () => {
  const recorder = createRecordingClient();

  const response = (await handleListSiteLists({
    siteId: "site-1",
    pageToken: "/sites/site-1/lists?$skiptoken=page2",
  })) as ToolEnvelope;

  assert.equal(response.isError, undefined);
  assert.deepEqual(recorder.requests, [
    {
      url: "https://graph.microsoft.com/v1.0/sites/site-1/lists?$skiptoken=page2",
      authorization: "Bearer test-token",
    },
  ]);
});

test("getAllPages stops before following a foreign @odata.nextLink", async () => {
  const recorder = createRecordingClient(() => ({
    value: [{ id: "item-1" }],
    "@odata.nextLink":
      "https://attacker.example/v1.0/me/drive/root/children?page=2",
  }));

  await assert.rejects(
    recorder.client.getAllPages("/me/drive/root/children"),
    (error: unknown) =>
      error instanceof GraphApiError &&
      /Refusing to send an authenticated request/.test(error.message),
  );
  assert.deepEqual(
    recorder.requests.map((request) => request.url),
    ["https://graph.microsoft.com/v1.0/me/drive/root/children"],
  );
  assert.equal(recorder.tokenFetches, 1);
});

test("the request URL guard is not retried", async () => {
  const recorder = createRecordingClient();

  await assert.rejects(
    recorder.client.get("https://attacker.example/v1.0/me"),
    (error: unknown) =>
      error instanceof GraphApiError &&
      error.category === "Validation" &&
      error.isRetryable === false,
  );
  assert.deepEqual(recorder.requests, []);
});

test("assertGraphRequestUrl accepts Graph v1.0 and beta URLs", () => {
  for (const url of [
    "https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=abc",
    "https://graph.microsoft.com/beta/sites/site-1/lists",
    "https://GRAPH.MICROSOFT.COM/v1.0/me",
    "https://graph.microsoft.com:443/v1.0/me",
  ]) {
    assert.doesNotThrow(() => assertGraphRequestUrl(url), url);
  }
});

test("assertGraphRequestUrl rejects foreign, downgraded, credentialed, ported and off-version URLs", () => {
  for (const url of [
    ...FOREIGN_PAGE_TOKENS.map(([, pageToken]) => pageToken),
    "https://graph.microsoft.com./v1.0/me",
    "https://graph.microsoft.com/v2.0/me",
    "https://graph.microsoft.com/me",
    "ftp://graph.microsoft.com/v1.0/me",
    "not a url",
    "",
  ]) {
    assert.throws(
      () => assertGraphRequestUrl(url),
      (error: unknown) =>
        error instanceof GraphApiError && error.category === "Validation",
      url,
    );
  }
});
