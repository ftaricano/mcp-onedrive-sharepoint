import test from "node:test";
import assert from "node:assert/strict";

import { buildUrl } from "../config/endpoints.js";
import {
  extractPaginatedResult,
  jsonTextResponse,
  toolErrorResponse,
} from "../graph/contracts.js";
import {
  buildDriveChildrenEndpoint,
  buildDriveItemEndpoint,
  buildDriveSearchEndpoint,
  describeDriveTarget,
} from "../graph/resource-resolver.js";
import {
  GraphApiError,
  createUserFriendlyError,
} from "../graph/error-handler.js";
import { GraphClient } from "../graph/client.js";
import { metadataCache } from "../utils/cache-manager.js";

test.afterEach(() => {
  metadataCache.clear();
});

test("buildUrl appends query params without duplicating path placeholders", () => {
  const url = buildUrl(
    "/sites/{siteId}/lists",
    {
      siteId: "site-123",
      $top: "25",
      $orderby: "displayName",
      $filter: "name ne 'Archive'",
    },
    false,
  );

  assert.equal(
    url,
    "/sites/site-123/lists?%24top=25&%24orderby=displayName&%24filter=name+ne+%27Archive%27",
  );
});

test("extractPaginatedResult returns items and next page metadata", () => {
  const result = extractPaginatedResult(
    {
      value: [{ id: "1" }, { id: "2" }],
      "@odata.nextLink":
        "https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=abc",
      "@odata.count": 12,
    },
    2,
  );

  assert.equal(result.items.length, 2);
  assert.equal(result.pagination.returned, 2);
  assert.equal(result.pagination.totalCount, 12);
  assert.equal(result.pagination.hasMore, true);
  assert.match(result.pagination.nextPageToken ?? "", /skiptoken=abc/);
});

test("extractPaginatedResult rejects Graph error payloads instead of masking them", () => {
  assert.throws(
    () =>
      extractPaginatedResult({
        error: { code: "accessDenied", message: "Access denied" },
      }),
    (error: unknown) => {
      assert.ok(error instanceof GraphApiError);
      assert.equal(error.code, "accessDenied");
      return true;
    },
  );
});

test("resource resolver supports me, site and drive scopes", () => {
  assert.equal(
    buildDriveChildrenEndpoint({ path: "/Documents/Plans" }),
    "/me/drive/root:/Documents/Plans:/children",
  );
  assert.equal(
    buildDriveChildrenEndpoint({ siteId: "site-123", path: "Docs" }),
    "/sites/site-123/drive/root:/Docs:/children",
  );
  assert.equal(
    buildDriveChildrenEndpoint({ driveId: "drive-123" }),
    "/drives/drive-123/root/children",
  );
  assert.equal(
    buildDriveItemEndpoint(
      { driveId: "drive-123", itemId: "item-456" },
      "/content",
    ),
    "/drives/drive-123/items/item-456/content",
  );
  assert.equal(
    buildDriveItemEndpoint({
      siteId: "site-123",
      itemPath: "/Docs/report.docx",
    }),
    "/sites/site-123/drive/root:/Docs/report.docx:",
  );
  assert.equal(
    buildDriveSearchEndpoint({ driveId: "drive-123" }, "budget 2026"),
    "/drives/drive-123/root/search(q='budget%202026')",
  );
  // OData strings escape `'` by doubling it. `encodeURIComponent` leaves
  // `'` untouched (RFC 3986 unreserved per spec), so the doubled quotes
  // reach Graph as literal `''` and are parsed as an escaped single quote.
  assert.equal(
    buildDriveSearchEndpoint({ driveId: "drive-123" }, "can't break"),
    "/drives/drive-123/root/search(q='can''t%20break')",
  );
  assert.equal(describeDriveTarget({}), "me");
  assert.equal(describeDriveTarget({ siteId: "site-123" }), "site:site-123");
  assert.equal(
    describeDriveTarget({ driveId: "drive-123" }),
    "drive:drive-123",
  );
});

test("JSON and error envelopes are MCP-compatible text responses", () => {
  const success = jsonTextResponse({ ok: true, nested: { value: 1 } });
  assert.equal(success.content.length, 1);
  assert.equal(success.content[0].type, "text");
  assert.match(success.content[0].text, /"ok": true/);

  const error = toolErrorResponse(
    "list_files",
    new GraphApiError({
      error: { code: "InvalidAuthenticationToken", message: "expired" },
    }),
  );
  assert.equal(error.isError, true);
  // First block: friendly summary text (backward-compat with plain-text readers).
  assert.equal(error.content.length, 2);
  assert.match(error.content[0].text, /Error in list_files/);
  assert.match(error.content[0].text, /Authentication Error/);

  // Second block: structured JSON so callers (LLMs) can decide retry vs bail.
  const parsed = JSON.parse(error.content[1].text);
  assert.equal(parsed.error.category, "Authentication");
  assert.equal(parsed.error.code, "InvalidAuthenticationToken");
  assert.equal(parsed.error.retryable, false);
  assert.match(parsed.summary, /Error in list_files/);
});

test("toolErrorResponse flags throttling errors as retryable", () => {
  const error = toolErrorResponse(
    "list_files",
    new GraphApiError(
      { error: { code: "TooManyRequests", message: "slow down" } },
      undefined,
      429,
    ),
  );
  const parsed = JSON.parse(error.content[1].text);
  assert.equal(parsed.error.category, "Throttling");
  assert.equal(parsed.error.retryable, true);
  assert.equal(parsed.error.statusCode, 429);
});

test("GraphApiError produces actionable messages", () => {
  const error = new GraphApiError(
    { error: { code: "ItemNotFound", message: "missing file" } },
    "GET /me/drive/items/x",
    404,
  );
  assert.equal(error.category, "NotFound");
  assert.equal(error.isRetryable, false);
  assert.match(createUserFriendlyError(error), /Suggested Action:/);
});

test("GraphApiError categorizes lowercase accessDenied as a permission error", () => {
  const error = new GraphApiError({
    error: { code: "accessDenied", message: "Access denied" },
  });

  assert.equal(error.category, "Permission");
  assert.equal(error.isRetryable, false);
});

test("GraphClient.get throws when Graph returns a fulfilled response containing an error payload", async () => {
  const client = new GraphClient();
  const endpoint = "/sites/site-1/lists";
  const params = { $top: "5" };

  (client as any).axios = {
    get: async () => ({
      data: {
        error: { code: "accessDenied", message: "Access denied" },
      },
      headers: {},
    }),
  };

  await assert.rejects(client.get(endpoint, params), (error: unknown) => {
    assert.ok(error instanceof GraphApiError);
    assert.equal(error.code, "accessDenied");
    return true;
  });

  assert.equal(metadataCache.get(`${endpoint}:${JSON.stringify(params)}`), null);
});

test("GraphClient.get preserves success payloads", async () => {
  const client = new GraphClient();

  (client as any).axios = {
    get: async () => ({
      data: {
        value: [{ id: "item-1" }],
      },
      headers: {},
    }),
  };

  const response = await client.get<{ id: string }>("/sites/site-1/lists", {
    $top: "1",
  });

  assert.equal(response.success, true);
  assert.deepEqual(response.data, { value: [{ id: "item-1" }] });
});

test("GraphClient.getAllPages truncates at maxItems and emits metadata", async () => {
  const client = new GraphClient();

  (client as any).axios = {
    get: async () => ({
      data: {
        value: [{ id: "a" }, { id: "b" }, { id: "c" }],
        "@odata.nextLink":
          "https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=x",
      },
      headers: {},
    }),
  };

  const response = await client.getAllPages<{ id: string }>(
    "/me/drive/root/children",
    undefined,
    { maxItems: 2 },
  );

  assert.equal(response.success, true);
  assert.equal(response.data?.length, 2);
  assert.equal(response.metadata?.truncated, true);
  assert.equal(response.metadata?.truncationReason, "maxItems");
  assert.match(response.metadata?.nextPageToken ?? "", /skiptoken=x/);
});

test("GraphClient.getAllPages truncates at maxPages and emits metadata", async () => {
  const client = new GraphClient();

  (client as any).axios = {
    get: async () => ({
      data: {
        value: [{ id: "a" }],
        "@odata.nextLink":
          "https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=x",
      },
      headers: {},
    }),
  };

  const response = await client.getAllPages<{ id: string }>(
    "/me/drive/root/children",
    undefined,
    { maxPages: 2 },
  );

  assert.equal(response.success, true);
  assert.equal(response.data?.length, 2);
  assert.equal(response.metadata?.truncated, true);
  assert.equal(response.metadata?.truncationReason, "maxPages");
});

test("GraphClient.getAllPages returns full result when within caps", async () => {
  const client = new GraphClient();

  const pages = [
    {
      value: [{ id: "a" }, { id: "b" }],
      "@odata.nextLink":
        "https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=p2",
    },
    {
      value: [{ id: "c" }],
    },
  ];
  let call = 0;

  (client as any).axios = {
    get: async () => ({
      data: pages[call++],
      headers: {},
    }),
  };

  const response = await client.getAllPages<{ id: string }>(
    "/me/drive/root/children",
    undefined,
    { maxItems: 100, maxPages: 10 },
  );

  assert.equal(response.success, true);
  assert.deepEqual(response.data, [{ id: "a" }, { id: "b" }, { id: "c" }]);
  assert.equal(call, 2);
});
