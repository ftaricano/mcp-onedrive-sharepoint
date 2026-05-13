import test from "node:test";
import assert from "node:assert/strict";

import { __setGraphClientInstanceForTests } from "../graph/client.js";
import { handleDiscoverSites } from "../tools/sharepoint/index.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";
import {
  createMockGraphClient,
  parsePayload,
  type ToolEnvelope,
} from "./helpers/tool-test-helpers.js";

registerGraphClientTestLifecycle();

test("discover_sites uses Graph site search endpoint for explicit searches", async () => {
  const mock = createMockGraphClient({
    get: async (endpoint: string) => {
      if (endpoint === "/sites/root") {
        return {
          success: true,
          data: {
            id: "root-site",
            displayName: "Root Site",
            name: "Root Site",
            webUrl: "https://contoso.sharepoint.com",
            root: {},
          },
        };
      }

      return {
        success: true,
        data: {
          value: [
            {
              id: "site-1",
              displayName: "Financeiro Área",
              name: "Financeiro Área",
              webUrl: "https://contoso.sharepoint.com/sites/financeiro",
            },
          ],
        },
      };
    },
  });

  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleDiscoverSites({
    search: "  Financeiro   Área  ",
    includePersonalSite: true,
    limit: 5,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.search, "Financeiro Área");
  assert.equal(payload.siteCount, 2);
  assert.equal(payload.sites[0].id, "root-site");
  assert.equal(
    mock.methodCalls("get")[0]?.args[0],
    "/sites?search=Financeiro%20%C3%81rea",
  );
  assert.deepEqual(mock.methodCalls("get")[0]?.args[1], { $top: "5" });
});

test("discover_sites falls back to wildcard Graph search when no search term is provided", async () => {
  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: {
        value: [
          {
            id: "site-1",
            displayName: "Finance",
            name: "Finance",
            webUrl: "https://contoso.sharepoint.com/sites/finance",
          },
        ],
      },
    }),
  });

  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleDiscoverSites({ limit: 10 })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.search, "all sites");
  assert.equal(payload.siteCount, 1);
  assert.equal(mock.methodCalls("get")[0]?.args[0], "/sites?search=*");
  assert.deepEqual(mock.methodCalls("get")[0]?.args[1], { $top: "10" });
});
