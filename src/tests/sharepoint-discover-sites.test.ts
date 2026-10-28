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

import { redactOutput, redactDiagnostic } from "../utils/redaction.js";
import { SecurityValidator } from "../utils/security-validator.js";
import { toolErrorResponse } from "../graph/contracts.js";
import { createUserFriendlyError } from "../graph/error-handler.js";

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

for (const search of [undefined, "Finance Area"]) {
  for (const cliSecondPass of [false, true]) {
    test(`round 14 discover_sites two pages / search=${search ?? "wildcard"} / CLI=${cliSecondPass}`, async () => {
      const searchQuery = search ? encodeURIComponent(search) : "*";
      const initialEndpoint = `/sites?search=${searchQuery}`;
      const cursorValue = ["fictional", "page", "cursor"].join("-");
      const cursor = `https://graph.microsoft.com/v1.0/sites?search=${searchQuery}&$top=1&$skiptoken=${cursorValue}%2B%2F%3D`;
      let page = 0;
      const mock = createMockGraphClient({
        get: async () => ({
          success: true,
          data: {
            value: [
              {
                id: `site-${++page}`,
                displayName: "Fictional Site",
                webUrl: "https://contoso.sharepoint.com/sites/fictional",
              },
            ],
            ...(page === 1 ? { "@odata.nextLink": cursor } : {}),
          },
        }),
      });
      __setGraphClientInstanceForTests(
        mock.client as unknown as Parameters<
          typeof __setGraphClientInstanceForTests
        >[0],
      );
      const first = (await handleDiscoverSites({
        search,
        limit: 1,
      })) as ToolEnvelope;
      assert.equal(first.isError, undefined);
      const payload = JSON.parse(
        cliSecondPass
          ? redactOutput(first.content[0].text)
          : first.content[0].text,
      );
      assert.equal(payload.pagination.nextPageToken, cursor);
      assert.equal(payload.pagination.hasMore, true);
      assert.equal(payload.sites[0].id, "site-1");
      const second = (await handleDiscoverSites({
        search,
        limit: 1,
        pageToken: payload.pagination.nextPageToken,
      })) as ToolEnvelope;
      assert.equal(second.isError, undefined);
      const last = JSON.parse(
        cliSecondPass
          ? redactOutput(second.content[0].text)
          : second.content[0].text,
      );
      assert.equal(last.pagination.hasMore, false);
      assert.equal(last.sites[0].id, "site-2");
      assert.equal(mock.methodCalls("get").length, 2);
      assert.equal(mock.methodCalls("get")[0].args[0], initialEndpoint);
      assert.deepEqual(mock.methodCalls("get")[0].args[1], { $top: "1" });
      assert.equal(mock.methodCalls("get")[1].args[0], cursor);
      assert.equal(mock.methodCalls("get")[1].args[1], undefined);
      for (const diagnostic of [cursor, JSON.stringify(payload)]) {
        for (const output of [
          redactDiagnostic(diagnostic),
          SecurityValidator.sanitizeForLogging(diagnostic),
          createUserFriendlyError(new Error(diagnostic)),
          JSON.stringify(
            toolErrorResponse("discover_sites", new Error(diagnostic)),
          ),
        ]) {
          assert.ok(!output.includes(cursorValue));
        }
      }
    });
  }
}
