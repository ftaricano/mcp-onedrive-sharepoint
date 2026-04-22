import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { __setGraphClientInstanceForTests } from "../graph/client.js";
import {
  __resetKnownSitesForTests,
  __setKnownSitesForTests,
} from "../sharepoint/site-resolver.js";
import { handleListFiles, handleUploadFile } from "../tools/files/index.js";
import {
  handleResolveSite,
  handleListSiteLists,
} from "../tools/sharepoint/index.js";
import { handleListDrives } from "../tools/utils/index.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";
import {
  createMockGraphClient,
  parsePayload,
  type ToolEnvelope,
} from "./helpers/tool-test-helpers.js";

registerGraphClientTestLifecycle();

const PRIMARY_SITE_ID =
  "example.sharepoint.com,00000000-0000-0000-0000-000000000001,11111111-1111-1111-1111-111111111111";
const PRIMARY_DRIVE_ID = "b!TESTDRIVEPRIMARY0000000000000000";
const SECONDARY_SITE_ID =
  "example.sharepoint.com,00000000-0000-0000-0000-000000000002,22222222-2222-2222-2222-222222222222";
const SECONDARY_DRIVE_ID = "b!TESTDRIVESECONDARY0000000000000";

test.beforeEach(() => {
  __setKnownSitesForTests([
    {
      key: "primary",
      name: "Primary",
      siteId: PRIMARY_SITE_ID,
      siteUrl: "https://example.sharepoint.com/sites/Primary",
      driveId: PRIMARY_DRIVE_ID,
      aliases: ["primary", "primary-site", "/sites/Primary"],
    },
    {
      key: "secondary",
      name: "Secondary",
      siteId: SECONDARY_SITE_ID,
      siteUrl: "https://example.sharepoint.com/sites/Secondary",
      driveId: SECONDARY_DRIVE_ID,
      aliases: ["secondary", "/sites/Secondary"],
    },
  ]);
});

test.afterEach(() => {
  __resetKnownSitesForTests();
});

test("resolve_site resolves a canonical alias locally without Graph lookup", async () => {
  const mock = createMockGraphClient();
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleResolveSite({
    site: "primary",
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.resolved, true);
  assert.equal(payload.site.key, "primary");
  assert.equal(payload.site.siteId, PRIMARY_SITE_ID);
  assert.equal(payload.site.driveId, PRIMARY_DRIVE_ID);
  assert.equal(mock.methodCalls("get").length, 0);
});

test("list_files accepts a site alias and resolves its drive without a raw siteId", async () => {
  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: {
        value: [
          {
            id: "item-1",
            name: "Budget.xlsx",
            size: 128,
            file: {
              mimeType:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
            lastModifiedDateTime: "2026-04-15T12:00:00.000Z",
            webUrl:
              "https://example.sharepoint.com/sites/Primary/Docs/Budget.xlsx",
            parentReference: {
              path: "/drive/root:/Docs",
              driveId: PRIMARY_DRIVE_ID,
            },
          },
        ],
      },
    }),
  });

  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleListFiles({
    site: "primary",
    path: "/Docs",
    limit: 10,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(
    mock.methodCalls("get")[0]?.args[0],
    `/drives/${PRIMARY_DRIVE_ID}/root:/Docs:/children`,
  );
  assert.equal(payload.target, `drive:${PRIMARY_DRIVE_ID}`);
  assert.equal(payload.site.key, "primary");
  assert.equal(payload.items[0].name, "Budget.xlsx");
});

test("list_files refuses unresolved SharePoint aliases instead of falling back to personal OneDrive", async () => {
  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: { value: [] },
    }),
  });

  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleListFiles({
    site: "missing-sharepoint-site",
    path: "/Docs",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.match(
    response.content[0].text,
    /Refusing to fall back to personal OneDrive/,
  );
  assert.equal(mock.methodCalls("get").length, 0);
});

test("list_site_lists accepts a canonical siteUrl and resolves the siteId from the registry", async () => {
  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: {
        value: [
          {
            id: "list-1",
            name: "Reports",
            displayName: "Reports",
            list: {
              hidden: false,
              template: "documentLibrary",
              contentTypesEnabled: true,
            },
            columns: [],
            contentTypes: [],
          },
        ],
      },
    }),
  });

  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleListSiteLists({
    siteUrl: "https://example.sharepoint.com/sites/Secondary",
    limit: 5,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(
    mock.methodCalls("get")[0]?.args[0],
    `/sites/${SECONDARY_SITE_ID}/lists`,
  );
  assert.equal(payload.site.key, "secondary");
  assert.equal(payload.siteId, SECONDARY_SITE_ID);
  assert.equal(payload.lists[0].id, "list-1");
});

test("list_drives can target a canonical site drive directly", async () => {
  const mock = createMockGraphClient({
    get: async (endpoint: string) => ({
      success: true,
      data: {
        id: endpoint.replace("/drives/", ""),
        name: "Documents",
        driveType: "documentLibrary",
        webUrl: "https://example.sharepoint.com/sites/Secondary/Shared%20Documents",
        quota: {
          total: 100,
          used: 40,
          remaining: 60,
          state: "normal",
        },
      },
    }),
  });

  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleListDrives({
    site: "secondary",
    includeQuota: true,
  })) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(
    mock.methodCalls("get")[0]?.args[0],
    `/drives/${SECONDARY_DRIVE_ID}`,
  );
  assert.equal(payload.driveCount, 1);
  assert.equal(payload.site.key, "secondary");
  assert.equal(payload.drives[0].id, SECONDARY_DRIVE_ID);
});

test("upload_file prefers the resolved SharePoint driveId for canonical site aliases", async () => {
  const tempDir = mkdtempSync(path.join(tmpdir(), "sp-upload-resolved-"));
  const localPath = path.join(tempDir, "report.txt");
  writeFileSync(localPath, "hello");

  const mock = createMockGraphClient({
    uploadFile: async (
      _endpoint: string,
      _uploadPath: string,
      fileName: string,
    ) => ({
      success: true,
      data: {
        id: "file-1",
        name: fileName,
        size: 5,
        webUrl: `https://example.sharepoint.com/sites/Primary/${fileName}`,
        lastModifiedDateTime: "2026-04-15T12:00:00.000Z",
      },
    }),
  });

  __setGraphClientInstanceForTests(mock.client as any);

  try {
    const response = (await handleUploadFile({
      localPath,
      remotePath: "report.txt",
      site: "primary",
    })) as ToolEnvelope;
    const payload = parsePayload(response);

    assert.equal(response.isError, undefined);
    assert.equal(
      mock.methodCalls("uploadFile")[0]?.args[0],
      `/drives/${PRIMARY_DRIVE_ID}/root:/report.txt:/content`,
    );
    assert.equal(payload.target, `drive:${PRIMARY_DRIVE_ID}`);
    assert.equal(payload.site.key, "primary");
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
});

test("upload_file honors explicit driveId during folder checks and upload", async () => {
  const tempDir = mkdtempSync(path.join(tmpdir(), "sp-upload-explicit-"));
  const localPath = path.join(tempDir, "report.txt");
  writeFileSync(localPath, "hello");

  const mock = createMockGraphClient({
    get: async () => ({
      success: true,
      data: {
        id: "folder-1",
        name: "safe",
        folder: {},
      },
    }),
    uploadFile: async (
      _endpoint: string,
      _uploadPath: string,
      fileName: string,
    ) => ({
      success: true,
      data: {
        id: "file-2",
        name: fileName,
        size: 5,
        webUrl: `https://example.invalid/${fileName}`,
        lastModifiedDateTime: "2026-04-15T12:00:00.000Z",
      },
    }),
  });

  __setGraphClientInstanceForTests(mock.client as any);

  try {
    const response = (await handleUploadFile({
      localPath,
      remotePath: "safe/report.txt",
      driveId: "drive-explicit-123",
    })) as ToolEnvelope;
    const payload = parsePayload(response);

    assert.equal(response.isError, undefined);
    assert.equal(
      mock.methodCalls("get")[0]?.args[0],
      "/drives/drive-explicit-123/root:/safe",
    );
    assert.equal(
      mock.methodCalls("uploadFile")[0]?.args[0],
      "/drives/drive-explicit-123/root:/safe/report.txt:/content",
    );
    assert.equal(payload.target, "drive:drive-explicit-123");
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
});
