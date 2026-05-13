import { test } from "node:test";
import assert from "node:assert/strict";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  writeFileSync,
} from "node:fs";
import * as path from "node:path";

import { __setGraphClientInstanceForTests } from "../graph/client.js";
import {
  handleBatchFileOperations,
  handleSyncFolder,
} from "../tools/advanced/sync.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";
import {
  cleanupTempDir,
  createMockGraphClient,
  createTempDir,
  parsePayload,
  writeFileWithMtime,
  type ToolEnvelope,
} from "./helpers/tool-test-helpers.js";

registerGraphClientTestLifecycle();

test("sync_folder covers upload conflict resolution branches", async () => {
  const scenarios = [
    {
      name: "local uploads even when remote exists",
      conflictResolution: "local",
      remoteModified: "2026-04-14T18:00:00.000Z",
      localModified: new Date("2026-04-14T10:00:00.000Z"),
      expectUpload: true,
      expectConflict: false,
    },
    {
      name: "remote skips upload when remote wins",
      conflictResolution: "remote",
      remoteModified: "2026-04-14T18:00:00.000Z",
      localModified: new Date("2026-04-14T20:00:00.000Z"),
      expectUpload: false,
      expectConflict: false,
    },
    {
      name: "newer uploads when local file is newer",
      conflictResolution: "newer",
      remoteModified: "2026-04-14T18:00:00.000Z",
      localModified: new Date("2026-04-14T20:00:00.000Z"),
      expectUpload: true,
      expectConflict: false,
    },
    {
      name: "rename records conflict instead of uploading",
      conflictResolution: "rename",
      remoteModified: "2026-04-14T18:00:00.000Z",
      localModified: new Date("2026-04-14T20:00:00.000Z"),
      expectUpload: false,
      expectConflict: true,
    },
  ] as const;

  for (const scenario of scenarios) {
    const localPath = createTempDir("sync-upload-branches-");

    try {
      const localFilePath = path.join(localPath, "report.txt");
      writeFileWithMtime(
        localFilePath,
        "local-content",
        scenario.localModified,
      );

      const mock = createMockGraphClient({
        get: async () => ({
          success: true,
          data: {
            value: [
              {
                id: "remote-1",
                name: "report.txt",
                size: 10,
                file: { mimeType: "text/plain" },
                lastModifiedDateTime: scenario.remoteModified,
              },
            ],
          },
        }),
        uploadFile: async () => ({ success: true, data: { id: "uploaded-1" } }),
      });
      __setGraphClientInstanceForTests(mock.client as any);

      const response = (await handleSyncFolder({
        localPath,
        remotePath: "Docs",
        direction: "upload",
        conflictResolution: scenario.conflictResolution,
      })) as ToolEnvelope;
      const payload = parsePayload<any>(response);

      assert.equal(response.isError, undefined, scenario.name);
      assert.equal(
        payload.summary.uploaded,
        scenario.expectUpload ? 1 : 0,
        scenario.name,
      );
      assert.equal(
        payload.summary.conflicts,
        scenario.expectConflict ? 1 : 0,
        scenario.name,
      );
      assert.equal(
        mock.methodCalls("uploadFile").length,
        scenario.expectUpload ? 1 : 0,
        scenario.name,
      );

      if (scenario.expectUpload) {
        const [uploadEndpoint, uploadPath, uploadName, options] =
          mock.methodCalls("uploadFile")[0].args;
        assert.equal(
          uploadEndpoint,
          "/me/drive/root:/Docs/report.txt:/content",
        );
        assert.equal(uploadPath, localFilePath);
        assert.equal(uploadName, "report.txt");
        assert.deepEqual(options, { conflictBehavior: "replace" });
      }

      if (scenario.expectConflict) {
        assert.match(
          payload.details.conflicts[0].newName,
          /^report_local_\d+\.txt$/,
        );
      }
    } finally {
      cleanupTempDir(localPath);
    }
  }
});

test("sync_folder covers download conflict resolution branches", async () => {
  const scenarios = [
    {
      name: "local skips download when local wins",
      conflictResolution: "local",
      remoteModified: "2026-04-14T20:00:00.000Z",
      localModified: new Date("2026-04-14T10:00:00.000Z"),
      expectDownload: false,
      expectConflict: false,
    },
    {
      name: "remote downloads even when local exists",
      conflictResolution: "remote",
      remoteModified: "2026-04-14T18:00:00.000Z",
      localModified: new Date("2026-04-14T20:00:00.000Z"),
      expectDownload: true,
      expectConflict: false,
    },
    {
      name: "newer downloads when remote file is newer",
      conflictResolution: "newer",
      remoteModified: "2026-04-14T20:00:00.000Z",
      localModified: new Date("2026-04-14T10:00:00.000Z"),
      expectDownload: true,
      expectConflict: false,
    },
    {
      name: "rename records conflict instead of overwriting local file",
      conflictResolution: "rename",
      remoteModified: "2026-04-14T20:00:00.000Z",
      localModified: new Date("2026-04-14T10:00:00.000Z"),
      expectDownload: false,
      expectConflict: true,
    },
  ] as const;

  for (const scenario of scenarios) {
    const localPath = createTempDir("sync-download-branches-");

    try {
      const localFilePath = path.join(localPath, "report.txt");
      writeFileWithMtime(localFilePath, "old-local", scenario.localModified);

      const mock = createMockGraphClient({
        get: async () => ({
          success: true,
          data: {
            value: [
              {
                id: "remote-1",
                name: "report.txt",
                size: 17,
                file: { mimeType: "text/plain" },
                lastModifiedDateTime: scenario.remoteModified,
              },
            ],
          },
        }),
        downloadFile: async () => ({
          success: true,
          data: Buffer.from("remote-buffer"),
        }),
      });
      __setGraphClientInstanceForTests(mock.client as any);

      const response = (await handleSyncFolder({
        localPath,
        remotePath: "Docs",
        direction: "download",
        conflictResolution: scenario.conflictResolution,
      })) as ToolEnvelope;
      const payload = parsePayload<any>(response);

      assert.equal(response.isError, undefined, scenario.name);
      assert.equal(
        payload.summary.downloaded,
        scenario.expectDownload ? 1 : 0,
        scenario.name,
      );
      assert.equal(
        payload.summary.conflicts,
        scenario.expectConflict ? 1 : 0,
        scenario.name,
      );
      assert.equal(
        mock.methodCalls("downloadFile").length,
        scenario.expectDownload ? 1 : 0,
        scenario.name,
      );

      if (scenario.expectDownload) {
        assert.equal(readFileSync(localFilePath, "utf8"), "remote-buffer");
        assert.equal(
          mock.methodCalls("downloadFile")[0].args[0],
          "/me/drive/items/remote-1/content",
        );
      }

      if (scenario.expectConflict) {
        assert.match(
          payload.details.conflicts[0].newName,
          /^report_remote_\d+\.txt$/,
        );
      }
    } finally {
      cleanupTempDir(localPath);
    }
  }
});

test("sync_folder honors includePatterns and excludePatterns", async () => {
  const localPath = createTempDir("sync-patterns-");

  try {
    writeFileSync(path.join(localPath, "report.xlsx"), "report");
    writeFileSync(path.join(localPath, "notes.txt"), "notes");
    writeFileSync(path.join(localPath, "skip.tmp"), "skip");

    const mock = createMockGraphClient({
      get: async () => ({ success: true, data: { value: [] } }),
      uploadFile: async () => ({ success: true, data: { id: "uploaded" } }),
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleSyncFolder({
      localPath,
      remotePath: "Docs",
      direction: "upload",
      includePatterns: ["*.xlsx", "*.tmp"],
      excludePatterns: ["*.tmp"],
    })) as ToolEnvelope;
    const payload = parsePayload<any>(response);

    assert.equal(payload.summary.uploaded, 1);
    assert.equal(mock.methodCalls("uploadFile").length, 1);
    assert.equal(mock.methodCalls("uploadFile")[0].args[2], "report.xlsx");
    assert.deepEqual(
      payload.details.skipped.map((entry: any) => ({
        name: entry.name,
        reason: entry.reason,
      })),
      [
        { name: "notes.txt", reason: "Not in include pattern" },
        { name: "skip.tmp", reason: "In exclude pattern" },
      ],
    );
  } finally {
    cleanupTempDir(localPath);
  }
});

test("sync_folder returns error when upload localPath does not exist", async () => {
  const mock = createMockGraphClient();
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleSyncFolder({
    localPath: "/tmp/path-that-does-not-exist-for-upload",
    remotePath: "Docs",
    direction: "upload",
  })) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.match(response.content[0].text, /Local path does not exist/);
  assert.equal(mock.methodCalls("get").length, 0);
});

test("sync_folder creates local directory for download when missing", async () => {
  const rootDir = createTempDir("sync-download-create-root-");
  const localPath = path.join(rootDir, "missing-folder");

  try {
    const mock = createMockGraphClient({
      get: async () => ({
        success: true,
        data: {
          value: [
            {
              id: "remote-1",
              name: "report.txt",
              size: 6,
              file: { mimeType: "text/plain" },
              lastModifiedDateTime: "2026-04-14T18:00:00.000Z",
            },
          ],
        },
      }),
      downloadFile: async () => ({
        success: true,
        data: Buffer.from("remote"),
      }),
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleSyncFolder({
      localPath,
      remotePath: "Docs",
      direction: "download",
    })) as ToolEnvelope;
    const payload = parsePayload<any>(response);

    assert.equal(response.isError, undefined);
    assert.equal(existsSync(localPath), true);
    assert.equal(
      readFileSync(path.join(localPath, "report.txt"), "utf8"),
      "remote",
    );
    assert.equal(payload.summary.downloaded, 1);
  } finally {
    cleanupTempDir(rootDir);
  }
});

test("sync_folder deletes local and remote orphans when requested", async () => {
  const localPath = createTempDir("sync-delete-orphans-");

  try {
    writeFileSync(path.join(localPath, "remote-shared.txt"), "shared");
    writeFileSync(path.join(localPath, "local-only.txt"), "local-only");
    mkdirSync(path.join(localPath, "local-folder-only"));

    const mock = createMockGraphClient({
      get: async () => ({
        success: true,
        data: {
          value: [
            {
              id: "shared-1",
              name: "remote-shared.txt",
              size: 10,
              file: { mimeType: "text/plain" },
              lastModifiedDateTime: "2026-04-14T18:00:00.000Z",
            },
            {
              id: "remote-only-1",
              name: "remote-only.txt",
              size: 10,
              file: { mimeType: "text/plain" },
              lastModifiedDateTime: "2026-04-14T18:00:00.000Z",
            },
          ],
        },
      }),
      delete: async () => ({ success: true }),
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleSyncFolder({
      localPath,
      remotePath: "Docs",
      direction: "bidirectional",
      deleteOrphans: true,
      conflictResolution: "remote",
    })) as ToolEnvelope;
    const payload = parsePayload<any>(response);

    assert.equal(response.isError, undefined);
    assert.equal(existsSync(path.join(localPath, "local-only.txt")), false);
    assert.equal(existsSync(path.join(localPath, "local-folder-only")), false);
    assert.equal(mock.methodCalls("delete").length, 1);
    assert.equal(
      mock.methodCalls("delete")[0].args[0],
      "/me/drive/items/remote-only-1",
    );
    assert.deepEqual(
      payload.details.deleted
        .map((entry: any) => `${entry.location}:${entry.name}`)
        .sort(),
      [
        "local:local-folder-only",
        "local:local-only.txt",
        "remote:remote-only.txt",
      ],
    );
  } finally {
    cleanupTempDir(localPath);
  }
});

test("batch_file_operations covers successful operation branches sequentially", async () => {
  const localPath = createTempDir("batch-ops-sequential-");

  try {
    const uploadSource = path.join(localPath, "upload.txt");
    writeFileSync(uploadSource, "upload-me");

    const downloadTarget = path.join(localPath, "downloaded.txt");

    const mock = createMockGraphClient({
      uploadFile: async () => ({ success: true, data: { id: "uploaded-1" } }),
      downloadFile: async () => ({
        success: true,
        data: Buffer.from("downloaded-content"),
      }),
      patch: async () => ({ success: true, data: { id: "patched" } }),
      post: async () => ({ success: true, data: { job: "copy-1" } }),
      delete: async () => ({ success: true }),
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleBatchFileOperations({
      operations: [
        {
          operation: "upload",
          source: uploadSource,
          destination: "Docs/uploaded.txt",
        },
        {
          operation: "download",
          source: "Docs/server.txt",
          destination: downloadTarget,
          itemId: "download-id",
        },
        {
          operation: "move",
          itemId: "move-id",
          destination: "/drive/root:/Archive",
        },
        {
          operation: "copy",
          itemId: "copy-id",
          destination: "/drive/root:/Copies",
          newName: "copy-name.txt",
        },
        { operation: "delete", itemId: "delete-id" },
        { operation: "delete", source: "Docs/delete-by-source.txt" },
        { operation: "rename", itemId: "rename-id", newName: "renamed.txt" },
      ],
    })) as ToolEnvelope;
    const payload = parsePayload<any>(response);

    assert.equal(payload.success, true);
    assert.deepEqual(payload.summary, {
      total: 7,
      successful: 7,
      failed: 0,
      parallel: false,
      stopOnError: false,
    });
    assert.equal(readFileSync(downloadTarget, "utf8"), "downloaded-content");

    assert.equal(
      mock.methodCalls("uploadFile")[0].args[0],
      "/me/drive/root:/Docs/uploaded.txt:/content",
    );
    assert.equal(
      mock.methodCalls("downloadFile")[0].args[0],
      "/me/drive/items/download-id/content",
    );

    const patchCalls = mock.methodCalls("patch");
    assert.deepEqual(
      patchCalls.map((call) => [call.args[0], call.args[1]]),
      [
        [
          "/me/drive/items/move-id",
          { parentReference: { path: "/drive/root:/Archive" } },
        ],
        ["/me/drive/items/rename-id", { name: "renamed.txt" }],
      ],
    );

    assert.deepEqual(mock.methodCalls("post")[0].args, [
      "/me/drive/items/copy-id/copy",
      {
        parentReference: { path: "/drive/root:/Copies" },
        name: "copy-name.txt",
      },
    ]);
    assert.deepEqual(
      mock.methodCalls("delete").map((call) => call.args[0]),
      [
        "/me/drive/items/delete-id",
        "/me/drive/root:/Docs/delete-by-source.txt",
      ],
    );
    assert.deepEqual(
      payload.results.map((result: any) => result.status),
      [
        "completed",
        "completed",
        "completed",
        "completed",
        "completed",
        "completed",
        "completed",
      ],
    );
  } finally {
    cleanupTempDir(localPath);
  }
});

test("batch_file_operations handles unknown operation and continues when stopOnError is false", async () => {
  const localPath = createTempDir("batch-ops-continue-");

  try {
    const uploadSource = path.join(localPath, "upload.txt");
    writeFileSync(uploadSource, "upload-me");

    const mock = createMockGraphClient({
      uploadFile: async () => ({ success: true, data: { id: "uploaded-1" } }),
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleBatchFileOperations({
      operations: [
        {
          operation: "upload",
          source: uploadSource,
          destination: "Docs/uploaded.txt",
        },
        { operation: "explode" },
        {
          operation: "upload",
          source: uploadSource,
          destination: "Docs/uploaded-2.txt",
        },
      ],
      stopOnError: false,
    })) as ToolEnvelope;
    const payload = parsePayload<any>(response);

    assert.equal(payload.success, false);
    assert.equal(payload.summary.total, 3);
    assert.equal(payload.summary.successful, 2);
    assert.equal(payload.summary.failed, 1);
    assert.equal(mock.methodCalls("uploadFile").length, 2);
    assert.equal(payload.results[1].status, "failed");
    assert.match(payload.results[1].error, /Unknown operation: explode/);
  } finally {
    cleanupTempDir(localPath);
  }
});

test("batch_file_operations stops sequential execution after first failure when stopOnError is true", async () => {
  const localPath = createTempDir("batch-ops-stop-");

  try {
    const uploadSource = path.join(localPath, "upload.txt");
    writeFileSync(uploadSource, "upload-me");

    const mock = createMockGraphClient({
      uploadFile: async () => ({ success: true, data: { id: "uploaded-1" } }),
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleBatchFileOperations({
      operations: [
        {
          operation: "upload",
          source: uploadSource,
          destination: "Docs/uploaded.txt",
        },
        { operation: "download", source: "Docs/server.txt" },
        {
          operation: "upload",
          source: uploadSource,
          destination: "Docs/should-not-run.txt",
        },
      ],
      stopOnError: true,
      parallel: false,
    })) as ToolEnvelope;

    assert.equal(response.isError, true);
    assert.match(
      response.content[0].text,
      /Source and destination required for download/,
    );
    assert.equal(mock.methodCalls("uploadFile").length, 1);
    assert.equal(mock.methodCalls("downloadFile").length, 0);
  } finally {
    cleanupTempDir(localPath);
  }
});

test("batch_file_operations executes all operations in parallel and preserves per-item failures", async () => {
  const localPath = createTempDir("batch-ops-parallel-");

  try {
    const uploadSource = path.join(localPath, "upload.txt");
    writeFileSync(uploadSource, "upload-me");

    const mock = createMockGraphClient({
      uploadFile: async () => {
        await new Promise((resolve) => setTimeout(resolve, 25));
        return { success: true, data: { id: "uploaded-1" } };
      },
      downloadFile: async () => {
        await new Promise((resolve) => setTimeout(resolve, 5));
        throw new Error("network failure");
      },
      delete: async () => {
        await new Promise((resolve) => setTimeout(resolve, 10));
        return { success: true };
      },
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleBatchFileOperations({
      operations: [
        {
          operation: "upload",
          source: uploadSource,
          destination: "Docs/uploaded.txt",
        },
        {
          operation: "download",
          source: "Docs/server.txt",
          destination: path.join(localPath, "download.txt"),
          itemId: "download-id",
        },
        { operation: "delete", itemId: "delete-id" },
      ],
      parallel: true,
      stopOnError: true,
    })) as ToolEnvelope;
    const payload = parsePayload<any>(response);

    assert.equal(payload.success, false);
    assert.deepEqual(payload.summary, {
      total: 3,
      successful: 2,
      failed: 1,
      parallel: true,
      stopOnError: true,
    });
    assert.equal(mock.methodCalls("uploadFile").length, 1);
    assert.equal(mock.methodCalls("downloadFile").length, 1);
    assert.equal(mock.methodCalls("delete").length, 1);
    assert.equal(payload.results[1].status, "failed");
    assert.match(payload.results[1].error, /network failure/);
  } finally {
    cleanupTempDir(localPath);
  }
});

test("batch_file_operations validates batch size and empty operations", async () => {
  const mock = createMockGraphClient();
  __setGraphClientInstanceForTests(mock.client as any);

  const emptyResponse = (await handleBatchFileOperations({
    operations: [],
  })) as ToolEnvelope;
  assert.equal(emptyResponse.isError, true);
  assert.match(
    emptyResponse.content[0].text,
    /At least one operation is required/,
  );

  const oversizedResponse = (await handleBatchFileOperations({
    operations: Array.from({ length: 51 }, () => ({
      operation: "delete",
      itemId: "x",
    })),
  })) as ToolEnvelope;
  assert.equal(oversizedResponse.isError, true);
  assert.match(
    oversizedResponse.content[0].text,
    /Maximum 50 operations allowed per batch/,
  );
});
