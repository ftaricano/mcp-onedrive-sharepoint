import test from "node:test";
import assert from "node:assert/strict";
import { mkdirSync, writeFileSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { __setGraphClientInstanceForTests } from "../graph/client.js";
import { handleDownloadFile, handleUploadFile } from "../tools/files/index.js";
import {
  createMockGraphClient,
  parsePayload,
} from "./helpers/tool-test-helpers.js";
import { registerGraphClientTestLifecycle } from "./helpers/test-lifecycle.js";

registerGraphClientTestLifecycle();
const tags = { eTag: '"version-1"', cTag: '"content-1"' };
const item = { id: "item-1", name: "report.txt", ...tags };

test("upload exposes remote eTag and cTag and masks signed webUrl", async () => {
  const localPath = join(tmpdir(), `upload-tags-${process.pid}.txt`);
  writeFileSync(localPath, "fixture");
  const mock = createMockGraphClient({
    uploadFile: async () => ({
      success: true,
      data: {
        ...item,
        webUrl:
          "https://contoso.sharepoint.com/report.txt?sig=fictional-canary",
      },
    }),
  });
  __setGraphClientInstanceForTests(mock.client as any);
  const response = await handleUploadFile({
    localPath,
    remotePath: "report.txt",
  });
  const payload = parsePayload(response);
  assert.equal(payload.success, true);
  assert.equal(payload.file.eTag, tags.eTag);
  assert.equal(payload.file.cTag, tags.cTag);
  assert.equal(
    payload.file.webUrl,
    "https://contoso.sharepoint.com/report.txt?***",
  );
  assert.ok(!JSON.stringify(response).includes("fictional-canary"));
});
for (const outputPath of [false, true]) {
  test(`download exposes tags and binds content request (outputPath=${outputPath})`, async () => {
    const mock = createMockGraphClient({
      get: async () => ({ success: true, data: item }),
      downloadFile: async () => ({
        success: true,
        data: Buffer.from("fixture"),
      }),
    });
    __setGraphClientInstanceForTests(mock.client as any);
    const target = join(tmpdir(), `download-tags-${process.pid}`, "report.txt");
    mkdirSync(join(tmpdir(), `download-tags-${process.pid}`), {
      recursive: true,
    });
    const payload = parsePayload(
      await handleDownloadFile({
        fileId: item.id,
        ...(outputPath ? { outputPath: target } : {}),
      }),
    );
    assert.equal(payload.success, true);
    assert.deepEqual(payload.file, item);
    assert.deepEqual(mock.methodCalls("get")[0].args, [
      "/me/drive/items/item-1",
      { $select: "id,name,eTag,cTag" },
    ]);
    assert.deepEqual(mock.methodCalls("downloadFile")[0].args, [
      "/me/drive/items/item-1/content",
      { headers: { "If-Match": tags.eTag } },
    ]);
    if (outputPath) assert.equal(readFileSync(target, "utf8"), "fixture");
  });
}
test("download reports missing tags explicitly without inventing versions", async () => {
  const mock = createMockGraphClient({
    get: async () => ({ success: true, data: { id: item.id } }),
    downloadFile: async () => ({ success: true, data: Buffer.from("fixture") }),
  });
  __setGraphClientInstanceForTests(mock.client as any);
  const payload = parsePayload(await handleDownloadFile({ fileId: item.id }));
  assert.equal(payload.file.eTag, null);
  assert.equal(payload.file.cTag, null);
  assert.deepEqual(mock.methodCalls("downloadFile")[0].args[1], {
    headers: undefined,
  });
});
test("metadata failure prevents download; changed eTag never produces success", async () => {
  const mock = createMockGraphClient({ get: async () => ({ success: false }) });
  __setGraphClientInstanceForTests(mock.client as any);
  assert.equal(
    "isError" in (await handleDownloadFile({ fileId: item.id })),
    true,
  );
  assert.equal(mock.methodCalls("downloadFile").length, 0);
  const changed = createMockGraphClient({
    get: async () => ({ success: true, data: item }),
    downloadFile: async () => {
      throw new Error("Precondition failed (412)");
    },
  });
  __setGraphClientInstanceForTests(changed.client as any);
  assert.equal(
    "isError" in (await handleDownloadFile({ fileId: item.id })),
    true,
  );
});
