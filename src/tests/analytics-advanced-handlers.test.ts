import { test } from 'node:test';
import assert from 'node:assert/strict';

import { __setGraphClientInstanceForTests } from '../graph/client.js';
import { handleStorageAnalytics, handleVersionManagement } from '../tools/advanced/analytics.js';
import { registerGraphClientTestLifecycle } from './helpers/test-lifecycle.js';
import { createMockGraphClient, parsePayload, type ToolEnvelope } from './helpers/tool-test-helpers.js';

registerGraphClientTestLifecycle();

test('storage_analytics covers summary, duplicates, large_files, file_types, detailed recursion, and version aggregation', async () => {
  const drive = {
    id: 'drive-1',
    name: 'OneDrive',
    driveType: 'business',
    quota: { total: 1000, used: 600, remaining: 400 },
  };

  const rootItems = {
    value: [
      {
        id: 'file-a',
        name: 'dup.docx',
        size: 200,
        file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
        lastModifiedDateTime: '2024-01-01T00:00:00.000Z',
        createdDateTime: '2023-12-01T00:00:00.000Z',
        webUrl: 'https://contoso/files/a',
        parentReference: { path: '/drive/root:' },
      },
      {
        id: 'file-b',
        name: 'dup.docx',
        size: 200,
        file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
        lastModifiedDateTime: '2024-02-01T00:00:00.000Z',
        createdDateTime: '2023-12-05T00:00:00.000Z',
        webUrl: 'https://contoso/files/b',
        parentReference: { path: '/drive/root:/Archive' },
      },
      {
        id: 'file-c',
        name: 'huge.pptx',
        size: 5 * 1024 * 1024,
        file: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
        lastModifiedDateTime: '2024-03-01T00:00:00.000Z',
        createdDateTime: '2024-02-15T00:00:00.000Z',
        webUrl: 'https://contoso/files/c',
        parentReference: { path: '/drive/root:' },
      },
      {
        id: 'folder-1',
        name: 'Nested',
        size: 0,
        folder: { childCount: 1 },
        lastModifiedDateTime: '2024-04-01T00:00:00.000Z',
        createdDateTime: '2024-04-01T00:00:00.000Z',
        webUrl: 'https://contoso/folder/1',
        parentReference: { path: '/drive/root:' },
      },
    ],
  };

  const nestedItems = {
    value: [
      {
        id: 'file-d',
        name: 'nested.xlsx',
        size: 1024,
        file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
        lastModifiedDateTime: '2024-04-02T00:00:00.000Z',
        createdDateTime: '2024-04-01T00:00:00.000Z',
        webUrl: 'https://contoso/files/d',
        parentReference: { path: '/drive/root:/Nested' },
      },
    ],
  };

  const versionMap: Record<string, any> = {
    '/me/drive/items/file-a/versions': { value: [{ id: 'current', size: 200 }, { id: '1.0', size: 180, lastModifiedDateTime: '2024-01-01T00:00:00.000Z' }] },
    '/me/drive/items/file-b/versions': { value: [{ id: 'current', size: 200 }] },
    '/me/drive/items/file-c/versions': { value: [{ id: 'current', size: 5 * 1024 * 1024 }, { id: '1.0', size: 1024, lastModifiedDateTime: '2024-02-01T00:00:00.000Z' }] },
    '/me/drive/items/file-d/versions': { value: [{ id: 'current', size: 1024 }, { id: '1.0', size: 512, lastModifiedDateTime: '2024-04-01T00:00:00.000Z' }] },
  };

  const runAnalysis = async (analysisType: string, extraArgs: Record<string, unknown> = {}) => {
    const mock = createMockGraphClient({
      get: async (endpoint: string) => {
        if (endpoint === '/me/drive') {
          return { success: true, data: drive };
        }
        if (endpoint === '/me/drive/root/children') {
          return { success: true, data: rootItems };
        }
        if (endpoint === '/me/drive/items/folder-1/children') {
          return { success: true, data: nestedItems };
        }
        if (versionMap[endpoint]) {
          return { success: true, data: versionMap[endpoint] };
        }
        throw new Error(`Unexpected GET ${endpoint}`);
      },
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleStorageAnalytics({
      analysisType,
      thresholds: { largeFileSize: 1, oldFileDays: 30 },
      includeVersions: true,
      ...extraArgs,
    })) as ToolEnvelope;

    return { payload: parsePayload<any>(response), mock };
  };

  let result = await runAnalysis('summary');
  assert.equal(result.payload.summary.totalFiles, 3);
  assert.equal(result.payload.summary.totalFolders, 1);
  assert.equal(result.payload.versions.filesWithVersions, 2);

  result = await runAnalysis('duplicates');
  assert.equal(result.payload.duplicates.totalDuplicateSets, 1);
  assert.equal(result.payload.duplicates.duplicates[0].duplicateCount, 2);

  result = await runAnalysis('large_files');
  assert.equal(result.payload.largeFiles.count, 1);
  assert.equal(result.payload.largeFiles.files[0].name, 'huge.pptx');

  result = await runAnalysis('file_types');
  assert.equal(result.payload.fileTypes.totalTypes, 2);
  assert.equal(result.payload.fileTypes.topBySize[0].extension, 'pptx');

  result = await runAnalysis('detailed');
  assert.equal(result.payload.detailed.fileCount, 4);
  assert.equal(result.payload.detailed.folderCount, 1);
  assert.equal(result.payload.detailed.duplicates.length, 1);
});

test('version_management covers list, restore, delete, cleanup, compare, and invalid action branches', async () => {
  const listMock = createMockGraphClient({
    get: async (endpoint: string) => {
      if (endpoint === '/sites/site-1/drive/root:/Docs/report.docx') {
        return { success: true, data: { id: 'item-1' } };
      }
      if (endpoint === '/sites/site-1/drive/items/item-1/versions') {
        return {
          success: true,
          data: {
            value: [
              { id: 'current', version: '3.0', size: 30, lastModifiedDateTime: '2026-04-14T00:00:00.000Z', lastModifiedBy: { user: { displayName: 'Alex' } } },
              { id: '2.0', version: '2.0', size: 20, lastModifiedDateTime: '2026-04-13T00:00:00.000Z', lastModifiedBy: { user: { displayName: 'Sam' } } },
            ],
          },
        };
      }
      if (endpoint === '/sites/site-1/drive/items/item-1') {
        return { success: true, data: { name: 'report.docx', size: 30, webUrl: 'https://contoso/report' } };
      }
      throw new Error(`Unexpected GET ${endpoint}`);
    },
  });
  __setGraphClientInstanceForTests(listMock.client as any);

  let response = (await handleVersionManagement({ itemPath: 'Docs/report.docx', siteId: 'site-1', action: 'list' })) as ToolEnvelope;
  let payload = parsePayload<any>(response);
  assert.equal(payload.versionCount, 2);
  assert.equal(payload.file.name, 'report.docx');

  const restoreMock = createMockGraphClient({
    post: async (endpoint: string) => {
      assert.equal(endpoint, '/me/drive/items/item-1/versions/2.0/restoreVersion');
      return { success: true };
    },
  });
  __setGraphClientInstanceForTests(restoreMock.client as any);

  response = (await handleVersionManagement({ itemId: 'item-1', action: 'restore', versionId: '2.0' })) as ToolEnvelope;
  payload = parsePayload<any>(response);
  assert.equal(payload.restoredVersion, '2.0');

  const deleteMock = createMockGraphClient({
    delete: async (endpoint: string) => {
      assert.equal(endpoint, '/me/drive/items/item-1/versions/1.0');
      return { success: true };
    },
  });
  __setGraphClientInstanceForTests(deleteMock.client as any);

  response = (await handleVersionManagement({ itemId: 'item-1', action: 'delete', versionId: '1.0' })) as ToolEnvelope;
  payload = parsePayload<any>(response);
  assert.equal(payload.deletedVersion, '1.0');

  const cleanupMock = createMockGraphClient({
    get: async (endpoint: string) => {
      assert.equal(endpoint, '/me/drive/items/item-1/versions');
      return {
        success: true,
        data: {
          value: [
            { id: 'current', version: '4.0', size: 40, lastModifiedDateTime: '2026-04-14T00:00:00.000Z' },
            { id: '3.0', version: '3.0', size: 30, lastModifiedDateTime: '2026-04-13T00:00:00.000Z' },
            { id: '2.0', version: '2.0', size: 20, lastModifiedDateTime: '2026-04-12T00:00:00.000Z' },
            { id: '1.0', version: '1.0', size: 10, lastModifiedDateTime: '2026-04-11T00:00:00.000Z' },
          ],
        },
      };
    },
    delete: async (endpoint: string) => {
      if (endpoint.endsWith('/2.0')) {
        throw new Error('cannot delete protected version');
      }
      return { success: true };
    },
  });
  __setGraphClientInstanceForTests(cleanupMock.client as any);

  response = (await handleVersionManagement({ itemId: 'item-1', action: 'cleanup', keepVersions: 2 })) as ToolEnvelope;
  payload = parsePayload<any>(response);
  assert.equal(payload.deletedCount, 1);
  assert.equal(payload.failedDeletions.length, 1);
  assert.equal(payload.failedDeletions[0].id, '2.0');

  const compareMock = createMockGraphClient({
    get: async (endpoint: string) => {
      if (endpoint.endsWith('/versions/1.0')) {
        return {
          success: true,
          data: { id: '1.0', version: '1.0', size: 10, lastModifiedDateTime: '2026-04-10T00:00:00.000Z', lastModifiedBy: { user: { displayName: 'Alex' } } },
        };
      }
      if (endpoint.endsWith('/versions/3.0')) {
        return {
          success: true,
          data: { id: '3.0', version: '3.0', size: 25, lastModifiedDateTime: '2026-04-13T00:00:00.000Z', lastModifiedBy: { user: { displayName: 'Sam' } } },
        };
      }
      throw new Error(`Unexpected GET ${endpoint}`);
    },
  });
  __setGraphClientInstanceForTests(compareMock.client as any);

  response = (await handleVersionManagement({ itemId: 'item-1', action: 'compare', versionId: '1.0', compareVersionId: '3.0' })) as ToolEnvelope;
  payload = parsePayload<any>(response);
  assert.equal(payload.comparison.differences.sizeDifference, 15);
  assert.equal(payload.comparison.differences.daysBetween, 3);

  const invalidMock = createMockGraphClient();
  __setGraphClientInstanceForTests(invalidMock.client as any);

  response = (await handleVersionManagement({ itemId: 'item-1', action: 'bogus' })) as ToolEnvelope;
  assert.equal(response.isError, true);
  assert.match(response.content[0].text, /Invalid action: bogus/);
});
