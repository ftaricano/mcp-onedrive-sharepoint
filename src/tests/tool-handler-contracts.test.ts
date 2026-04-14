import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';

import { __setGraphClientInstanceForTests } from '../graph/client.js';
import { __setUtilityDependenciesForTests, handleHealthCheck } from '../tools/utils/index.js';
import { handleListFiles, handleUploadFile } from '../tools/files/index.js';
import { handleDiscoverSites, handleListItems } from '../tools/sharepoint/index.js';
import { handleBatchFileOperations, handleSyncFolder } from '../tools/advanced/sync.js';
import { handleExcelAnalysis, handleExcelOperations } from '../tools/advanced/excel.js';
import { handleAdvancedShare, handleManagePermissions } from '../tools/advanced/collaboration.js';
import { handleStorageAnalytics, handleVersionManagement } from '../tools/advanced/analytics.js';
import { cleanupAllCaches } from '../utils/cache-manager.js';

type ToolEnvelope = {
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
};

function parsePayload(response: ToolEnvelope) {
  return JSON.parse(response.content[0].text);
}

test.afterEach(() => {
  __setGraphClientInstanceForTests(null);
  __setUtilityDependenciesForTests();
  cleanupAllCaches();
});

test.after(() => {
  cleanupAllCaches();
});

test('health_check handler returns centralized MCP contract on success', async () => {
  __setUtilityDependenciesForTests({
    getAuthInstance: async () => ({
      isAuthenticated: async () => true,
    }),
  });

  __setGraphClientInstanceForTests({
    healthCheck: async () => ({
      success: true,
      data: {
        user: {
          id: 'user-1',
          displayName: 'Jarvis',
          mail: 'jarvis@example.com',
          userPrincipalName: 'jarvis@example.com',
        },
      },
    }),
    get: async () => ({
      success: true,
      data: {
        id: 'drive-1',
        name: 'OneDrive',
        driveType: 'business',
        quota: {
          total: 100,
          used: 25,
          remaining: 75,
          state: 'normal',
        },
      },
    }),
    cleanup: () => {},
  } as any);

  const response = await handleHealthCheck({ includeUserInfo: true, includeDriveInfo: true }) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(response.content[0].type, 'text');
  assert.equal(payload.status, 'healthy');
  assert.equal(payload.authentication.isAuthenticated, true);
  assert.equal(payload.defaultDrive.id, 'drive-1');
});

test('health_check handler returns centralized MCP contract on error', async () => {
  __setUtilityDependenciesForTests({
    getAuthInstance: async () => ({
      isAuthenticated: async () => {
        throw new Error('auth unavailable');
      },
    }),
  });

  __setGraphClientInstanceForTests({ cleanup: () => {} } as any);

  const response = await handleHealthCheck({}) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(response.content[0].type, 'text');
  assert.match(response.content[0].text, /Error in health_check/);
});

test('list_files handler returns centralized MCP contract on success', async () => {
  __setGraphClientInstanceForTests({
    get: async () => ({
      success: true,
      data: {
        value: [
          {
            id: 'item-1',
            name: 'Report.xlsx',
            size: 42,
            file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
            lastModifiedDateTime: '2026-04-14T18:00:00.000Z',
            webUrl: 'https://contoso.sharepoint.com/report',
            parentReference: { path: '/drive/root:/Docs', driveId: 'drive-1' },
          },
        ],
      },
    }),
    cleanup: () => {},
  } as any);

  const response = await handleListFiles({ path: '/Docs', driveId: 'drive-1', limit: 10 }) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.target, 'drive:drive-1');
  assert.equal(payload.itemCount, 1);
  assert.equal(payload.items[0].name, 'Report.xlsx');
});

test('upload_file handler preserves JSON error envelope via centralized helper', async () => {
  __setGraphClientInstanceForTests({
    get: async () => {
      throw new Error('missing folder');
    },
    post: async () => ({ success: false }),
    cleanup: () => {},
  } as any);

  const response = await handleUploadFile({
    localPath: '/tmp/non-existent-upload.txt',
    remotePath: 'unsafe/report.txt',
  }) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, true);
  assert.equal(response.content[0].type, 'text');
  assert.equal(payload.success, false);
  assert.equal(payload.error, 'Path preparation failed');
});

test('discover_sites handler returns centralized MCP contract on success', async () => {
  __setGraphClientInstanceForTests({
    get: async (endpoint: string) => {
      if (endpoint === '/sites/root') {
        return {
          success: true,
          data: {
            id: 'root-site',
            displayName: 'Root Site',
            name: 'Root Site',
            webUrl: 'https://contoso.sharepoint.com',
            root: {},
          },
        };
      }

      return {
        success: true,
        data: {
          value: [
            {
              id: 'site-1',
              displayName: 'Finance',
              name: 'Finance',
              webUrl: 'https://contoso.sharepoint.com/sites/finance',
            },
          ],
        },
      };
    },
    cleanup: () => {},
  } as any);

  const response = await handleDiscoverSites({ search: 'fin', includePersonalSite: true, limit: 5 }) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.siteCount, 2);
  assert.equal(payload.sites[0].id, 'root-site');
  assert.equal(payload.sites[1].id, 'site-1');
});

test('list_items handler returns centralized MCP contract on error', async () => {
  __setGraphClientInstanceForTests({
    get: async () => {
      throw new Error('graph down');
    },
    cleanup: () => {},
  } as any);

  const response = await handleListItems({ siteId: 'site-1', listId: 'list-1' }) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(response.content[0].type, 'text');
  assert.match(response.content[0].text, /Error in list_items/);
});

test('sync_folder handler returns centralized MCP contract on success', async () => {
  const localPath = mkdtempSync(path.join(tmpdir(), 'sync-folder-contract-'));

  try {
    __setGraphClientInstanceForTests({
      get: async () => ({ success: true, data: { value: [] } }),
      cleanup: () => {},
    } as any);

    const response = await handleSyncFolder({
      localPath,
      remotePath: 'Docs',
      direction: 'bidirectional',
    }) as ToolEnvelope;
    const payload = parsePayload(response);

    assert.equal(response.isError, undefined);
    assert.equal(response.content[0].type, 'text');
    assert.equal(payload.success, true);
    assert.equal(payload.summary.uploaded, 0);
    assert.equal(payload.summary.downloaded, 0);
  } finally {
    rmSync(localPath, { recursive: true, force: true });
  }
});

test('batch_file_operations handler returns centralized MCP contract on error', async () => {
  __setGraphClientInstanceForTests({ cleanup: () => {} } as any);

  const response = await handleBatchFileOperations({ operations: [] }) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(response.content[0].type, 'text');
  assert.match(response.content[0].text, /Error in batch_file_operations/);
});

test('excel_operations handler returns centralized MCP contract on success', async () => {
  __setGraphClientInstanceForTests({
    get: async () => ({
      success: true,
      data: {
        value: [{ id: 'ws-1', name: 'Sheet1', position: 0, visibility: 'Visible' }],
      },
    }),
    cleanup: () => {},
  } as any);

  const response = await handleExcelOperations({
    itemId: 'file-1',
    operation: 'list_worksheets',
  }) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(response.content[0].type, 'text');
  assert.equal(payload.operation, 'list_worksheets');
  assert.equal(payload.count, 1);
  assert.equal(payload.worksheets[0].id, 'ws-1');
});

test('excel_analysis handler returns centralized MCP contract on error', async () => {
  __setGraphClientInstanceForTests({ cleanup: () => {} } as any);

  const response = await handleExcelAnalysis({ analysisType: 'used_range' }) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(response.content[0].type, 'text');
  assert.match(response.content[0].text, /Error in excel_analysis/);
});

test('advanced_share handler returns centralized MCP contract on success', async () => {
  __setGraphClientInstanceForTests({
    post: async () => ({
      success: true,
      data: {
        value: [{
          id: 'perm-1',
          grantedTo: { user: { email: 'teammate@example.com' } },
          roles: ['write'],
          hasPassword: false,
          expirationDateTime: '2026-04-30T00:00:00.000Z',
          shareId: 'share-1',
        }],
      },
    }),
    cleanup: () => {},
  } as any);

  const response = await handleAdvancedShare({
    itemId: 'item-1',
    recipients: ['teammate@example.com'],
    permission: 'write',
  }) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(response.content[0].type, 'text');
  assert.equal(payload.success, true);
  assert.equal(payload.recipientCount, 1);
  assert.equal(payload.permissions[0].id, 'perm-1');
});

test('manage_permissions handler returns centralized MCP contract on error', async () => {
  __setGraphClientInstanceForTests({ cleanup: () => {} } as any);

  const response = await handleManagePermissions({ action: 'list' }) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(response.content[0].type, 'text');
  assert.match(response.content[0].text, /Error in manage_permissions/);
});

test('storage_analytics handler returns centralized MCP contract on success', async () => {
  __setGraphClientInstanceForTests({
    get: async (endpoint: string) => {
      if (endpoint === '/me/drive') {
        return {
          success: true,
          data: {
            id: 'drive-1',
            name: 'OneDrive',
            driveType: 'business',
            quota: { total: 1000, used: 250, remaining: 750 },
          },
        };
      }

      return {
        success: true,
        data: {
          value: [{
            id: 'item-1',
            name: 'Report.xlsx',
            size: 1024,
            file: { mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
            lastModifiedDateTime: '2026-04-01T00:00:00.000Z',
            createdDateTime: '2026-03-01T00:00:00.000Z',
            webUrl: 'https://contoso.sharepoint.com/report',
            parentReference: { path: '/drive/root:/Docs' },
          }],
        },
      };
    },
    cleanup: () => {},
  } as any);

  const response = await handleStorageAnalytics({ analysisType: 'summary' }) as ToolEnvelope;
  const payload = parsePayload(response);

  assert.equal(response.isError, undefined);
  assert.equal(response.content[0].type, 'text');
  assert.equal(payload.analysisType, 'summary');
  assert.equal(payload.summary.totalFiles, 1);
  assert.equal(payload.drive.id, 'drive-1');
});

test('version_management handler returns centralized MCP contract on error', async () => {
  __setGraphClientInstanceForTests({ cleanup: () => {} } as any);

  const response = await handleVersionManagement({ action: 'list' }) as ToolEnvelope;

  assert.equal(response.isError, true);
  assert.equal(response.content[0].type, 'text');
  assert.match(response.content[0].text, /Error in version_management/);
});
