import test from 'node:test';
import assert from 'node:assert/strict';

import { buildUrl } from '../config/endpoints.js';
import { extractPaginatedResult, jsonTextResponse, toolErrorResponse } from '../graph/contracts.js';
import {
  buildDriveChildrenEndpoint,
  buildDriveItemEndpoint,
  buildDriveSearchEndpoint,
  describeDriveTarget,
} from '../graph/resource-resolver.js';
import { GraphApiError, createUserFriendlyError } from '../graph/error-handler.js';

test('buildUrl appends query params without duplicating path placeholders', () => {
  const url = buildUrl('/sites/{siteId}/lists', {
    siteId: 'site-123',
    '$top': '25',
    '$orderby': 'displayName',
    '$filter': "name ne 'Archive'",
  }, false);

  assert.equal(
    url,
    "/sites/site-123/lists?%24top=25&%24orderby=displayName&%24filter=name+ne+%27Archive%27",
  );
});

test('extractPaginatedResult returns items and next page metadata', () => {
  const result = extractPaginatedResult(
    {
      value: [{ id: '1' }, { id: '2' }],
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/drive/root/children?$skiptoken=abc',
      '@odata.count': 12,
    },
    2,
  );

  assert.equal(result.items.length, 2);
  assert.equal(result.pagination.returned, 2);
  assert.equal(result.pagination.totalCount, 12);
  assert.equal(result.pagination.hasMore, true);
  assert.match(result.pagination.nextPageToken ?? '', /skiptoken=abc/);
});

test('resource resolver supports me, site and drive scopes', () => {
  assert.equal(buildDriveChildrenEndpoint({ path: '/Documents/Plans' }), '/me/drive/root:/Documents/Plans:/children');
  assert.equal(buildDriveChildrenEndpoint({ siteId: 'site-123', path: 'Docs' }), '/sites/site-123/drive/root:/Docs:/children');
  assert.equal(buildDriveChildrenEndpoint({ driveId: 'drive-123' }), '/drives/drive-123/root/children');
  assert.equal(buildDriveItemEndpoint({ driveId: 'drive-123', itemId: 'item-456' }, '/content'), '/drives/drive-123/items/item-456/content');
  assert.equal(buildDriveItemEndpoint({ siteId: 'site-123', itemPath: '/Docs/report.docx' }), '/sites/site-123/drive/root:/Docs/report.docx:');
  assert.equal(buildDriveSearchEndpoint({ driveId: 'drive-123' }, 'budget 2026'), "/drives/drive-123/root/search(q='budget%202026')");
  assert.equal(describeDriveTarget({}), 'me');
  assert.equal(describeDriveTarget({ siteId: 'site-123' }), 'site:site-123');
  assert.equal(describeDriveTarget({ driveId: 'drive-123' }), 'drive:drive-123');
});

test('JSON and error envelopes are MCP-compatible text responses', () => {
  const success = jsonTextResponse({ ok: true, nested: { value: 1 } });
  assert.equal(success.content.length, 1);
  assert.equal(success.content[0].type, 'text');
  assert.match(success.content[0].text, /"ok": true/);

  const error = toolErrorResponse('list_files', new GraphApiError({ error: { code: 'InvalidAuthenticationToken', message: 'expired' } }));
  assert.equal(error.isError, true);
  assert.match(error.content[0].text, /Error in list_files/);
  assert.match(error.content[0].text, /Authentication Error/);
});

test('GraphApiError produces actionable messages', () => {
  const error = new GraphApiError({ error: { code: 'ItemNotFound', message: 'missing file' } }, 'GET /me/drive/items/x', 404);
  assert.equal(error.category, 'NotFound');
  assert.equal(error.isRetryable, false);
  assert.match(createUserFriendlyError(error), /Suggested Action:/);
});
