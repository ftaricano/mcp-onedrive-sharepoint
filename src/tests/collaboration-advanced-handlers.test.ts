import { test } from 'node:test';
import assert from 'node:assert/strict';

import { __setGraphClientInstanceForTests } from '../graph/client.js';
import {
  handleAdvancedShare,
  handleCheckUserAccess,
  handleManagePermissions,
} from '../tools/advanced/collaboration.js';
import { registerGraphClientTestLifecycle } from './helpers/test-lifecycle.js';
import { createMockGraphClient, parsePayload, type ToolEnvelope } from './helpers/tool-test-helpers.js';

registerGraphClientTestLifecycle();

test('advanced_share covers itemPath payload composition including owner role and optional fields', async () => {
  const mock = createMockGraphClient({
    post: async (endpoint: string, data?: any) => {
      assert.equal(endpoint, '/sites/site-1/drive/root:/Docs/report.docx:/invite');
      assert.deepEqual(data, {
        recipients: [{ email: 'owner@example.com' }],
        requireSignIn: false,
        sendInvitation: false,
        retainInheritedPermissions: false,
        roles: ['owner'],
        message: 'Please review',
        expirationDateTime: '2026-05-01T00:00:00.000Z',
      });

      return {
        success: true,
        data: {
          value: [
            {
              id: 'perm-1',
              grantedTo: { user: { email: 'owner@example.com' } },
              roles: ['owner'],
              hasPassword: false,
              expirationDateTime: '2026-05-01T00:00:00.000Z',
              shareId: 'share-1',
            },
          ],
        },
      };
    },
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleAdvancedShare({
    itemPath: 'Docs/report.docx',
    siteId: 'site-1',
    recipients: ['owner@example.com'],
    permission: 'owner',
    requireSignIn: false,
    sendInvitation: false,
    message: 'Please review',
    expirationDateTime: '2026-05-01T00:00:00.000Z',
    retainInheritedPermissions: false,
  })) as ToolEnvelope;
  const payload = parsePayload<any>(response);

  assert.equal(payload.success, true);
  assert.equal(payload.permissions[0].grantedTo, 'owner@example.com');
  assert.deepEqual(payload.permissions[0].roles, ['owner']);
});

test('manage_permissions covers list, update, revoke, and invalid action branches', async () => {
  const listMock = createMockGraphClient({
    get: async (endpoint: string, params?: Record<string, unknown>) => {
      assert.equal(endpoint, '/me/drive/items/item-1/permissions');
      assert.deepEqual(params, { '$expand': 'grantedTo,grantedToIdentities' });
      return {
        success: true,
        data: {
          value: [
            {
              id: 'perm-1',
              roles: ['write'],
              grantedTo: { user: { email: 'user@example.com', displayName: 'User', id: 'user-1' } },
              link: { type: 'edit', scope: 'users', webUrl: 'https://contoso/link' },
              inheritedFrom: null,
              expirationDateTime: null,
              hasPassword: false,
            },
          ],
        },
      };
    },
  });
  __setGraphClientInstanceForTests(listMock.client as any);

  let response = (await handleManagePermissions({ itemId: 'item-1', action: 'list' })) as ToolEnvelope;
  let payload = parsePayload<any>(response);
  assert.equal(payload.permissionCount, 1);
  assert.equal(payload.permissions[0].grantedTo.user.email, 'user@example.com');

  const updateMock = createMockGraphClient({
    patch: async (endpoint: string, data?: any) => {
      assert.equal(endpoint, '/sites/site-1/drive/root:/Docs/report.docx/permissions/perm-1');
      assert.deepEqual(data, { roles: ['read'] });
      return { success: true, data: { id: 'perm-1', roles: ['read'] } };
    },
  });
  __setGraphClientInstanceForTests(updateMock.client as any);

  response = (await handleManagePermissions({
    itemPath: 'Docs/report.docx',
    siteId: 'site-1',
    action: 'update',
    permissionId: 'perm-1',
    newRoles: ['read'],
  })) as ToolEnvelope;
  payload = parsePayload<any>(response);
  assert.equal(payload.success, true);
  assert.deepEqual(payload.newRoles, ['read']);

  const revokeMock = createMockGraphClient({
    delete: async (endpoint: string) => {
      assert.equal(endpoint, '/me/drive/items/item-1/permissions/perm-9');
      return { success: true };
    },
  });
  __setGraphClientInstanceForTests(revokeMock.client as any);

  response = (await handleManagePermissions({ itemId: 'item-1', action: 'revoke', permissionId: 'perm-9' })) as ToolEnvelope;
  payload = parsePayload<any>(response);
  assert.equal(payload.success, true);
  assert.equal(payload.permissionId, 'perm-9');

  const invalidMock = createMockGraphClient();
  __setGraphClientInstanceForTests(invalidMock.client as any);

  response = (await handleManagePermissions({ itemId: 'item-1', action: 'bogus' })) as ToolEnvelope;
  assert.equal(response.isError, true);
  assert.match(response.content[0].text, /Invalid action: bogus/);
});

test('check_user_access merges direct, shared-link, and inherited roles', async () => {
  const mock = createMockGraphClient({
    get: async (endpoint: string, params?: Record<string, unknown>) => {
      assert.equal(endpoint, '/me/drive/items/item-1/permissions');
      assert.deepEqual(params, { '$expand': 'grantedTo,grantedToIdentities' });
      return {
        success: true,
        data: {
          value: [
            {
              id: 'perm-direct',
              roles: ['read'],
              grantedTo: { user: { email: 'member@example.com' } },
              expirationDateTime: null,
            },
            {
              id: 'perm-link',
              roles: ['write'],
              grantedToIdentities: [{ user: { email: 'member@example.com' } }],
              link: { type: 'edit' },
              expirationDateTime: null,
            },
            {
              id: 'perm-inherited',
              roles: ['owner'],
              inheritedFrom: { path: '/parents/folder' },
            },
          ],
        },
      };
    },
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleCheckUserAccess({
    itemId: 'item-1',
    userEmail: 'member@example.com',
    includeInherited: true,
  })) as ToolEnvelope;
  const payload = parsePayload<any>(response);

  assert.equal(payload.hasAccess, true);
  assert.deepEqual(payload.effectiveRoles.sort(), ['owner', 'read', 'write']);
  assert.equal(payload.directPermissions.length, 2);
  assert.equal(payload.inheritedPermissions.length, 1);
  assert.deepEqual(payload.summary, { canRead: true, canWrite: true, isOwner: true });
});
