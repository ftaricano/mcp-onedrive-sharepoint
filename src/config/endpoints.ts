/**
 * Microsoft Graph API endpoint configurations
 */

export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
export const GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';

export const ENDPOINTS = {
  // Authentication
  AUTH: {
    DEVICE_CODE: 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode',
    TOKEN: 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token',
  },
  
  // User and profile
  USER: {
    ME: '/me',
    DRIVES: '/me/drives',
    DRIVE: '/me/drive',
  },
  
  // Files and drives
  FILES: {
    DRIVE_ROOT: '/me/drive/root',
    DRIVE_ITEMS: '/me/drive/items',
    DRIVE_ITEM: '/me/drive/items/{itemId}',
    DRIVE_ITEM_CONTENT: '/me/drive/items/{itemId}/content',
    DRIVE_ITEM_CHILDREN: '/me/drive/items/{itemId}/children',
    DRIVE_SEARCH: '/me/drive/root/search(q=\'{query}\')',
    DRIVE_UPLOAD_SESSION: '/me/drive/items/{itemId}/createUploadSession',
    DRIVE_PERMISSIONS: '/me/drive/items/{itemId}/permissions',
    
    // SharePoint drive access
    SITE_DRIVE_ROOT: '/sites/{siteId}/drive/root',
    SITE_DRIVE_ITEMS: '/sites/{siteId}/drive/items',
    SITE_DRIVE_ITEM: '/sites/{siteId}/drive/items/{itemId}',
  },
  
  // SharePoint sites and lists
  SHAREPOINT: {
    SITES: '/sites',
    SITE: '/sites/{siteId}',
    SITE_LISTS: '/sites/{siteId}/lists',
    SITE_LIST: '/sites/{siteId}/lists/{listId}',
    LIST_ITEMS: '/sites/{siteId}/lists/{listId}/items',
    LIST_ITEM: '/sites/{siteId}/lists/{listId}/items/{itemId}',
    LIST_COLUMNS: '/sites/{siteId}/lists/{listId}/columns',
    SITE_SEARCH: '/sites/{siteId}/search(q=\'{query}\')',
    ROOT_SITE: '/sites/root',
  },
  
  // Excel workbooks
  EXCEL: {
    WORKBOOK: '/me/drive/items/{itemId}/workbook',
    WORKSHEETS: '/me/drive/items/{itemId}/workbook/worksheets',
    WORKSHEET: '/me/drive/items/{itemId}/workbook/worksheets/{worksheetId}',
    RANGE: '/me/drive/items/{itemId}/workbook/worksheets/{worksheetId}/range(address=\'{address}\')',
    USED_RANGE: '/me/drive/items/{itemId}/workbook/worksheets/{worksheetId}/usedRange',
    TABLES: '/me/drive/items/{itemId}/workbook/tables',
    TABLE: '/me/drive/items/{itemId}/workbook/tables/{tableId}',
    TABLE_ROWS: '/me/drive/items/{itemId}/workbook/tables/{tableId}/rows',
    TABLE_COLUMNS: '/me/drive/items/{itemId}/workbook/tables/{tableId}/columns',
    CHARTS: '/me/drive/items/{itemId}/workbook/worksheets/{worksheetId}/charts',
    NAMED_ITEMS: '/me/drive/items/{itemId}/workbook/names',
    
    // Session management for better performance
    CREATE_SESSION: '/me/drive/items/{itemId}/workbook/createSession',
    CLOSE_SESSION: '/me/drive/items/{itemId}/workbook/closeSession',
    
    // SharePoint Excel files
    SITE_WORKBOOK: '/sites/{siteId}/drive/items/{itemId}/workbook',
    SITE_WORKSHEETS: '/sites/{siteId}/drive/items/{itemId}/workbook/worksheets',
  },
  
  // Search
  SEARCH: {
    GLOBAL: '/search/query',
    DRIVE_SEARCH: '/me/drive/root/search(q=\'{query}\')',
    SITE_SEARCH: '/sites/{siteId}/drive/root/search(q=\'{query}\')',
  }
} as const;

// Helper function to build URLs with path parameters and OData query parameters.
export function buildUrl(endpoint: string, params: Record<string, string> = {}, useBase = true): string {
  let url = endpoint;
  const queryParams = new URLSearchParams();

  for (const [key, rawValue] of Object.entries(params)) {
    const value = String(rawValue);
    const placeholder = `{${key}}`;

    if (url.includes(placeholder)) {
      url = url.split(placeholder).join(encodeURIComponent(value));
      continue;
    }

    if (value.length === 0 || value === 'undefined' || value === 'null') {
      continue;
    }

    queryParams.append(key, value);
  }

  if (queryParams.size > 0) {
    url += `${url.includes('?') ? '&' : '?'}${queryParams.toString()}`;
  }

  // Add base URL if needed
  if (useBase && !url.startsWith('http')) {
    url = GRAPH_BASE_URL + url;
  }
  
  return url;
}

// Helper to build SharePoint site URL from domain/path
export function buildSiteUrl(siteDomain: string, sitePath = ''): string {
  if (sitePath) {
    return `/sites/${siteDomain}:/${sitePath}`;
  }
  return `/sites/${siteDomain}`;
}