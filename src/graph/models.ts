/**
 * TypeScript models for Microsoft Graph API responses
 * Comprehensive models for OneDrive, SharePoint, and Excel operations
 */

// Base Microsoft Graph models
export interface GraphError {
  code: string;
  message: string;
  details?: Array<{
    code: string;
    message: string;
    target?: string;
  }>;
  innerError?: {
    'request-id': string;
    date: string;
  };
}

export interface GraphResponse<T> {
  '@odata.context'?: string;
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
  value?: T[];
  error?: GraphError;
}

// User and authentication models
export interface User {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
}

// Drive and file models
export interface Drive {
  id: string;
  name: string;
  description?: string;
  driveType: 'personal' | 'business' | 'documentLibrary' | 'sharepoint';
  owner: {
    user?: User;
    group?: {
      id: string;
      displayName: string;
    };
  };
  quota: {
    total: number;
    used: number;
    remaining: number;
    state: 'normal' | 'nearing' | 'critical' | 'exceeded';
  };
  webUrl: string;
}

export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl: string;
  downloadUrl?: string;
  '@microsoft.graph.downloadUrl'?: string;
  file?: {
    mimeType: string;
    hashes?: {
      quickXorHash?: string;
      sha1Hash?: string;
      sha256Hash?: string;
    };
  };
  folder?: {
    childCount: number;
  };
  image?: {
    height: number;
    width: number;
  };
  parentReference: {
    driveId: string;
    driveType: string;
    id: string;
    path: string;
  };
  createdBy: {
    user: User;
  };
  lastModifiedBy: {
    user: User;
  };
  shared?: {
    scope: 'anonymous' | 'organization' | 'users';
  };
}

export interface Permission {
  id: string;
  roles: string[];
  shareId?: string;
  hasPassword?: boolean;
  grantedToV2?: {
    user?: User;
    group?: {
      id: string;
      displayName: string;
    };
    siteUser?: {
      id: string;
      displayName: string;
      loginName: string;
    };
  };
  link?: {
    type: 'view' | 'edit' | 'embed';
    scope: 'anonymous' | 'organization' | 'users';
    webUrl: string;
    application?: {
      id: string;
      displayName: string;
    };
  };
  expirationDateTime?: string;
}

// Upload session for large files
export interface UploadSession {
  uploadUrl: string;
  expirationDateTime: string;
  nextExpectedRanges: string[];
}

// SharePoint models
export interface Site {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  webUrl: string;
  siteCollection?: {
    hostname: string;
  };
  root?: {};
  sharepointIds?: {
    siteId: string;
    siteUrl: string;
    tenantId: string;
    webId: string;
  };
  createdDateTime: string;
  lastModifiedDateTime: string;
}

export interface List {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  webUrl: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  list: {
    contentTypesEnabled: boolean;
    hidden: boolean;
    template: string;
  };
  columns: ListColumn[];
  contentTypes?: ContentType[];
}

export interface ListColumn {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  required: boolean;
  hidden: boolean;
  indexed: boolean;
  type: string;
  text?: {
    allowMultipleLines: boolean;
    appendChangesToExistingText: boolean;
    linesForEditing: number;
    maxLength: number;
  };
  number?: {
    decimalPlaces: number;
    displayAs: string;
    maximum?: number;
    minimum?: number;
  };
  choice?: {
    allowTextEntry: boolean;
    choices: string[];
    displayAs: string;
  };
  lookup?: {
    allowMultipleValues: boolean;
    allowUnlimitedLength: boolean;
    columnName: string;
    listId: string;
    primaryLookupColumnId: string;
  };
}

export interface ContentType {
  id: string;
  name: string;
  description?: string;
  group: string;
  hidden: boolean;
  readOnly: boolean;
  sealed: boolean;
  documentSet?: {
    shouldPrefixNameToFile: boolean;
    allowedContentTypes: Array<{
      id: string;
      name: string;
    }>;
  };
}

export interface ListItem {
  id: string;
  webUrl: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  createdBy: {
    user: User;
  };
  lastModifiedBy: {
    user: User;
  };
  parentReference: {
    id: string;
    siteId: string;
  };
  contentType: {
    id: string;
    name: string;
  };
  fields: Record<string, any>;
  sharepointIds: {
    listId: string;
    listItemId: string;
    listItemUniqueId: string;
  };
}

// Excel models
export interface Workbook {
  id: string;
  names: WorkbookNamedItem[];
  worksheets: Worksheet[];
  tables: WorkbookTable[];
  application: {
    calculationMode: string;
  };
}

export interface Worksheet {
  id: string;
  name: string;
  position: number;
  visibility: 'visible' | 'hidden' | 'veryHidden';
  charts: WorkbookChart[];
  tables: WorkbookTable[];
  protection?: {
    protected: boolean;
  };
}

export interface WorkbookRange {
  address: string;
  addressLocal: string;
  cellCount: number;
  columnCount: number;
  rowCount: number;
  columnIndex: number;
  rowIndex: number;
  values: any[][];
  formulas: any[][];
  formulasLocal: any[][];
  formulasR1C1: any[][];
  hidden: boolean;
  numberFormat: any[][];
  text: any[][];
  valueTypes: any[][];
  format?: {
    fill: {
      color: string;
    };
    font: {
      bold: boolean;
      color: string;
      italic: boolean;
      name: string;
      size: number;
      underline: string;
    };
    borders: any;
    protection: {
      formulaHidden: boolean;
      locked: boolean;
    };
  };
}

export interface WorkbookTable {
  id: string;
  name: string;
  showHeaders: boolean;
  showTotals: boolean;
  style: string;
  highlightFirstColumn: boolean;
  highlightLastColumn: boolean;
  showBandedColumns: boolean;
  showBandedRows: boolean;
  showFilterButton: boolean;
  legacyId: string;
  columns: WorkbookTableColumn[];
  rows: WorkbookTableRow[];
  sort?: {
    fields: Array<{
      key: number;
      sortOn: string;
      ascending: boolean;
      color?: string;
      dataOption: string;
    }>;
    matchCase: boolean;
    hasHeaders: boolean;
  };
  worksheet: {
    id: string;
    name: string;
  };
}

export interface WorkbookTableColumn {
  id: string;
  name: string;
  index: number;
  values: any[][];
  filter?: {
    criteria: {
      filterOn: string;
      values?: string[];
      color?: string;
      operator?: string;
      criterion1?: string;
      criterion2?: string;
    };
  };
}

export interface WorkbookTableRow {
  index: number;
  values: any[][];
}

export interface WorkbookChart {
  id: string;
  name: string;
  height: number;
  width: number;
  left: number;
  top: number;
  axes: {
    categoryAxis: any;
    seriesAxis: any;
    valueAxis: any;
  };
  dataLabels: {
    position: string;
    showValue: boolean;
    showSeriesName: boolean;
    showCategoryName: boolean;
    showLegendKey: boolean;
    showPercentage: boolean;
    showBubbleSize: boolean;
    separator: string;
  };
  format: {
    fill: {
      setSolidColor: (color: string) => void;
    };
  };
  legend: {
    visible: boolean;
    position: string;
  };
  series: any[];
  title: {
    text: string;
    visible: boolean;
    overlay: boolean;
  };
  worksheet: {
    id: string;
    name: string;
  };
}

export interface WorkbookNamedItem {
  name: string;
  comment?: string;
  scope: string;
  type: string;
  value: any;
  visible: boolean;
  worksheet?: {
    id: string;
    name: string;
  };
}

// Session management for Excel
export interface WorkbookSession {
  id: string;
  persistChanges: boolean;
}

// Search models
export interface SearchResult {
  '@odata.type': string;
  name: string;
  size?: number;
  lastModifiedDateTime: string;
  webUrl: string;
  id: string;
  parentReference?: {
    driveId: string;
    id: string;
    path: string;
  };
  file?: {
    mimeType: string;
  };
  folder?: {
    childCount: number;
  };
}

// Error handling models
export interface ApiError {
  code: string;
  message: string;
  details?: string;
  statusCode?: number;
  context?: string;
}

// Request/Response helpers
export interface BatchRequest {
  id: string;
  method: 'GET' | 'POST' | 'PUT' | 'PATCH' | 'DELETE';
  url: string;
  headers?: Record<string, string>;
  body?: any;
}

export interface BatchResponse {
  id: string;
  status: number;
  headers?: Record<string, string>;
  body?: any;
}

// Configuration and metadata
export interface DriveQuota {
  total: number;
  used: number;
  remaining: number;
  deleted: number;
  state: 'normal' | 'nearing' | 'critical' | 'exceeded';
}

export interface SiteInformation {
  id: string;
  displayName: string;
  name: string;
  description?: string;
  webUrl: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  isPersonalSite: boolean;
  root?: any;
}

// Utility types
export type FileExtension = '.xlsx' | '.docx' | '.pptx' | '.pdf' | '.txt' | '.jpg' | '.png' | string;
export type MimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' | 
                      'application/vnd.openxmlformats-officedocument.wordprocessingml.document' |
                      'application/vnd.openxmlformats-officedocument.presentationml.presentation' |
                      'application/pdf' | 'text/plain' | 'image/jpeg' | 'image/png' | string;

export type ShareScope = 'anonymous' | 'organization' | 'users';
export type ShareRole = 'read' | 'write' | 'owner';

// Response wrapper for consistent API responses
export interface McpResponse<T> {
  success: boolean;
  data?: T;
  error?: string;
  metadata?: {
    requestId?: string;
    timestamp: string;
    source: 'onedrive' | 'sharepoint' | 'excel';
  };
}