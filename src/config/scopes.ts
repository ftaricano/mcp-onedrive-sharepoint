/**
 * Microsoft Graph API scopes configuration for OneDrive/SharePoint/Excel integration
 */

export const GRAPH_SCOPES = {
  // File and drive access
  FILES_READ: 'Files.Read',
  FILES_READ_ALL: 'Files.Read.All',
  FILES_READWRITE: 'Files.ReadWrite',
  FILES_READWRITE_ALL: 'Files.ReadWrite.All',
  
  // SharePoint access
  SITES_READ_ALL: 'Sites.Read.All',
  SITES_READWRITE_ALL: 'Sites.ReadWrite.All',
  SITES_MANAGE_ALL: 'Sites.Manage.All',
  
  // User profile
  USER_READ: 'User.Read',
  
  // Directory access (for business accounts)
  DIRECTORY_READ_ALL: 'Directory.Read.All',
  
  // Application permissions (for unattended scenarios)
  APP_FILES_READ_ALL: 'Files.Read.All',
  APP_FILES_READWRITE_ALL: 'Files.ReadWrite.All',
  APP_SITES_READ_ALL: 'Sites.Read.All',
  APP_SITES_READWRITE_ALL: 'Sites.ReadWrite.All'
} as const;

// Scope configurations for different use cases
export const SCOPE_CONFIGURATIONS = {
  // Personal OneDrive access
  PERSONAL: [
    GRAPH_SCOPES.USER_READ,
    GRAPH_SCOPES.FILES_READWRITE,
  ],
  
  // Business OneDrive + SharePoint
  BUSINESS: [
    GRAPH_SCOPES.USER_READ,
    GRAPH_SCOPES.FILES_READWRITE_ALL,
    GRAPH_SCOPES.SITES_READWRITE_ALL,
  ],
  
  // Full enterprise access
  ENTERPRISE: [
    GRAPH_SCOPES.USER_READ,
    GRAPH_SCOPES.FILES_READWRITE_ALL,
    GRAPH_SCOPES.SITES_READWRITE_ALL,
    GRAPH_SCOPES.DIRECTORY_READ_ALL,
  ],
  
  // Application-only (service principal)
  APPLICATION: [
    GRAPH_SCOPES.APP_FILES_READWRITE_ALL,
    GRAPH_SCOPES.APP_SITES_READWRITE_ALL,
  ]
} as const;

// Default scope for device code flow
export const DEFAULT_SCOPES = SCOPE_CONFIGURATIONS.BUSINESS;

export type ScopeConfiguration = keyof typeof SCOPE_CONFIGURATIONS;