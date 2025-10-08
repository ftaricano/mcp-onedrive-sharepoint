# MCP OneDrive/SharePoint Server

A comprehensive Model Context Protocol (MCP) server providing unified access to OneDrive and SharePoint through Microsoft Graph API.

## Features

### 🔄 Unified File Operations
- **OneDrive & SharePoint**: Single API for both personal OneDrive and SharePoint document libraries
- **Complete CRUD**: Upload, download, move, rename, delete, search files and folders
- **Advanced Search**: Content search with metadata filtering
- **Sharing Management**: Create and manage sharing links with permissions
- **Large File Support**: Resumable uploads for files > 4MB

### 💼 File Management
- **Multi-format Support**: Handle various file types including Office documents
- **Metadata Access**: Read and write file properties and custom metadata
- **Version Control**: Access file version history and restore previous versions
- **Bulk Operations**: Perform operations on multiple files efficiently

### 📋 SharePoint Lists Management
- **Site Discovery**: List and search SharePoint sites
- **List Operations**: Create, read, update, delete SharePoint lists
- **Item Management**: Full CRUD operations on list items
- **Schema Access**: Get list schemas and column definitions
- **Content Types**: Work with SharePoint content types

### 🏢 Business Account Optimizations
- **Multi-Tenant Support**: Works with both personal and business accounts
- **Device Code Flow**: Secure authentication optimized for CLI/MCP usage
- **Enterprise Features**: Advanced permissions, compliance, audit logging
- **Performance**: Intelligent caching, batch operations, retry logic

## Installation

### Prerequisites
- Node.js 18.0.0 or higher
- A Microsoft Azure application registration

### Setup

1. **Clone and install dependencies:**
```bash
git clone <repository-url>
cd mcp-onedrive-sharepoint
npm install
```

2. **Azure App Registration:**
   - Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
   - Create a new registration or use existing one
   - Configure authentication:
     - Platform: Mobile and desktop applications
     - Redirect URI: `http://localhost` (for device code flow)
   - API Permissions:
     - Microsoft Graph:
       - `Files.ReadWrite.All` (Application)
       - `Sites.ReadWrite.All` (Application)
       - `User.Read` (Delegated)
       - `offline_access` (Delegated)

3. **Environment Configuration:**
```bash
cp .env.example .env
# Edit .env with your Azure app credentials
```

4. **Build the server:**
```bash
npm run build
```

## Configuration

### Environment Variables

```bash
# Required
MICROSOFT_CLIENT_ID=your-application-client-id

# Optional (defaults to 'common' for multi-tenant)
MICROSOFT_TENANT_ID=common  # or specific tenant ID for single-tenant
```

### Business vs Personal Accounts

**Multi-Tenant (Personal + Business):**
```bash
MICROSOFT_TENANT_ID=common
```

**Single Business Tenant:**
```bash
MICROSOFT_TENANT_ID=your-tenant-id-here
```

## Usage

### Start the MCP Server
```bash
npm start
```

### Authentication
```bash
# First time authentication
use tool: authenticate

# Force re-authentication
use tool: authenticate with forceReauth: true
```

The server uses device code flow - you'll get a code to enter at https://microsoft.com/devicelogin.

## Tool Reference

### File Management (10 tools)

#### `list_files`
List files and folders in OneDrive or SharePoint.
```json
{
  "driveId": "optional-drive-id",
  "folderId": "optional-folder-id", 
  "folderPath": "/Documents/Projects",
  "top": 100,
  "orderBy": "name",
  "filter": "file ne null"
}
```

#### `upload_file`
Upload files with automatic large file handling.
```json
{
  "driveId": "optional-drive-id",
  "folderPath": "/Documents",
  "fileName": "report.xlsx",
  "content": "base64-encoded-content",
  "conflictBehavior": "rename"
}
```

#### `download_file`
Download files as base64-encoded content.
```json
{
  "driveId": "optional-drive-id",
  "fileId": "file-id-here",
  "filePath": "/Documents/report.xlsx"
}
```

#### `search_files`
Advanced file search with filters.
```json
{
  "query": "quarterly report",
  "fileType": "xlsx",
  "modifiedAfter": "2024-01-01T00:00:00Z",
  "top": 50
}
```

#### `share_file`
Create sharing links with permissions.
```json
{
  "itemId": "file-id-here",
  "type": "edit",
  "scope": "organization",
  "expirationDateTime": "2024-12-31T23:59:59Z"
}
```

### SharePoint Lists (8 tools)

#### `list_sharepoint_sites`
Discover available SharePoint sites.
```json
{
  "search": "project sites",
  "top": 50
}
```

#### `get_list_items`
Get items from SharePoint lists.
```json
{
  "siteId": "site-id-here",
  "listId": "list-id-here",
  "filter": "Status eq 'Active'",
  "orderBy": "Modified desc",
  "top": 100
}
```

#### `create_list_item`
Create new SharePoint list items.
```json
{
  "siteId": "site-id-here", 
  "listId": "list-id-here",
  "fields": {
    "Title": "New Project",
    "Status": "Planning",
    "DueDate": "2024-12-31"
  }
}
```

### Excel Integration (12 tools)

#### `get_workbook_info`
Get Excel workbook metadata without downloading.
```json
{
  "driveId": "optional-drive-id",
  "fileId": "excel-file-id",
  "filePath": "/Documents/budget.xlsx"
}
```

#### `get_worksheet_data`
Extract data from Excel worksheets.
```json
{
  "fileId": "excel-file-id",
  "worksheetId": "Sheet1",
  "range": "A1:E10",
  "includeFormulas": true
}
```

#### `set_range_values`
Update Excel ranges directly.
```json
{
  "fileId": "excel-file-id",
  "worksheetId": "Sheet1", 
  "range": "A1:B2",
  "values": [["Name", "Value"], ["Product A", 100]]
}
```

#### `get_table_data`
Access Excel table data.
```json
{
  "fileId": "excel-file-id",
  "worksheetId": "Sheet1",
  "tableId": "Table1",
  "includeHeaders": true
}
```

#### `execute_formula`
Run Excel formulas and get results.
```json
{
  "fileId": "excel-file-id",
  "worksheetId": "Sheet1",
  "formula": "=SUM(A1:A10)"
}
```

#### `create_chart`
Create charts in Excel programmatically.
```json
{
  "fileId": "excel-file-id",
  "worksheetId": "Sheet1",
  "chartType": "ColumnClustered",
  "sourceData": "A1:B10",
  "title": "Sales Data"
}
```

### Utility Tools (5 tools)

#### `get_drive_info`
Get drive information and storage quota.
```json
{
  "driveId": "optional-drive-id"
}
```

#### `search_content`
Universal content search across all accessible resources.
```json
{
  "query": "project timeline",
  "entityTypes": ["driveItem", "listItem", "site"],
  "top": 50
}
```

## Architecture

### Core Components

```
src/
├── server.ts              # Main MCP server
├── auth/
│   └── microsoft-graph-auth.ts  # OAuth 2.0 + device code flow
├── graph/
│   ├── client.ts          # Microsoft Graph API client
│   └── models.ts          # TypeScript models
├── tools/
│   ├── files/             # File management tools
│   ├── sharepoint/        # SharePoint list tools  
│   ├── excel/             # Excel integration tools
│   └── utils/             # Utility tools
└── config/
    ├── endpoints.ts       # API endpoint definitions
    └── scopes.ts          # OAuth scopes configuration
```

### Authentication Flow

1. **Device Code Flow**: Optimized for CLI/MCP environments
2. **Secure Storage**: Tokens stored in system keychain
3. **Auto-Refresh**: Automatic token refresh with retry logic
4. **Business Ready**: Multi-tenant and single-tenant support

### Error Handling

- **Retry Logic**: Exponential backoff for rate limits
- **Graceful Degradation**: Fallback strategies for failures
- **Detailed Errors**: Clear error messages with context
- **Business Context**: Enterprise-specific error handling

## Business Account Features

### Multi-Tenant Support
Works seamlessly with:
- Personal Microsoft accounts (outlook.com, hotmail.com)
- Business accounts (Office 365, Azure AD)
- Government clouds (with configuration)

### Enterprise Optimizations
- **Admin Consent**: Support for admin-consented permissions
- **Compliance**: Audit logging and data classification awareness
- **Performance**: Batch operations and intelligent caching
- **Security**: Minimal scope requests and secure token storage

### SharePoint Integration
- **Site Discovery**: Find sites across the organization
- **Document Libraries**: Access as unified drives
- **List Management**: Full CRUD on business lists
- **Content Types**: Support for SharePoint content types

## Development

### Build and Test
```bash
# Development mode
npm run dev

# Build for production  
npm run build

# Run tests
npm test

# Linting
npm run lint
```

### Authentication Setup
```bash
# Initial authentication setup
npm run auth
```

### MCP Integration
Add to your MCP client configuration:
```json
{
  "mcpServers": {
    "onedrive-sharepoint": {
      "command": "node",
      "args": ["/path/to/mcp-onedrive-sharepoint/build/server.js"]
    }
  }
}
```

## Troubleshooting

### Authentication Issues
- Ensure correct client ID and tenant configuration
- Check that required permissions are granted
- For business accounts, admin consent may be required

### Permission Errors
- Verify Azure app has correct Graph API permissions
- Business accounts may need additional scopes
- Check user permissions in SharePoint/OneDrive

### File Access Issues
- Ensure files exist and user has access
- Check file locks (Excel files may be locked by other users)
- Verify drive ID resolution for SharePoint document libraries

## Security Considerations

### Token Management
- Tokens stored securely in system keychain
- Automatic refresh with secure storage
- No tokens logged or exposed in errors

### Permissions
- Follows principle of least privilege
- Requests only necessary scopes
- Supports admin consent for business scenarios

### Data Protection
- No local file caching by default
- Secure HTTPS connections only
- Enterprise compliance awareness

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## Support

For issues and questions:
- Check the troubleshooting section
- Review Microsoft Graph API documentation
- Open an issue on GitHub