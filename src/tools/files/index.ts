/**
 * File management tools for OneDrive and SharePoint
 * Unified operations for both platforms using Microsoft Graph
 */

import { Tool } from '@modelcontextprotocol/sdk/types.js';
import { getGraphClient } from '../../graph/client.js';
import { DriveItem, Permission, UploadSession, GraphResponse } from '../../graph/models.js';
import { extractPaginatedResult, jsonTextResponse, toolErrorResponse } from '../../graph/contracts.js';
import { createUserFriendlyError } from '../../graph/error-handler.js';
import {
  buildDriveChildrenEndpoint,
  buildDriveItemEndpoint,
  buildDriveSearchEndpoint,
  describeDriveTarget
} from '../../graph/resource-resolver.js';

// Tool 1: List files and folders
export const listFiles: Tool = {
  name: 'list_files',
  description: 'List files and folders in OneDrive or SharePoint drive',
  inputSchema: {
    type: 'object',
    properties: {
      path: {
        type: 'string',
        description: 'Folder path (e.g., "/Documents" or "" for root)',
        default: ''
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional, if not provided uses personal OneDrive)'
      },
      driveId: {
        type: 'string',
        description: 'Drive ID for a specific OneDrive or SharePoint document library'
      },
      filter: {
        type: 'string',
        description: 'OData filter (e.g., "file ne null" for files only)',
      },
      orderBy: {
        type: 'string',
        description: 'Sort order (e.g., "name", "lastModifiedDateTime desc")',
        default: 'name'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of items to return',
        default: 100
      },
      pageToken: {
        type: 'string',
        description: 'Opaque pagination token from a previous response (Graph nextLink)'
      }
    }
  }
};

export async function handleListFiles(args: any) {
  try {
    const client = getGraphClient();
    const { path = '', siteId, driveId, filter, orderBy = 'name', limit = 100, pageToken } = args;

    const endpoint = pageToken || buildDriveChildrenEndpoint({ siteId, driveId, path });

    const params: any = {
      '$orderby': orderBy,
      '$top': limit.toString()
    };

    if (filter) {
      params['$filter'] = filter;
    }

    const response = await client.get<GraphResponse<DriveItem>>(endpoint, pageToken ? undefined : params);

    if (response.success && response.data) {
      const { items, pagination } = extractPaginatedResult(response.data, limit);

      return jsonTextResponse({
        target: describeDriveTarget({ siteId, driveId }),
        path: path || '/',
        itemCount: items.length,
        pagination,
        items: items.map((item: DriveItem) => ({
          id: item.id,
          name: item.name,
          size: item.size,
          type: item.file ? 'file' : 'folder',
          mimeType: item.file?.mimeType,
          lastModified: item.lastModifiedDateTime,
          webUrl: item.webUrl,
          path: item.parentReference.path,
          driveId: item.parentReference.driveId
        }))
      });
    }

    throw new Error('Failed to retrieve files');
  } catch (error) {
    return toolErrorResponse('list_files', error);
  }
}

// Tool 2: Download file
export const downloadFile: Tool = {
  name: 'download_file',
  description: 'Download a file from OneDrive or SharePoint',
  inputSchema: {
    type: 'object',
    properties: {
      fileId: {
        type: 'string',
        description: 'File item ID'
      },
      filePath: {
        type: 'string',
        description: 'Alternative: file path (e.g., "/Documents/report.xlsx")'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      driveId: {
        type: 'string',
        description: 'Drive ID for a specific document library (optional)'
      },
      outputPath: {
        type: 'string',
        description: 'Local path to save the file (optional)'
      }
    },
    required: [],
  }
};

export async function handleDownloadFile(args: any) {
  try {
    const client = getGraphClient();
    const { fileId, filePath, siteId, driveId, outputPath } = args;
    const endpoint = buildDriveItemEndpoint({ itemId: fileId, itemPath: filePath, siteId, driveId }, '/content');

    const response = await client.downloadFile(endpoint);

    if (response.success && response.data) {
      const buffer = response.data as Buffer;
      
      if (outputPath) {
        const fs = await import('fs');
        await fs.promises.writeFile(outputPath, buffer);
        
        return jsonTextResponse({
          success: true,
          target: describeDriveTarget({ siteId, driveId }),
          message: `File downloaded successfully to ${outputPath}`,
          size: buffer.length,
          path: outputPath
        });
      } else {
        return jsonTextResponse({
          success: true,
          target: describeDriveTarget({ siteId, driveId }),
          message: 'File downloaded successfully',
          size: buffer.length,
          contentType: 'application/octet-stream',
          data: `<${buffer.length} bytes of binary data>`
        });
      }
    }

    throw new Error('Failed to download file');
  } catch (error) {
    return toolErrorResponse('download_file', error);
  }
}

// Tool 3: Upload file
export const uploadFile: Tool = {
  name: 'upload_file',
  description: 'Upload a file to OneDrive or SharePoint',
  inputSchema: {
    type: 'object',
    properties: {
      localPath: {
        type: 'string',
        description: 'Local file path to upload'
      },
      remotePath: {
        type: 'string',
        description: 'Remote path where to upload (e.g., "/Documents/report.xlsx")'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      conflictBehavior: {
        type: 'string',
        enum: ['fail', 'replace', 'rename'],
        description: 'What to do if file already exists',
        default: 'rename'
      }
    },
    required: ['localPath', 'remotePath']
  }
};

export async function handleUploadFile(args: any) {
  try {
    const client = getGraphClient();
    const { localPath, remotePath, siteId, conflictBehavior = 'rename' } = args;

    // Import path helper utilities
    const { prepareUploadPath } = await import('../utils/path-helper.js');
    
    // Prepare and sanitize the upload path
    const pathPrep = await prepareUploadPath(remotePath, siteId);
    
    if (!pathPrep.success) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: false,
            error: 'Path preparation failed',
            details: pathPrep.error,
            originalPath: remotePath,
            changes: pathPrep.changes
          }, null, 2)
        }],
        isError: true
      };
    }

    // Use the sanitized path for upload
    let endpoint: string;
    if (siteId) {
      endpoint = `/sites/${siteId}/drive/root:/${pathPrep.sanitizedPath}:/content`;
    } else {
      endpoint = `/me/drive/root:/${pathPrep.sanitizedPath}:/content`;
    }

    const fileName = require('path').basename(pathPrep.sanitizedPath);
    
    const response = await client.uploadFile(endpoint, localPath, fileName, {
      conflictBehavior
    });

    if (response.success && response.data) {
      const item = response.data as DriveItem;
      
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'File uploaded successfully',
            pathChanges: pathPrep.changes,
            originalPath: pathPrep.originalPath,
            finalPath: pathPrep.sanitizedPath,
            file: {
              id: item.id,
              name: item.name,
              size: item.size,
              webUrl: item.webUrl,
              lastModified: item.lastModifiedDateTime
            }
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to upload file');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error uploading file: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 4: Create folder
export const createFolder: Tool = {
  name: 'create_folder',
  description: 'Create a new folder in OneDrive or SharePoint',
  inputSchema: {
    type: 'object',
    properties: {
      name: {
        type: 'string',
        description: 'Folder name'
      },
      parentPath: {
        type: 'string',
        description: 'Parent folder path (e.g., "/Documents" or "" for root)',
        default: ''
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      }
    },
    required: ['name']
  }
};

export async function handleCreateFolder(args: any) {
  try {
    const client = getGraphClient();
    const { name, parentPath = '', siteId } = args;

    let endpoint: string;
    if (siteId) {
      endpoint = parentPath 
        ? `/sites/${siteId}/drive/root:/${parentPath}:/children`
        : `/sites/${siteId}/drive/root/children`;
    } else {
      endpoint = parentPath
        ? `/me/drive/root:/${parentPath}:/children`
        : `/me/drive/root/children`;
    }

    const folderData = {
      name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename'
    };

    const response = await client.post<DriveItem>(endpoint, folderData);

    if (response.success && response.data) {
      const folder = response.data;
      
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'Folder created successfully',
            folder: {
              id: folder.id,
              name: folder.name,
              webUrl: folder.webUrl,
              path: folder.parentReference.path + '/' + folder.name
            }
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to create folder');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error creating folder: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 5: Move/rename item
export const moveItem: Tool = {
  name: 'move_item',
  description: 'Move or rename a file/folder in OneDrive or SharePoint',
  inputSchema: {
    type: 'object',
    properties: {
      itemId: {
        type: 'string',
        description: 'Item ID to move/rename'
      },
      newName: {
        type: 'string',
        description: 'New name for the item (optional)'
      },
      parentFolderId: {
        type: 'string',
        description: 'ID of destination folder (optional)'
      },
      parentFolderPath: {
        type: 'string',
        description: 'Path of destination folder (optional)'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      }
    },
    required: ['itemId']
  }
};

export async function handleMoveItem(args: any) {
  try {
    const client = getGraphClient();
    const { itemId, newName, parentFolderId, parentFolderPath, siteId } = args;

    let endpoint: string = siteId 
      ? `/sites/${siteId}/drive/items/${itemId}`
      : `/me/drive/items/${itemId}`;

    const updateData: any = {};

    if (newName) {
      updateData.name = newName;
    }

    if (parentFolderId) {
      updateData.parentReference = { id: parentFolderId };
    } else if (parentFolderPath) {
      // Need to resolve folder path to ID first
      const folderEndpoint = siteId
        ? `/sites/${siteId}/drive/root:/${parentFolderPath}`
        : `/me/drive/root:/${parentFolderPath}`;
      
      const folderResponse = await client.get<DriveItem>(folderEndpoint);
      if (folderResponse.success && folderResponse.data) {
        updateData.parentReference = { id: folderResponse.data.id };
      }
    }

    const response = await client.patch<DriveItem>(endpoint, updateData);

    if (response.success && response.data) {
      const item = response.data;
      
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'Item moved/renamed successfully',
            item: {
              id: item.id,
              name: item.name,
              webUrl: item.webUrl,
              path: item.parentReference.path + '/' + item.name
            }
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to move/rename item');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error moving/renaming item: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 6: Delete item
export const deleteItem: Tool = {
  name: 'delete_item',
  description: 'Delete a file or folder from OneDrive or SharePoint',
  inputSchema: {
    type: 'object',
    properties: {
      itemId: {
        type: 'string',
        description: 'Item ID to delete'
      },
      itemPath: {
        type: 'string',
        description: 'Alternative: item path to delete'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      permanent: {
        type: 'boolean',
        description: 'Permanently delete (bypass recycle bin)',
        default: false
      }
    },
    required: [],
  }
};

export async function handleDeleteItem(args: any) {
  try {
    const client = getGraphClient();
    const { itemId, itemPath, siteId, permanent = false } = args;

    let endpoint: string;
    if (itemId) {
      endpoint = siteId 
        ? `/sites/${siteId}/drive/items/${itemId}`
        : `/me/drive/items/${itemId}`;
    } else {
      endpoint = siteId
        ? `/sites/${siteId}/drive/root:/${itemPath}`
        : `/me/drive/root:/${itemPath}`;
    }

    const response = await client.delete(endpoint);

    if (response.success) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: permanent 
              ? 'Item permanently deleted' 
              : 'Item moved to recycle bin',
            itemId: itemId || 'path-based',
            itemPath: itemPath || 'id-based'
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to delete item');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error deleting item: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 7: Search files
export const searchFiles: Tool = {
  name: 'search_files',
  description: 'Search for files and folders in OneDrive or SharePoint',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Search query (file name, content, etc.)'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      driveId: {
        type: 'string',
        description: 'Drive ID for a specific document library (optional)'
      },
      fileTypes: {
        type: 'array',
        items: { type: 'string' },
        description: 'Filter by file types (e.g., ["xlsx", "docx"])'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of results',
        default: 50
      },
      pageToken: {
        type: 'string',
        description: 'Opaque pagination token from a previous response (Graph nextLink)'
      }
    },
    required: ['query']
  }
};

export async function handleSearchFiles(args: any) {
  try {
    const client = getGraphClient();
    const { query, siteId, driveId, fileTypes, limit = 50, pageToken } = args;
    const endpoint = pageToken || buildDriveSearchEndpoint({ siteId, driveId }, query);

    const params: any = {
      '$top': limit.toString()
    };

    // Add file type filter if specified
    if (fileTypes && fileTypes.length > 0) {
      const typeFilter = fileTypes.map((type: string) => 
        `endswith(name,'.${type}')`
      ).join(' or ');
      params['$filter'] = typeFilter;
    }

    const response = await client.get<GraphResponse<DriveItem>>(endpoint, pageToken ? undefined : params);

    if (response.success && response.data) {
      const { items, pagination } = extractPaginatedResult(response.data, limit);
      
      return jsonTextResponse({
        query,
        target: describeDriveTarget({ siteId, driveId }),
        resultCount: items.length,
        pagination,
        results: items.map((item: DriveItem) => ({
          id: item.id,
          name: item.name,
          size: item.size,
          type: item.file ? 'file' : 'folder',
          mimeType: item.file?.mimeType,
          lastModified: item.lastModifiedDateTime,
          webUrl: item.webUrl,
          path: item.parentReference.path,
          driveId: item.parentReference.driveId
        }))
      });
    }

    throw new Error('Search failed');
  } catch (error) {
    return toolErrorResponse('search_files', error);
  }
}

// Tool 8: Get file metadata
export const getFileMetadata: Tool = {
  name: 'get_file_metadata',
  description: 'Get detailed metadata for a file or folder',
  inputSchema: {
    type: 'object',
    properties: {
      itemId: {
        type: 'string',
        description: 'Item ID'
      },
      itemPath: {
        type: 'string',
        description: 'Alternative: item path'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      includeVersions: {
        type: 'boolean',
        description: 'Include version history',
        default: false
      }
    },
    required: [],
  }
};

export async function handleGetFileMetadata(args: any) {
  try {
    const client = getGraphClient();
    const { itemId, itemPath, siteId, includeVersions = false } = args;

    let endpoint: string;
    if (itemId) {
      endpoint = siteId 
        ? `/sites/${siteId}/drive/items/${itemId}`
        : `/me/drive/items/${itemId}`;
    } else {
      endpoint = siteId
        ? `/sites/${siteId}/drive/root:/${itemPath}`
        : `/me/drive/root:/${itemPath}`;
    }

    const response = await client.get<DriveItem>(endpoint);

    if (response.success && response.data) {
      const item = response.data;
      
      const metadata: any = {
        id: item.id,
        name: item.name,
        size: item.size,
        type: item.file ? 'file' : 'folder',
        created: item.createdDateTime,
        lastModified: item.lastModifiedDateTime,
        webUrl: item.webUrl,
        downloadUrl: item['@microsoft.graph.downloadUrl'],
        path: item.parentReference.path,
        createdBy: item.createdBy.user.displayName,
        lastModifiedBy: item.lastModifiedBy.user.displayName
      };

      if (item.file) {
        metadata.mimeType = item.file.mimeType;
        metadata.hashes = item.file.hashes;
      }

      if (item.folder) {
        metadata.childCount = item.folder.childCount;
      }

      if (includeVersions && item.file) {
        try {
          const versionsEndpoint = `${endpoint}/versions`;
          const versionsResponse = await client.get(versionsEndpoint);
          if (versionsResponse.success) {
            metadata.versions = versionsResponse.data;
          }
        } catch (versionError) {
          metadata.versionsError = 'Could not retrieve versions';
        }
      }

      return {
        content: [{
          type: 'text',
          text: JSON.stringify(metadata, null, 2)
        }]
      };
    }

    throw new Error('Failed to get metadata');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error getting metadata: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 9: Share file/folder
export const shareItem: Tool = {
  name: 'share_item',
  description: 'Create a sharing link for a file or folder',
  inputSchema: {
    type: 'object',
    properties: {
      itemId: {
        type: 'string',
        description: 'Item ID to share'
      },
      itemPath: {
        type: 'string',
        description: 'Alternative: item path to share'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      type: {
        type: 'string',
        enum: ['view', 'edit', 'embed'],
        description: 'Type of sharing link',
        default: 'view'
      },
      scope: {
        type: 'string',
        enum: ['anonymous', 'organization', 'users'],
        description: 'Who can access the link',
        default: 'organization'
      },
      expirationDateTime: {
        type: 'string',
        description: 'Link expiration (ISO 8601 format, optional)'
      },
      password: {
        type: 'string',
        description: 'Password protection (optional)'
      }
    },
    required: []
  }
};

export async function handleShareItem(args: any) {
  try {
    const client = getGraphClient();
    const { 
      itemId, 
      itemPath, 
      siteId, 
      type = 'view', 
      scope = 'organization',
      expirationDateTime,
      password 
    } = args;

    let endpoint: string;
    if (itemId) {
      endpoint = siteId 
        ? `/sites/${siteId}/drive/items/${itemId}/createLink`
        : `/me/drive/items/${itemId}/createLink`;
    } else {
      endpoint = siteId
        ? `/sites/${siteId}/drive/root:/${itemPath}:/createLink`
        : `/me/drive/root:/${itemPath}:/createLink`;
    }

    const shareData: any = {
      type,
      scope
    };

    if (expirationDateTime) {
      shareData.expirationDateTime = expirationDateTime;
    }

    if (password) {
      shareData.password = password;
    }

    const response = await client.post<Permission>(endpoint, shareData);

    if (response.success && response.data) {
      const permission = response.data;
      
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'Sharing link created successfully',
            link: {
              id: permission.id,
              url: permission.link?.webUrl,
              type: permission.link?.type,
              scope: permission.link?.scope,
              expirationDateTime: permission.expirationDateTime,
              hasPassword: permission.hasPassword
            }
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to create sharing link');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error creating sharing link: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 10: Copy item
export const copyItem: Tool = {
  name: 'copy_item',
  description: 'Copy a file or folder in OneDrive or SharePoint',
  inputSchema: {
    type: 'object',
    properties: {
      itemId: {
        type: 'string',
        description: 'Item ID to copy'
      },
      itemPath: {
        type: 'string',
        description: 'Alternative: item path to copy'
      },
      destinationFolderId: {
        type: 'string',
        description: 'Destination folder ID'
      },
      destinationFolderPath: {
        type: 'string',
        description: 'Alternative: destination folder path'
      },
      newName: {
        type: 'string',
        description: 'New name for the copied item (optional)'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      destinationSiteId: {
        type: 'string',
        description: 'Destination SharePoint site ID (optional)'
      }
    },
    required: [],
  }
};

export async function handleCopyItem(args: any) {
  try {
    const client = getGraphClient();
    const { 
      itemId, 
      itemPath, 
      destinationFolderId,
      destinationFolderPath,
      newName,
      siteId,
      destinationSiteId 
    } = args;

    let endpoint: string;
    if (itemId) {
      endpoint = siteId 
        ? `/sites/${siteId}/drive/items/${itemId}/copy`
        : `/me/drive/items/${itemId}/copy`;
    } else {
      endpoint = siteId
        ? `/sites/${siteId}/drive/root:/${itemPath}:/copy`
        : `/me/drive/root:/${itemPath}:/copy`;
    }

    const copyData: any = {
      parentReference: {}
    };

    if (destinationFolderId) {
      copyData.parentReference.id = destinationFolderId;
    } else if (destinationFolderPath) {
      // Resolve destination path to ID
      const destEndpoint = destinationSiteId
        ? `/sites/${destinationSiteId}/drive/root:/${destinationFolderPath}`
        : `/me/drive/root:/${destinationFolderPath}`;
      
      const destResponse = await client.get<DriveItem>(destEndpoint);
      if (destResponse.success && destResponse.data) {
        copyData.parentReference.id = destResponse.data.id;
      }
    }

    if (destinationSiteId) {
      copyData.parentReference.driveId = destinationSiteId;
    }

    if (newName) {
      copyData.name = newName;
    }

    const response = await client.post(endpoint, copyData);

    if (response.success) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'Item copy initiated successfully',
            note: 'Copy operation is asynchronous and may take some time to complete',
            itemId: itemId || 'path-based',
            itemPath: itemPath || 'id-based',
            destinationFolderId,
            destinationFolderPath,
            newName
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to copy item');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error copying item: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Export all tools and handlers
export const fileTools = [
  listFiles,
  downloadFile,
  uploadFile,
  createFolder,
  moveItem,
  deleteItem,
  searchFiles,
  getFileMetadata,
  shareItem,
  copyItem
];

export const fileHandlers = {
  list_files: handleListFiles,
  download_file: handleDownloadFile,
  upload_file: handleUploadFile,
  create_folder: handleCreateFolder,
  move_item: handleMoveItem,
  delete_item: handleDeleteItem,
  search_files: handleSearchFiles,
  get_file_metadata: handleGetFileMetadata,
  share_item: handleShareItem,
  copy_item: handleCopyItem
};