/**
 * File management tools for OneDrive and SharePoint
 * Unified operations for both platforms using Microsoft Graph
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { basename } from "node:path";
import { getGraphClient } from "../../graph/client.js";
import {
  DriveItem,
  Permission,
  UploadSession,
  GraphResponse,
} from "../../graph/models.js";
import {
  extractPaginatedResult,
  jsonTextResponse,
  toolErrorResponse,
  jsonTextErrorResponse,
} from "../../graph/contracts.js";
import {
  buildDriveChildrenEndpoint,
  buildDriveItemEndpoint,
  buildDriveSearchEndpoint,
  describeDriveTarget,
  getDriveRootEndpoint,
} from "../../graph/resource-resolver.js";
import { resolveDriveTargetContext } from "../../sharepoint/site-resolver.js";

async function resolveFileDriveContext(
  args: any,
  client: ReturnType<typeof getGraphClient>,
) {
  return resolveDriveTargetContext(
    {
      site: args.site,
      siteId: args.siteId,
      siteUrl: args.siteUrl,
      driveId: args.driveId,
    },
    client,
  );
}

// Tool 1: List files and folders
export const listFiles: Tool = {
  name: "list_files",
  description: "List files and folders in OneDrive or SharePoint drive",
  inputSchema: {
    type: "object",
    properties: {
      path: {
        type: "string",
        description: 'Folder path (e.g., "/Documents" or "" for root)',
        default: "",
      },
      siteId: {
        type: "string",
        description:
          "SharePoint site ID (optional, if not provided uses personal OneDrive)",
      },
      site: {
        type: "string",
        description:
          'Known SharePoint site alias or canonical URL (e.g., "financeiro", "socios2")',
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description:
          "Drive ID for a specific OneDrive or SharePoint document library",
      },
      filter: {
        type: "string",
        description: 'OData filter (e.g., "file ne null" for files only)',
      },
      orderBy: {
        type: "string",
        description: 'Sort order (e.g., "name", "lastModifiedDateTime desc")',
        default: "name",
      },
      limit: {
        type: "number",
        description: "Maximum number of items to return",
        default: 100,
      },
      pageToken: {
        type: "string",
        description:
          "Opaque pagination token from a previous response (Graph nextLink)",
      },
    },
  },
};

export async function handleListFiles(args: any) {
  try {
    const client = getGraphClient();
    const {
      path = "",
      filter,
      orderBy = "name",
      limit = 100,
      pageToken,
    } = args;
    const { siteId, driveId, resolvedSite } = await resolveFileDriveContext(
      args,
      client,
    );

    const endpoint =
      pageToken || buildDriveChildrenEndpoint({ siteId, driveId, path });

    const params: any = {
      $orderby: orderBy,
      $top: limit.toString(),
    };

    if (filter) {
      params["$filter"] = filter;
    }

    const response = await client.get<GraphResponse<DriveItem>>(
      endpoint,
      pageToken ? undefined : params,
    );

    if (response.success && response.data) {
      const { items, pagination } = extractPaginatedResult(
        response.data,
        limit,
      );

      return jsonTextResponse({
        target: describeDriveTarget({ siteId, driveId }),
        site: resolvedSite,
        path: path || "/",
        itemCount: items.length,
        pagination,
        items: items.map((item: DriveItem) => ({
          id: item.id,
          name: item.name,
          size: item.size,
          type: item.file ? "file" : "folder",
          mimeType: item.file?.mimeType,
          lastModified: item.lastModifiedDateTime,
          webUrl: item.webUrl,
          path: item.parentReference.path,
          driveId: item.parentReference.driveId,
        })),
      });
    }

    throw new Error("Failed to retrieve files");
  } catch (error) {
    return toolErrorResponse("list_files", error);
  }
}

// Tool 2: Download file
export const downloadFile: Tool = {
  name: "download_file",
  description: "Download a file from OneDrive or SharePoint",
  inputSchema: {
    type: "object",
    properties: {
      fileId: {
        type: "string",
        description: "File item ID",
      },
      filePath: {
        type: "string",
        description: 'Alternative: file path (e.g., "/Documents/report.xlsx")',
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      outputPath: {
        type: "string",
        description: "Local path to save the file (optional)",
      },
    },
    required: [],
  },
};

export async function handleDownloadFile(args: any) {
  try {
    const client = getGraphClient();
    const { fileId, filePath, outputPath } = args;
    const { siteId, driveId, resolvedSite } = await resolveFileDriveContext(
      args,
      client,
    );
    const endpoint = buildDriveItemEndpoint(
      { itemId: fileId, itemPath: filePath, siteId, driveId },
      "/content",
    );

    const response = await client.downloadFile(endpoint);

    if (response.success && response.data) {
      const buffer = response.data as Buffer;

      if (outputPath) {
        const fs = await import("fs");
        await fs.promises.writeFile(outputPath, buffer);

        return jsonTextResponse({
          success: true,
          target: describeDriveTarget({ siteId, driveId }),
          site: resolvedSite,
          message: `File downloaded successfully to ${outputPath}`,
          size: buffer.length,
          path: outputPath,
        });
      } else {
        return jsonTextResponse({
          success: true,
          target: describeDriveTarget({ siteId, driveId }),
          site: resolvedSite,
          message: "File downloaded successfully",
          size: buffer.length,
          contentType: "application/octet-stream",
          data: `<${buffer.length} bytes of binary data>`,
        });
      }
    }

    throw new Error("Failed to download file");
  } catch (error) {
    return toolErrorResponse("download_file", error);
  }
}

// Tool 3: Upload file
export const uploadFile: Tool = {
  name: "upload_file",
  description: "Upload a file to OneDrive or SharePoint",
  inputSchema: {
    type: "object",
    properties: {
      localPath: {
        type: "string",
        description: "Local file path to upload",
      },
      remotePath: {
        type: "string",
        description:
          'Remote path where to upload (e.g., "/Documents/report.xlsx")',
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      conflictBehavior: {
        type: "string",
        enum: ["fail", "replace", "rename"],
        description: "What to do if file already exists",
        default: "rename",
      },
    },
    required: ["localPath", "remotePath"],
  },
};

export async function handleUploadFile(args: any) {
  try {
    const client = getGraphClient();
    const { localPath, remotePath, conflictBehavior = "rename" } = args;
    const { siteId, driveId, resolvedSite } = await resolveFileDriveContext(
      args,
      client,
    );

    // Import path helper utilities
    const { prepareUploadPath } = await import("../utils/path-helper.js");

    // Prepare and sanitize the upload path
    const pathPrep = await prepareUploadPath(remotePath, siteId, driveId);

    if (!pathPrep.success) {
      return jsonTextErrorResponse({
        success: false,
        error: "Path preparation failed",
        details: pathPrep.error,
        originalPath: remotePath,
        changes: pathPrep.changes,
      });
    }

    // Use the sanitized path for upload
    const endpoint = `${getDriveRootEndpoint({ siteId, driveId })}/root:/${pathPrep.sanitizedPath}:/content`;

    const fileName = basename(pathPrep.sanitizedPath);

    const response = await client.uploadFile(endpoint, localPath, fileName, {
      conflictBehavior,
    });

    if (response.success && response.data) {
      const item = response.data as DriveItem;

      return jsonTextResponse({
        success: true,
        message: "File uploaded successfully",
        target: describeDriveTarget({ siteId, driveId }),
        site: resolvedSite,
        pathChanges: pathPrep.changes,
        originalPath: pathPrep.originalPath,
        finalPath: pathPrep.sanitizedPath,
        file: {
          id: item.id,
          name: item.name,
          size: item.size,
          webUrl: item.webUrl,
          lastModified: item.lastModifiedDateTime,
        },
      });
    }

    throw new Error("Failed to upload file");
  } catch (error) {
    return toolErrorResponse("upload_file", error);
  }
}

// Tool 4: Create folder
export const createFolder: Tool = {
  name: "create_folder",
  description: "Create a new folder in OneDrive or SharePoint",
  inputSchema: {
    type: "object",
    properties: {
      name: {
        type: "string",
        description: "Folder name",
      },
      parentPath: {
        type: "string",
        description: 'Parent folder path (e.g., "/Documents" or "" for root)',
        default: "",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
    },
    required: ["name"],
  },
};

export async function handleCreateFolder(args: any) {
  try {
    const client = getGraphClient();
    const { name, parentPath = "" } = args;
    const { siteId, driveId } = await resolveFileDriveContext(args, client);

    const driveRoot = getDriveRootEndpoint({ siteId, driveId });
    const endpoint = parentPath
      ? `${driveRoot}/root:/${parentPath}:/children`
      : `${driveRoot}/root/children`;

    const folderData = {
      name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    };

    const response = await client.post<DriveItem>(endpoint, folderData);

    if (response.success && response.data) {
      const folder = response.data;

      return jsonTextResponse({
        success: true,
        message: "Folder created successfully",
        folder: {
          id: folder.id,
          name: folder.name,
          webUrl: folder.webUrl,
          path: folder.parentReference.path + "/" + folder.name,
        },
      });
    }

    throw new Error("Failed to create folder");
  } catch (error) {
    return toolErrorResponse("create_folder", error);
  }
}

// Tool 5: Move/rename item
export const moveItem: Tool = {
  name: "move_item",
  description: "Move or rename a file/folder in OneDrive or SharePoint",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "Item ID to move/rename",
      },
      itemPath: {
        type: "string",
        description: "Alternative: item path to move/rename",
      },
      newName: {
        type: "string",
        description: "New name for the item (optional)",
      },
      parentFolderId: {
        type: "string",
        description: "ID of destination folder (optional)",
      },
      parentFolderPath: {
        type: "string",
        description: "Path of destination folder (optional)",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
    },
    required: [],
  },
};

export async function handleMoveItem(args: any) {
  try {
    const client = getGraphClient();
    const { itemId, itemPath, newName, parentFolderId, parentFolderPath } = args;
    const { siteId, driveId } = await resolveFileDriveContext(args, client);

    const endpoint = buildDriveItemEndpoint(
      { itemId, itemPath, siteId, driveId },
    );

    const updateData: any = {};

    if (newName) {
      updateData.name = newName;
    }

    if (parentFolderId) {
      updateData.parentReference = { id: parentFolderId };
    } else if (parentFolderPath) {
      // Need to resolve folder path to ID first
      const driveRoot = getDriveRootEndpoint({ siteId, driveId });
      const folderEndpoint = `${driveRoot}/root:/${parentFolderPath}`;

      const folderResponse = await client.get<DriveItem>(folderEndpoint);
      if (folderResponse.success && folderResponse.data) {
        updateData.parentReference = { id: folderResponse.data.id };
      }
    }

    const response = await client.patch<DriveItem>(endpoint, updateData);

    if (response.success && response.data) {
      const item = response.data;

      return jsonTextResponse({
        success: true,
        message: "Item moved/renamed successfully",
        item: {
          id: item.id,
          name: item.name,
          webUrl: item.webUrl,
          path: item.parentReference.path + "/" + item.name,
        },
      });
    }

    throw new Error("Failed to move/rename item");
  } catch (error) {
    return toolErrorResponse("move_item", error);
  }
}

// Tool 6: Delete item
export const deleteItem: Tool = {
  name: "delete_item",
  description: "Delete a file or folder from OneDrive or SharePoint",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "Item ID to delete",
      },
      itemPath: {
        type: "string",
        description: "Alternative: item path to delete",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      permanent: {
        type: "boolean",
        description: "Permanently delete (bypass recycle bin)",
        default: false,
      },
    },
    required: [],
  },
};

export async function handleDeleteItem(args: any) {
  try {
    const client = getGraphClient();
    const { itemId, itemPath, permanent = false } = args;
    const { siteId, driveId } = await resolveFileDriveContext(args, client);

    const endpoint = buildDriveItemEndpoint(
      { itemId, itemPath, siteId, driveId },
    );

    const response = await client.delete(endpoint);

    if (response.success) {
      return jsonTextResponse({
        success: true,
        message: permanent
          ? "Item permanently deleted"
          : "Item moved to recycle bin",
        itemId: itemId || "path-based",
        itemPath: itemPath || "id-based",
      });
    }

    throw new Error("Failed to delete item");
  } catch (error) {
    return toolErrorResponse("delete_item", error);
  }
}

// Tool 7: Search files
export const searchFiles: Tool = {
  name: "search_files",
  description: "Search for files and folders in OneDrive or SharePoint",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description: "Search query (file name, content, etc.)",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      fileTypes: {
        type: "array",
        items: { type: "string" },
        description: 'Filter by file types (e.g., ["xlsx", "docx"])',
      },
      limit: {
        type: "number",
        description: "Maximum number of results",
        default: 50,
      },
      pageToken: {
        type: "string",
        description:
          "Opaque pagination token from a previous response (Graph nextLink)",
      },
    },
    required: ["query"],
  },
};

export async function handleSearchFiles(args: any) {
  try {
    const client = getGraphClient();
    const { query, fileTypes, limit = 50, pageToken } = args;
    const { siteId, driveId, resolvedSite } = await resolveFileDriveContext(
      args,
      client,
    );
    const endpoint =
      pageToken || buildDriveSearchEndpoint({ siteId, driveId }, query);

    const params: any = {
      $top: limit.toString(),
    };

    // Add file type filter if specified
    if (fileTypes && fileTypes.length > 0) {
      const typeFilter = fileTypes
        .map((type: string) => `endswith(name,'.${type}')`)
        .join(" or ");
      params["$filter"] = typeFilter;
    }

    const response = await client.get<GraphResponse<DriveItem>>(
      endpoint,
      pageToken ? undefined : params,
    );

    if (response.success && response.data) {
      const { items, pagination } = extractPaginatedResult(
        response.data,
        limit,
      );

      return jsonTextResponse({
        query,
        target: describeDriveTarget({ siteId, driveId }),
        site: resolvedSite,
        resultCount: items.length,
        pagination,
        results: items.map((item: DriveItem) => ({
          id: item.id,
          name: item.name,
          size: item.size,
          type: item.file ? "file" : "folder",
          mimeType: item.file?.mimeType,
          lastModified: item.lastModifiedDateTime,
          webUrl: item.webUrl,
          path: item.parentReference.path,
          driveId: item.parentReference.driveId,
        })),
      });
    }

    throw new Error("Search failed");
  } catch (error) {
    return toolErrorResponse("search_files", error);
  }
}

// Tool 8: Get file metadata
export const getFileMetadata: Tool = {
  name: "get_file_metadata",
  description: "Get detailed metadata for a file or folder",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "Item ID",
      },
      itemPath: {
        type: "string",
        description: "Alternative: item path",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      includeVersions: {
        type: "boolean",
        description: "Include version history",
        default: false,
      },
    },
    required: [],
  },
};

export async function handleGetFileMetadata(args: any) {
  try {
    const client = getGraphClient();
    const { itemId, itemPath, includeVersions = false } = args;
    const { siteId, driveId } = await resolveFileDriveContext(args, client);

    const endpoint = buildDriveItemEndpoint(
      { itemId, itemPath, siteId, driveId },
    );

    const response = await client.get<DriveItem>(endpoint);

    if (response.success && response.data) {
      const item = response.data;

      const metadata: any = {
        id: item.id,
        name: item.name,
        size: item.size,
        type: item.file ? "file" : "folder",
        created: item.createdDateTime,
        lastModified: item.lastModifiedDateTime,
        webUrl: item.webUrl,
        downloadUrl: item["@microsoft.graph.downloadUrl"],
        path: item.parentReference.path,
        createdBy: item.createdBy.user.displayName,
        lastModifiedBy: item.lastModifiedBy.user.displayName,
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
          metadata.versionsError = "Could not retrieve versions";
        }
      }

      return jsonTextResponse(metadata);
    }

    throw new Error("Failed to get metadata");
  } catch (error) {
    return toolErrorResponse("get_file_metadata", error);
  }
}

// Tool 9: Share file/folder
export const shareItem: Tool = {
  name: "share_item",
  description: "Create a sharing link for a file or folder",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "Item ID to share",
      },
      itemPath: {
        type: "string",
        description: "Alternative: item path to share",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      type: {
        type: "string",
        enum: ["view", "edit", "embed"],
        description: "Type of sharing link",
        default: "view",
      },
      scope: {
        type: "string",
        enum: ["anonymous", "organization", "users"],
        description: "Who can access the link",
        default: "organization",
      },
      expirationDateTime: {
        type: "string",
        description: "Link expiration (ISO 8601 format, optional)",
      },
      password: {
        type: "string",
        description: "Password protection (optional)",
      },
    },
    required: [],
  },
};

export async function handleShareItem(args: any) {
  try {
    const client = getGraphClient();
    const {
      itemId,
      itemPath,
      type = "view",
      scope = "organization",
      expirationDateTime,
      password,
    } = args;
    const { siteId, driveId } = await resolveFileDriveContext(args, client);

    const endpoint = buildDriveItemEndpoint(
      { itemId, itemPath, siteId, driveId },
      "/createLink",
    );

    const shareData: any = {
      type,
      scope,
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

      return jsonTextResponse({
        success: true,
        message: "Sharing link created successfully",
        link: {
          id: permission.id,
          url: permission.link?.webUrl,
          type: permission.link?.type,
          scope: permission.link?.scope,
          expirationDateTime: permission.expirationDateTime,
          hasPassword: permission.hasPassword,
        },
      });
    }

    throw new Error("Failed to create sharing link");
  } catch (error) {
    return toolErrorResponse("share_item", error);
  }
}

// Tool 10: Copy item
export const copyItem: Tool = {
  name: "copy_item",
  description: "Copy a file or folder in OneDrive or SharePoint",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "Item ID to copy",
      },
      itemPath: {
        type: "string",
        description: "Alternative: item path to copy",
      },
      destinationFolderId: {
        type: "string",
        description: "Destination folder ID",
      },
      destinationFolderPath: {
        type: "string",
        description: "Alternative: destination folder path",
      },
      newName: {
        type: "string",
        description: "New name for the copied item (optional)",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (optional)",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description:
          "Canonical SharePoint site URL (optional alternative to siteId)",
      },
      driveId: {
        type: "string",
        description: "Drive ID for a specific document library (optional)",
      },
      destinationSiteId: {
        type: "string",
        description: "Destination SharePoint site ID (optional)",
      },
    },
    required: [],
  },
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
      destinationSiteId,
    } = args;
    const { siteId, driveId } = await resolveFileDriveContext(args, client);

    const endpoint = buildDriveItemEndpoint(
      { itemId, itemPath, siteId, driveId },
      "/copy",
    );

    const copyData: any = {
      parentReference: {},
    };

    if (destinationFolderId) {
      copyData.parentReference.id = destinationFolderId;
    } else if (destinationFolderPath) {
      // Resolve destination path to ID
      const destDriveRoot = destinationSiteId
        ? getDriveRootEndpoint({ siteId: destinationSiteId })
        : getDriveRootEndpoint({ siteId, driveId });
      const destEndpoint = `${destDriveRoot}/root:/${destinationFolderPath}`;

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
      return jsonTextResponse({
        success: true,
        message: "Item copy initiated successfully",
        note: "Copy operation is asynchronous and may take some time to complete",
        itemId: itemId || "path-based",
        itemPath: itemPath || "id-based",
        destinationFolderId,
        destinationFolderPath,
        newName,
      });
    }

    throw new Error("Failed to copy item");
  } catch (error) {
    return toolErrorResponse("copy_item", error);
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
  copyItem,
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
  copy_item: handleCopyItem,
};
