/**
 * Path utilities for OneDrive/SharePoint operations
 * Handles path sanitization, validation, and automatic folder creation
 */

import { getGraphClient } from "../../graph/client.js";
import { DriveItem } from "../../graph/models.js";
import { getDriveRootEndpoint } from "../../graph/resource-resolver.js";

export interface PathInfo {
  sanitizedPath: string;
  folderPath: string;
  fileName: string;
  needsFolderCreation: boolean;
}

/**
 * Sanitize file and folder names for OneDrive/SharePoint compatibility
 */
export function sanitizeFileName(name: string): string {
  // Remove or replace invalid characters for OneDrive/SharePoint
  const invalidChars = /[<>:"/\\|?*\x00-\x1f]/g;
  const reservedNames = [
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "COM1",
    "COM2",
    "COM3",
    "COM4",
    "COM5",
    "COM6",
    "COM7",
    "COM8",
    "COM9",
    "LPT1",
    "LPT2",
    "LPT3",
    "LPT4",
    "LPT5",
    "LPT6",
    "LPT7",
    "LPT8",
    "LPT9",
  ];

  let sanitized = name
    // Replace invalid characters with underscores
    .replace(invalidChars, "_")
    // Replace multiple spaces with single space
    .replace(/\s+/g, " ")
    // Remove leading/trailing spaces and dots
    .trim()
    .replace(/^\.+|\.+$/g, "")
    // Replace spaces with underscores for better compatibility
    .replace(/\s/g, "_");

  // Check for reserved names
  const nameWithoutExt = sanitized.split(".")[0].toUpperCase();
  if (reservedNames.includes(nameWithoutExt)) {
    sanitized = `file_${sanitized}`;
  }

  // Ensure it's not empty
  if (!sanitized) {
    sanitized = "untitled_file";
  }

  // Limit length (OneDrive has 400 char limit for full path)
  if (sanitized.length > 200) {
    const ext = sanitized.includes(".") ? "." + sanitized.split(".").pop() : "";
    sanitized = sanitized.substring(0, 200 - ext.length) + ext;
  }

  return sanitized;
}

/**
 * Sanitize and analyze a remote path
 */
export function analyzePath(remotePath: string): PathInfo {
  // Remove leading/trailing slashes and normalize
  const normalizedPath = remotePath
    .replace(/^\/+|\/+$/g, "")
    .replace(/\/+/g, "/");

  if (!normalizedPath) {
    return {
      sanitizedPath: "",
      folderPath: "",
      fileName: "untitled_file",
      needsFolderCreation: false,
    };
  }

  const pathParts = normalizedPath.split("/");
  const originalFileName = pathParts.pop() || "untitled_file";
  const folderParts = pathParts.map((part) => sanitizeFileName(part));
  const sanitizedFileName = sanitizeFileName(originalFileName);

  const folderPath = folderParts.join("/");
  const sanitizedPath = folderPath
    ? `${folderPath}/${sanitizedFileName}`
    : sanitizedFileName;

  return {
    sanitizedPath,
    folderPath,
    fileName: sanitizedFileName,
    needsFolderCreation: folderParts.length > 0,
  };
}

/**
 * Check if a folder exists, create it if it doesn't
 */
export async function ensureFolderExists(
  folderPath: string,
  siteId?: string,
  driveId?: string,
): Promise<{ success: boolean; folderId?: string; error?: string }> {
  if (!folderPath) {
    return { success: true }; // Root folder always exists
  }

  const client = getGraphClient();
  const driveRootEndpoint = getDriveRootEndpoint({ siteId, driveId });

  try {
    // Try to get the folder first
    const checkEndpoint = `${driveRootEndpoint}/root:/${folderPath}`;

    const checkResponse = await client.get<DriveItem>(checkEndpoint);

    if (checkResponse.success && checkResponse.data) {
      // Folder exists
      return {
        success: true,
        folderId: checkResponse.data.id,
      };
    }
  } catch (error) {
    // Folder doesn't exist, need to create it
  }

  // Create folder structure recursively
  const folderParts = folderPath.split("/");
  let currentPath = "";

  for (const part of folderParts) {
    const parentPath = currentPath;
    currentPath = currentPath ? `${currentPath}/${part}` : part;

    try {
      // Check if current folder exists
      const checkEndpoint = `${driveRootEndpoint}/root:/${currentPath}`;

      const checkResponse = await client.get<DriveItem>(checkEndpoint);
      if (checkResponse.success) {
        continue; // Folder exists, move to next
      }
    } catch (error) {
      // Folder doesn't exist, create it
    }

    // Create the folder
    const createEndpoint = parentPath
      ? `${driveRootEndpoint}/root:/${parentPath}:/children`
      : `${driveRootEndpoint}/root/children`;

    const folderData = {
      name: part,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    };

    const createResponse = await client.post<DriveItem>(
      createEndpoint,
      folderData,
    );
    if (!createResponse.success) {
      return {
        success: false,
        error: `Failed to create folder: ${part}`,
      };
    }
  }

  return { success: true };
}

/**
 * Prepare path for upload with automatic sanitization and folder creation
 */
export async function prepareUploadPath(
  remotePath: string,
  siteId?: string,
  driveId?: string,
): Promise<{
  success: boolean;
  sanitizedPath: string;
  originalPath: string;
  changes: string[];
  error?: string;
}> {
  const changes: string[] = [];
  const originalPath = remotePath;

  // Analyze and sanitize the path
  const pathInfo = analyzePath(remotePath);

  // Track changes
  if (
    pathInfo.sanitizedPath !==
    remotePath.replace(/^\/+|\/+$/g, "").replace(/\/+/g, "/")
  ) {
    changes.push(
      `Path sanitized: "${remotePath}" → "${pathInfo.sanitizedPath}"`,
    );
  }

  // Ensure folder exists if needed
  if (pathInfo.needsFolderCreation) {
    const folderResult = await ensureFolderExists(
      pathInfo.folderPath,
      siteId,
      driveId,
    );
    if (!folderResult.success) {
      return {
        success: false,
        sanitizedPath: pathInfo.sanitizedPath,
        originalPath,
        changes,
        error: folderResult.error,
      };
    }
    changes.push(`Folder verified/created: "${pathInfo.folderPath}"`);
  }

  return {
    success: true,
    sanitizedPath: pathInfo.sanitizedPath,
    originalPath,
    changes,
  };
}
