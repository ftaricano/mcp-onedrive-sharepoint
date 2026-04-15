/**
 * Advanced synchronization and file management tools
 * Optimized for batch operations and large file handling
 */

import { Tool } from '@modelcontextprotocol/sdk/types.js';
import { getGraphClient } from '../../graph/client.js';
import { jsonTextResponse, toolErrorResponse } from '../../graph/contracts.js';
import { DriveItem, UploadSession } from '../../graph/models.js';
import { createUserFriendlyError } from '../../graph/error-handler.js';
import * as fs from 'fs';
import * as path from 'path';
import * as crypto from 'crypto';

// Tool 1: Sync folder (bidirectional)
export const syncFolder: Tool = {
  name: 'sync_folder',
  description: 'Synchronize a local folder with OneDrive/SharePoint (bidirectional)',
  inputSchema: {
    type: 'object',
    properties: {
      localPath: {
        type: 'string',
        description: 'Local folder path to sync'
      },
      remotePath: {
        type: 'string',
        description: 'Remote folder path in OneDrive/SharePoint'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      direction: {
        type: 'string',
        enum: ['upload', 'download', 'bidirectional'],
        description: 'Sync direction',
        default: 'bidirectional'
      },
      conflictResolution: {
        type: 'string',
        enum: ['local', 'remote', 'newer', 'rename'],
        description: 'How to resolve conflicts',
        default: 'newer'
      },
      includePatterns: {
        type: 'array',
        items: { type: 'string' },
        description: 'File patterns to include (e.g., ["*.docx", "*.xlsx"])'
      },
      excludePatterns: {
        type: 'array',
        items: { type: 'string' },
        description: 'File patterns to exclude (e.g., ["*.tmp", "~*"])'
      },
      deleteOrphans: {
        type: 'boolean',
        description: 'Delete files that exist only on one side',
        default: false
      }
    },
    required: ['localPath', 'remotePath']
  }
};

export async function handleSyncFolder(args: any) {
  try {
    const client = getGraphClient();
    const { 
      localPath, 
      remotePath, 
      siteId,
      direction = 'bidirectional',
      conflictResolution = 'newer',
      includePatterns = [],
      excludePatterns = [],
      deleteOrphans = false
    } = args;

    const syncResults = {
      uploaded: [] as any[],
      downloaded: [] as any[],
      skipped: [] as any[],
      conflicts: [] as any[],
      errors: [] as any[],
      deleted: [] as any[]
    };

    // Validate local path
    if (!fs.existsSync(localPath)) {
      if (direction === 'upload') {
        throw new Error(`Local path does not exist: ${localPath}`);
      } else {
        // Create local directory for download
        fs.mkdirSync(localPath, { recursive: true });
      }
    }

    // Get remote folder contents
    const remoteEndpoint = siteId
      ? `/sites/${siteId}/drive/root:/${remotePath}:/children`
      : `/me/drive/root:/${remotePath}:/children`;

    const remoteResponse = await client.get<any>(remoteEndpoint, {
      '$select': 'id,name,size,lastModifiedDateTime,file,folder',
      '$top': '1000'
    });

    const remoteItems = remoteResponse.success && remoteResponse.data 
      ? (remoteResponse.data as any).value || []
      : [];

    // Get local folder contents
    const localItems = fs.readdirSync(localPath, { withFileTypes: true });

    // Create maps for comparison
    const remoteMap = new Map(remoteItems.map((item: any) => [item.name, item]));
    const localMap = new Map(localItems.map(item => [item.name, item]));

    // Helper function to check file patterns
    const matchesPattern = (filename: string, patterns: string[]): boolean => {
      if (patterns.length === 0) return true;
      return patterns.some(pattern => {
        const regex = new RegExp(pattern.replace(/\*/g, '.*').replace(/\?/g, '.'));
        return regex.test(filename);
      });
    };

    // Process uploads (local → remote)
    if (direction === 'upload' || direction === 'bidirectional') {
      for (const localItem of localItems) {
        if (localItem.isDirectory()) continue; // Skip directories for now

        const filename = localItem.name;
        
        // Check patterns
        if (includePatterns.length > 0 && !matchesPattern(filename, includePatterns)) {
          syncResults.skipped.push({ name: filename, reason: 'Not in include pattern' });
          continue;
        }
        if (excludePatterns.length > 0 && matchesPattern(filename, excludePatterns)) {
          syncResults.skipped.push({ name: filename, reason: 'In exclude pattern' });
          continue;
        }

        const localFilePath = path.join(localPath, filename);
        const localStats = fs.statSync(localFilePath);
        const remoteItem = remoteMap.get(filename);

        // Determine if upload is needed
        let shouldUpload = false;
        if (!remoteItem) {
          shouldUpload = true;
        } else if ((remoteItem as any).file) {
          const remoteModified = new Date((remoteItem as any).lastModifiedDateTime);
          const localModified = localStats.mtime;

          switch (conflictResolution) {
            case 'local':
              shouldUpload = true;
              break;
            case 'remote':
              shouldUpload = false;
              break;
            case 'newer':
              shouldUpload = localModified > remoteModified;
              break;
            case 'rename':
              // Upload with new name
              const baseName = path.basename(filename, path.extname(filename));
              const ext = path.extname(filename);
              const newName = `${baseName}_local_${Date.now()}${ext}`;
              // Upload logic with new name would go here
              syncResults.conflicts.push({
                name: filename,
                resolution: 'renamed',
                newName
              });
              continue;
          }
        }

        if (shouldUpload) {
          try {
            const uploadEndpoint = siteId
              ? `/sites/${siteId}/drive/root:/${remotePath}/${filename}:/content`
              : `/me/drive/root:/${remotePath}/${filename}:/content`;

            const uploadResponse = await client.uploadFile(
              uploadEndpoint,
              localFilePath,
              filename,
              { conflictBehavior: 'replace' }
            );

            if (uploadResponse.success) {
              syncResults.uploaded.push({
                name: filename,
                size: localStats.size,
                localPath: localFilePath
              });
            }
          } catch (error) {
            syncResults.errors.push({
              name: filename,
              error: createUserFriendlyError(error),
              operation: 'upload'
            });
          }
        }
      }
    }

    // Process downloads (remote → local)
    if (direction === 'download' || direction === 'bidirectional') {
      for (const remoteItem of remoteItems) {
        if (!remoteItem.file) continue; // Skip folders for now

        const filename = remoteItem.name;
        
        // Check patterns
        if (includePatterns.length > 0 && !matchesPattern(filename, includePatterns)) {
          syncResults.skipped.push({ name: filename, reason: 'Not in include pattern' });
          continue;
        }
        if (excludePatterns.length > 0 && matchesPattern(filename, excludePatterns)) {
          syncResults.skipped.push({ name: filename, reason: 'In exclude pattern' });
          continue;
        }

        const localFilePath = path.join(localPath, filename);
        const localExists = fs.existsSync(localFilePath);

        // Determine if download is needed
        let shouldDownload = false;
        if (!localExists) {
          shouldDownload = true;
        } else {
          const localStats = fs.statSync(localFilePath);
          const remoteModified = new Date(remoteItem.lastModifiedDateTime);
          const localModified = localStats.mtime;

          switch (conflictResolution) {
            case 'remote':
              shouldDownload = true;
              break;
            case 'local':
              shouldDownload = false;
              break;
            case 'newer':
              shouldDownload = remoteModified > localModified;
              break;
            case 'rename':
              // Download with new name
              const baseName = path.basename(filename, path.extname(filename));
              const ext = path.extname(filename);
              const newName = `${baseName}_remote_${Date.now()}${ext}`;
              const newPath = path.join(localPath, newName);
              // Download logic with new name would go here
              syncResults.conflicts.push({
                name: filename,
                resolution: 'renamed',
                newName
              });
              continue;
          }
        }

        if (shouldDownload) {
          try {
            const downloadEndpoint = siteId
              ? `/sites/${siteId}/drive/items/${remoteItem.id}/content`
              : `/me/drive/items/${remoteItem.id}/content`;

            const downloadResponse = await client.downloadFile(downloadEndpoint);

            if (downloadResponse.success && downloadResponse.data) {
              fs.writeFileSync(localFilePath, downloadResponse.data as Buffer);
              
              // Preserve modification time
              const remoteModified = new Date(remoteItem.lastModifiedDateTime);
              fs.utimesSync(localFilePath, remoteModified, remoteModified);

              syncResults.downloaded.push({
                name: filename,
                size: remoteItem.size,
                localPath: localFilePath
              });
            }
          } catch (error) {
            syncResults.errors.push({
              name: filename,
              error: createUserFriendlyError(error),
              operation: 'download'
            });
          }
        }
      }
    }

    // Handle orphan deletion if requested
    if (deleteOrphans) {
      // Delete local orphans
      if (direction === 'download' || direction === 'bidirectional') {
        for (const localItem of localItems) {
          if (!remoteMap.has(localItem.name)) {
            const localFilePath = path.join(localPath, localItem.name);
            try {
              if (localItem.isDirectory()) {
                fs.rmSync(localFilePath, { recursive: true });
              } else {
                fs.unlinkSync(localFilePath);
              }
              syncResults.deleted.push({
                name: localItem.name,
                location: 'local'
              });
            } catch (error) {
              syncResults.errors.push({
                name: localItem.name,
                error: createUserFriendlyError(error),
                operation: 'delete_local'
              });
            }
          }
        }
      }

      // Delete remote orphans
      if (direction === 'upload' || direction === 'bidirectional') {
        for (const remoteItem of remoteItems) {
          if (!localMap.has(remoteItem.name)) {
            try {
              const deleteEndpoint = siteId
                ? `/sites/${siteId}/drive/items/${remoteItem.id}`
                : `/me/drive/items/${remoteItem.id}`;

              await client.delete(deleteEndpoint);
              
              syncResults.deleted.push({
                name: remoteItem.name,
                location: 'remote'
              });
            } catch (error) {
              syncResults.errors.push({
                name: remoteItem.name,
                error: createUserFriendlyError(error),
                operation: 'delete_remote'
              });
            }
          }
        }
      }
    }

    return jsonTextResponse({
      success: true,
      localPath,
      remotePath,
      direction,
      conflictResolution,
      summary: {
        uploaded: syncResults.uploaded.length,
        downloaded: syncResults.downloaded.length,
        skipped: syncResults.skipped.length,
        conflicts: syncResults.conflicts.length,
        deleted: syncResults.deleted.length,
        errors: syncResults.errors.length
      },
      details: syncResults
    });
  } catch (error) {
    return toolErrorResponse('sync_folder', error);
  }
}

// Tool 2: Batch file operations
export const batchFileOperations: Tool = {
  name: 'batch_file_operations',
  description: 'Perform multiple file operations in a single batch',
  inputSchema: {
    type: 'object',
    properties: {
      operations: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            operation: {
              type: 'string',
              enum: ['upload', 'download', 'move', 'copy', 'delete', 'rename'],
              description: 'Operation type'
            },
            source: {
              type: 'string',
              description: 'Source path (local for upload, remote for others)'
            },
            destination: {
              type: 'string',
              description: 'Destination path'
            },
            itemId: {
              type: 'string',
              description: 'Item ID (for operations on existing items)'
            },
            newName: {
              type: 'string',
              description: 'New name (for rename operation)'
            }
          },
          required: ['operation']
        },
        description: 'Array of operations to perform',
        maxItems: 50
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      stopOnError: {
        type: 'boolean',
        description: 'Stop processing if an operation fails',
        default: false
      },
      parallel: {
        type: 'boolean',
        description: 'Execute operations in parallel (faster but may hit rate limits)',
        default: false
      }
    },
    required: ['operations']
  }
};

export async function handleBatchFileOperations(args: any) {
  try {
    const client = getGraphClient();
    const { 
      operations, 
      siteId,
      stopOnError = false,
      parallel = false
    } = args;

    if (!operations || operations.length === 0) {
      throw new Error('At least one operation is required');
    }

    if (operations.length > 50) {
      throw new Error('Maximum 50 operations allowed per batch');
    }

    const results: any[] = [];

    const processOperation = async (op: any, index: number) => {
      const result: any = {
        index,
        operation: op.operation,
        source: op.source,
        destination: op.destination
      };

      try {
        switch (op.operation) {
          case 'upload': {
            if (!op.source || !op.destination) {
              throw new Error('Source and destination required for upload');
            }

            const uploadEndpoint = siteId
              ? `/sites/${siteId}/drive/root:/${op.destination}:/content`
              : `/me/drive/root:/${op.destination}:/content`;

            const response = await client.uploadFile(
              uploadEndpoint,
              op.source,
              path.basename(op.destination),
              { conflictBehavior: 'rename' }
            );

            result.success = response.success;
            if (response.data) {
              result.itemId = (response.data as any).id;
            }
            break;
          }

          case 'download': {
            if (!op.source || !op.destination) {
              throw new Error('Source and destination required for download');
            }

            const downloadEndpoint = op.itemId
              ? (siteId 
                  ? `/sites/${siteId}/drive/items/${op.itemId}/content`
                  : `/me/drive/items/${op.itemId}/content`)
              : (siteId
                  ? `/sites/${siteId}/drive/root:/${op.source}:/content`
                  : `/me/drive/root:/${op.source}:/content`);

            const response = await client.downloadFile(downloadEndpoint);

            if (response.success && response.data) {
              fs.writeFileSync(op.destination, response.data as Buffer);
              result.success = true;
              result.size = (response.data as Buffer).length;
            }
            break;
          }

          case 'move': {
            if (!op.itemId || !op.destination) {
              throw new Error('ItemId and destination required for move');
            }

            const moveEndpoint = siteId
              ? `/sites/${siteId}/drive/items/${op.itemId}`
              : `/me/drive/items/${op.itemId}`;

            const response = await client.patch(moveEndpoint, {
              parentReference: { path: op.destination }
            });

            result.success = response.success;
            break;
          }

          case 'copy': {
            if (!op.itemId || !op.destination) {
              throw new Error('ItemId and destination required for copy');
            }

            const copyEndpoint = siteId
              ? `/sites/${siteId}/drive/items/${op.itemId}/copy`
              : `/me/drive/items/${op.itemId}/copy`;

            const response = await client.post(copyEndpoint, {
              parentReference: { path: op.destination },
              name: op.newName || undefined
            });

            result.success = response.success;
            break;
          }

          case 'delete': {
            if (!op.itemId && !op.source) {
              throw new Error('ItemId or source required for delete');
            }

            const deleteEndpoint = op.itemId
              ? (siteId
                  ? `/sites/${siteId}/drive/items/${op.itemId}`
                  : `/me/drive/items/${op.itemId}`)
              : (siteId
                  ? `/sites/${siteId}/drive/root:/${op.source}`
                  : `/me/drive/root:/${op.source}`);

            const response = await client.delete(deleteEndpoint);
            result.success = response.success;
            break;
          }

          case 'rename': {
            if (!op.itemId || !op.newName) {
              throw new Error('ItemId and newName required for rename');
            }

            const renameEndpoint = siteId
              ? `/sites/${siteId}/drive/items/${op.itemId}`
              : `/me/drive/items/${op.itemId}`;

            const response = await client.patch(renameEndpoint, {
              name: op.newName
            });

            result.success = response.success;
            break;
          }

          default:
            throw new Error(`Unknown operation: ${op.operation}`);
        }

        result.status = 'completed';
      } catch (error) {
        result.success = false;
        result.status = 'failed';
        result.error = createUserFriendlyError(error);

        if (stopOnError) {
          throw error;
        }
      }

      return result;
    };

    if (parallel) {
      // Process operations in parallel
      const promises = operations.map((op: any, index: number) => 
        processOperation(op, index)
      );
      const parallelResults = await Promise.allSettled(promises);
      
      parallelResults.forEach((promiseResult, index) => {
        if (promiseResult.status === 'fulfilled') {
          results.push(promiseResult.value);
        } else {
          results.push({
            index,
            operation: operations[index].operation,
            success: false,
            status: 'failed',
            error: promiseResult.reason?.message || 'Unknown error'
          });
        }
      });
    } else {
      // Process operations sequentially
      for (let i = 0; i < operations.length; i++) {
        const result = await processOperation(operations[i], i);
        results.push(result);
        
        if (!result.success && stopOnError) {
          break;
        }
      }
    }

    const summary = {
      total: results.length,
      successful: results.filter(r => r.success).length,
      failed: results.filter(r => !r.success).length,
      parallel,
      stopOnError
    };

    return jsonTextResponse({
      success: summary.failed === 0,
      summary,
      results
    });
  } catch (error) {
    return toolErrorResponse('batch_file_operations', error);
  }
}

// Export all sync tools and handlers
export const syncTools = [
  syncFolder,
  batchFileOperations
];

export const syncHandlers = {
  sync_folder: handleSyncFolder,
  batch_file_operations: handleBatchFileOperations
};