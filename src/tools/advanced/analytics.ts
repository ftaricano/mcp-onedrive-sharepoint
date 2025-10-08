/**
 * Storage analytics and version management tools
 * Advanced features for storage optimization and file versioning
 */

import { Tool } from '@modelcontextprotocol/sdk/types.js';
import { getGraphClient } from '../../graph/client.js';
import { DriveItem, Drive, GraphResponse } from '../../graph/models.js';
import { createUserFriendlyError } from '../../graph/error-handler.js';

// Tool 1: Storage analytics
export const storageAnalytics: Tool = {
  name: 'storage_analytics',
  description: 'Analyze storage usage patterns and identify optimization opportunities',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional, defaults to personal drive)'
      },
      analysisType: {
        type: 'string',
        enum: ['summary', 'detailed', 'duplicates', 'large_files', 'old_files', 'file_types'],
        description: 'Type of analysis to perform',
        default: 'summary'
      },
      path: {
        type: 'string',
        description: 'Specific path to analyze (optional, defaults to root)'
      },
      thresholds: {
        type: 'object',
        properties: {
          largeFileSize: {
            type: 'number',
            description: 'Size in MB to consider a file large',
            default: 100
          },
          oldFileDays: {
            type: 'number',
            description: 'Days since last modified to consider a file old',
            default: 365
          }
        }
      },
      includeVersions: {
        type: 'boolean',
        description: 'Include version history in size calculations',
        default: false
      }
    }
  }
};

export async function handleStorageAnalytics(args: any) {
  try {
    const client = getGraphClient();
    const { 
      siteId,
      analysisType = 'summary',
      path = '',
      thresholds = { largeFileSize: 100, oldFileDays: 365 },
      includeVersions = false
    } = args;

    const results: any = {
      analysisType,
      timestamp: new Date().toISOString(),
      path: path || 'root'
    };

    // Get drive information
    const driveEndpoint = siteId ? `/sites/${siteId}/drive` : '/me/drive';
    const driveResponse = await client.get<Drive>(driveEndpoint);

    if (!driveResponse.success || !driveResponse.data) {
      throw new Error('Failed to get drive information');
    }

    const drive = driveResponse.data;
    results.drive = {
      id: drive.id,
      name: drive.name,
      type: drive.driveType,
      quota: drive.quota ? {
        total: drive.quota.total,
        used: drive.quota.used,
        remaining: drive.quota.remaining,
        usedPercentage: drive.quota.total ? 
          Math.round((drive.quota.used / drive.quota.total) * 100) : 0
      } : null
    };

    // Get items based on path
    const itemsEndpoint = path
      ? (siteId 
          ? `/sites/${siteId}/drive/root:/${path}:/children`
          : `/me/drive/root:/${path}:/children`)
      : (siteId
          ? `/sites/${siteId}/drive/root/children`
          : '/me/drive/root/children');

    // Recursive function to analyze items
    const analyzeItems = async (endpoint: string, depth: number = 0): Promise<any> => {
      const analytics = {
        totalSize: 0,
        fileCount: 0,
        folderCount: 0,
        largeFiles: [] as any[],
        oldFiles: [] as any[],
        fileTypes: {} as Record<string, { count: number; size: number }>,
        duplicates: new Map<string, any[]>(),
        versionedFiles: [] as any[]
      };

      const response = await client.get<GraphResponse<DriveItem>>(endpoint, {
        '$select': 'id,name,size,file,folder,lastModifiedDateTime,createdDateTime,webUrl',
        '$top': '1000'
      });

      if (!response.success || !response.data) {
        return analytics;
      }

      const items = (response.data as any).value || [];
      const now = Date.now();
      const oldThreshold = now - (thresholds.oldFileDays * 24 * 60 * 60 * 1000);
      const largeSizeThreshold = thresholds.largeFileSize * 1024 * 1024;

      for (const item of items) {
        if (item.file) {
          // File analytics
          analytics.fileCount++;
          analytics.totalSize += item.size || 0;

          // File type analysis
          const ext = item.name.split('.').pop()?.toLowerCase() || 'no_extension';
          if (!analytics.fileTypes[ext]) {
            analytics.fileTypes[ext] = { count: 0, size: 0 };
          }
          analytics.fileTypes[ext].count++;
          analytics.fileTypes[ext].size += item.size || 0;

          // Large files
          if (item.size && item.size > largeSizeThreshold) {
            analytics.largeFiles.push({
              id: item.id,
              name: item.name,
              size: item.size,
              sizeMB: Math.round(item.size / (1024 * 1024)),
              lastModified: item.lastModifiedDateTime,
              webUrl: item.webUrl
            });
          }

          // Old files
          const lastModified = new Date(item.lastModifiedDateTime).getTime();
          if (lastModified < oldThreshold) {
            analytics.oldFiles.push({
              id: item.id,
              name: item.name,
              size: item.size,
              lastModified: item.lastModifiedDateTime,
              daysSinceModified: Math.floor((now - lastModified) / (24 * 60 * 60 * 1000)),
              webUrl: item.webUrl
            });
          }

          // Check for duplicates (by name and size)
          const duplicateKey = `${item.name}_${item.size}`;
          if (!analytics.duplicates.has(duplicateKey)) {
            analytics.duplicates.set(duplicateKey, []);
          }
          analytics.duplicates.get(duplicateKey)!.push({
            id: item.id,
            name: item.name,
            size: item.size,
            path: item.parentReference?.path,
            lastModified: item.lastModifiedDateTime
          });

          // Version analysis if requested
          if (includeVersions) {
            try {
              const versionsEndpoint = siteId
                ? `/sites/${siteId}/drive/items/${item.id}/versions`
                : `/me/drive/items/${item.id}/versions`;
              
              const versionsResponse = await client.get<any>(versionsEndpoint);
              if (versionsResponse.success && versionsResponse.data) {
                const versions = (versionsResponse.data as any).value || [];
                if (versions.length > 1) {
                  const versionSizes = versions.reduce((sum: number, v: any) => 
                    sum + (v.size || 0), 0
                  );
                  analytics.versionedFiles.push({
                    id: item.id,
                    name: item.name,
                    currentSize: item.size,
                    versionCount: versions.length,
                    totalVersionSize: versionSizes,
                    oldestVersion: versions[versions.length - 1]?.lastModifiedDateTime
                  });
                }
              }
            } catch (versionError) {
              // Skip version analysis for this file
            }
          }
        } else if (item.folder) {
          // Folder analytics
          analytics.folderCount++;

          // Recursively analyze subfolders if doing detailed analysis
          if (analysisType === 'detailed' && depth < 3) {
            const subfolderEndpoint = siteId
              ? `/sites/${siteId}/drive/items/${item.id}/children`
              : `/me/drive/items/${item.id}/children`;
            
            const subAnalytics = await analyzeItems(subfolderEndpoint, depth + 1);
            
            // Merge results
            analytics.totalSize += subAnalytics.totalSize;
            analytics.fileCount += subAnalytics.fileCount;
            analytics.folderCount += subAnalytics.folderCount;
            analytics.largeFiles.push(...subAnalytics.largeFiles);
            analytics.oldFiles.push(...subAnalytics.oldFiles);
            
            // Merge file types
            for (const [ext, data] of Object.entries(subAnalytics.fileTypes)) {
              if (!analytics.fileTypes[ext]) {
                analytics.fileTypes[ext] = { count: 0, size: 0 };
              }
              analytics.fileTypes[ext].count += (data as any).count;
              analytics.fileTypes[ext].size += (data as any).size;
            }
            
            // Merge duplicates
            for (const [key, files] of subAnalytics.duplicates) {
              if (!analytics.duplicates.has(key)) {
                analytics.duplicates.set(key, []);
              }
              analytics.duplicates.get(key)!.push(...files);
            }
            
            analytics.versionedFiles.push(...subAnalytics.versionedFiles);
          }
        }
      }

      return analytics;
    };

    const analytics = await analyzeItems(itemsEndpoint);

    // Process results based on analysis type
    switch (analysisType) {
      case 'summary':
        results.summary = {
          totalFiles: analytics.fileCount,
          totalFolders: analytics.folderCount,
          totalSize: analytics.totalSize,
          totalSizeMB: Math.round(analytics.totalSize / (1024 * 1024)),
          largeFilesCount: analytics.largeFiles.length,
          oldFilesCount: analytics.oldFiles.length,
          topFileTypes: Object.entries(analytics.fileTypes)
            .sort((a: [string, any], b: [string, any]) => b[1].size - a[1].size)
            .slice(0, 10)
            .map(([ext, data]: [string, any]) => ({
              extension: ext,
              count: data.count,
              totalSizeMB: Math.round(data.size / (1024 * 1024))
            }))
        };
        break;

      case 'detailed':
        results.detailed = {
          ...analytics,
          duplicates: Array.from<[string, any[]]>(analytics.duplicates.entries())
            .filter(([_, files]: [string, any[]]) => files.length > 1)
            .map(([key, files]: [string, any[]]) => ({
              key,
              count: files.length,
              totalSize: files[0].size * files.length,
              files
            }))
        };
        break;

      case 'duplicates':
        const duplicates = Array.from<[string, any[]]>(analytics.duplicates.entries())
          .filter(([_, files]: [string, any[]]) => files.length > 1)
          .map(([key, files]: [string, any[]]) => ({
            name: files[0].name,
            duplicateCount: files.length,
            sizeEach: files[0].size,
            totalWastedSpace: files[0].size * (files.length - 1),
            locations: files.map((f: any) => ({
              id: f.id,
              path: f.path,
              lastModified: f.lastModified
            }))
          }))
          .sort((a: any, b: any) => b.totalWastedSpace - a.totalWastedSpace);

        results.duplicates = {
          totalDuplicateSets: duplicates.length,
          totalWastedSpace: duplicates.reduce((sum: number, d: any) => sum + d.totalWastedSpace, 0),
          duplicates: duplicates.slice(0, 50) // Top 50 duplicates
        };
        break;

      case 'large_files':
        results.largeFiles = {
          threshold: thresholds.largeFileSize,
          count: analytics.largeFiles.length,
          totalSize: analytics.largeFiles.reduce((sum: number, f: any) => sum + f.size, 0),
          files: analytics.largeFiles
            .sort((a: any, b: any) => b.size - a.size)
            .slice(0, 100) // Top 100 large files
        };
        break;

      case 'old_files':
        results.oldFiles = {
          thresholdDays: thresholds.oldFileDays,
          count: analytics.oldFiles.length,
          totalSize: analytics.oldFiles.reduce((sum: number, f: any) => sum + (f.size || 0), 0),
          files: analytics.oldFiles
            .sort((a: any, b: any) => b.daysSinceModified - a.daysSinceModified)
            .slice(0, 100) // Top 100 old files
        };
        break;

      case 'file_types':
        const sortedTypes = Object.entries(analytics.fileTypes)
          .map(([ext, data]: [string, any]) => ({
            extension: ext,
            count: data.count,
            totalSize: data.size,
            totalSizeMB: Math.round(data.size / (1024 * 1024)),
            averageSizeMB: Math.round((data.size / data.count) / (1024 * 1024))
          }))
          .sort((a: any, b: any) => b.totalSize - a.totalSize);

        results.fileTypes = {
          totalTypes: sortedTypes.length,
          topBySize: sortedTypes.slice(0, 20),
          topByCount: [...sortedTypes].sort((a: any, b: any) => b.count - a.count).slice(0, 20)
        };
        break;
    }

    // Add version analysis if included
    if (includeVersions && analytics.versionedFiles.length > 0) {
      results.versions = {
        filesWithVersions: analytics.versionedFiles.length,
        totalVersionSpace: analytics.versionedFiles.reduce((sum: number, f: any) => 
          sum + f.totalVersionSize, 0
        ),
        topVersionedFiles: analytics.versionedFiles
          .sort((a: any, b: any) => b.totalVersionSize - a.totalVersionSize)
          .slice(0, 20)
      };
    }

    return {
      content: [{
        type: 'text',
        text: JSON.stringify(results, null, 2)
      }]
    };
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error analyzing storage: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 2: Version management
export const versionManagement: Tool = {
  name: 'version_management',
  description: 'Manage file versions including restore, cleanup, and comparison',
  inputSchema: {
    type: 'object',
    properties: {
      itemId: {
        type: 'string',
        description: 'File item ID'
      },
      itemPath: {
        type: 'string',
        description: 'Alternative: file path'
      },
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)'
      },
      action: {
        type: 'string',
        enum: ['list', 'restore', 'delete', 'cleanup', 'compare'],
        description: 'Version management action',
        default: 'list'
      },
      versionId: {
        type: 'string',
        description: 'Version ID for restore/delete/compare actions'
      },
      keepVersions: {
        type: 'number',
        description: 'Number of versions to keep (for cleanup)',
        default: 5
      },
      compareVersionId: {
        type: 'string',
        description: 'Second version ID for comparison'
      }
    },
    required: ['action']
  }
};

export async function handleVersionManagement(args: any) {
  try {
    const client = getGraphClient();
    const { 
      itemId,
      itemPath,
      siteId,
      action = 'list',
      versionId,
      keepVersions = 5,
      compareVersionId
    } = args;

    // Get item ID if path is provided
    let actualItemId = itemId;
    if (!actualItemId && itemPath) {
      const itemEndpoint = siteId
        ? `/sites/${siteId}/drive/root:/${itemPath}`
        : `/me/drive/root:/${itemPath}`;
      
      const itemResponse = await client.get<DriveItem>(itemEndpoint);
      if (itemResponse.success && itemResponse.data) {
        actualItemId = itemResponse.data.id;
      } else {
        throw new Error('Failed to resolve item path');
      }
    }

    if (!actualItemId) {
      throw new Error('Either itemId or itemPath is required');
    }

    const baseEndpoint = siteId
      ? `/sites/${siteId}/drive/items/${actualItemId}`
      : `/me/drive/items/${actualItemId}`;

    switch (action) {
      case 'list': {
        const versionsEndpoint = `${baseEndpoint}/versions`;
        const response = await client.get<any>(versionsEndpoint);

        if (response.success && response.data) {
          const versions = (response.data as any).value || [];
          
          // Get current item info
          const currentResponse = await client.get<DriveItem>(baseEndpoint);
          const currentItem = currentResponse.data;

          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                action: 'list',
                file: {
                  id: actualItemId,
                  name: currentItem?.name,
                  currentSize: currentItem?.size,
                  webUrl: currentItem?.webUrl
                },
                versionCount: versions.length,
                totalVersionSize: versions.reduce((sum: number, v: any) => 
                  sum + (v.size || 0), 0
                ),
                versions: versions.map((v: any) => ({
                  id: v.id,
                  version: v.version,
                  size: v.size,
                  sizeMB: v.size ? Math.round(v.size / (1024 * 1024)) : 0,
                  lastModifiedDateTime: v.lastModifiedDateTime,
                  lastModifiedBy: v.lastModifiedBy?.user?.displayName,
                  isCurrent: v.id === 'current'
                }))
              }, null, 2)
            }]
          };
        }
        break;
      }

      case 'restore': {
        if (!versionId) {
          throw new Error('versionId is required for restore action');
        }

        const restoreEndpoint = `${baseEndpoint}/versions/${versionId}/restoreVersion`;
        const response = await client.post(restoreEndpoint, {});

        if (response.success) {
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                action: 'restore',
                success: true,
                message: `Version ${versionId} restored successfully`,
                itemId: actualItemId,
                restoredVersion: versionId
              }, null, 2)
            }]
          };
        }
        break;
      }

      case 'delete': {
        if (!versionId) {
          throw new Error('versionId is required for delete action');
        }

        const deleteEndpoint = `${baseEndpoint}/versions/${versionId}`;
        const response = await client.delete(deleteEndpoint);

        if (response.success) {
          return {
            content: [{
              type: 'text',
              text: JSON.stringify({
                action: 'delete',
                success: true,
                message: `Version ${versionId} deleted successfully`,
                itemId: actualItemId,
                deletedVersion: versionId
              }, null, 2)
            }]
          };
        }
        break;
      }

      case 'cleanup': {
        const versionsEndpoint = `${baseEndpoint}/versions`;
        const listResponse = await client.get<any>(versionsEndpoint);

        if (!listResponse.success || !listResponse.data) {
          throw new Error('Failed to get versions');
        }

        const versions = (listResponse.data as any).value || [];
        
        // Sort versions by date (newest first)
        versions.sort((a: any, b: any) => 
          new Date(b.lastModifiedDateTime).getTime() - 
          new Date(a.lastModifiedDateTime).getTime()
        );

        // Keep only the specified number of versions
        const versionsToDelete = versions.slice(keepVersions);
        const deletedVersions: any[] = [];
        const failedDeletions: any[] = [];

        for (const version of versionsToDelete) {
          if (version.id === 'current') continue; // Skip current version

          try {
            const deleteEndpoint = `${baseEndpoint}/versions/${version.id}`;
            await client.delete(deleteEndpoint);
            deletedVersions.push({
              id: version.id,
              version: version.version,
              size: version.size,
              lastModified: version.lastModifiedDateTime
            });
          } catch (error) {
            failedDeletions.push({
              id: version.id,
              error: createUserFriendlyError(error)
            });
          }
        }

        const spaceSaved = deletedVersions.reduce((sum, v) => sum + (v.size || 0), 0);

        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              action: 'cleanup',
              success: failedDeletions.length === 0,
              keepVersions,
              originalVersionCount: versions.length,
              deletedCount: deletedVersions.length,
              remainingCount: versions.length - deletedVersions.length,
              spaceSaved,
              spaceSavedMB: Math.round(spaceSaved / (1024 * 1024)),
              deletedVersions,
              failedDeletions
            }, null, 2)
          }]
        };
      }

      case 'compare': {
        if (!versionId || !compareVersionId) {
          throw new Error('versionId and compareVersionId are required for compare action');
        }

        // Get both versions
        const [version1Response, version2Response] = await Promise.all([
          client.get<any>(`${baseEndpoint}/versions/${versionId}`),
          client.get<any>(`${baseEndpoint}/versions/${compareVersionId}`)
        ]);

        if (!version1Response.success || !version2Response.success) {
          throw new Error('Failed to get versions for comparison');
        }

        const version1 = version1Response.data;
        const version2 = version2Response.data;

        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              action: 'compare',
              itemId: actualItemId,
              comparison: {
                version1: {
                  id: version1.id,
                  version: version1.version,
                  size: version1.size,
                  lastModified: version1.lastModifiedDateTime,
                  lastModifiedBy: version1.lastModifiedBy?.user?.displayName
                },
                version2: {
                  id: version2.id,
                  version: version2.version,
                  size: version2.size,
                  lastModified: version2.lastModifiedDateTime,
                  lastModifiedBy: version2.lastModifiedBy?.user?.displayName
                },
                differences: {
                  sizeDifference: (version2.size || 0) - (version1.size || 0),
                  timeDifference: new Date(version2.lastModifiedDateTime).getTime() - 
                                 new Date(version1.lastModifiedDateTime).getTime(),
                  daysBetween: Math.floor(
                    (new Date(version2.lastModifiedDateTime).getTime() - 
                     new Date(version1.lastModifiedDateTime).getTime()) / 
                    (24 * 60 * 60 * 1000)
                  )
                }
              }
            }, null, 2)
          }]
        };
      }

      default:
        throw new Error(`Invalid action: ${action}`);
    }

    throw new Error(`Failed to ${action} versions`);
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error managing versions: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Export all analytics tools and handlers
export const analyticsTools = [
  storageAnalytics,
  versionManagement
];

export const analyticsHandlers = {
  storage_analytics: handleStorageAnalytics,
  version_management: handleVersionManagement
};