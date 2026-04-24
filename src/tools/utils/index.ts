/**
 * Utility tools for OneDrive/SharePoint MCP Server
 * General utilities for authentication, health checks, and system management
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../../graph/client.js";
import { User, Drive, Site, GraphResponse } from "../../graph/models.js";
import { jsonTextResponse, toolErrorResponse } from "../../graph/contracts.js";
import { escapeODataString } from "../../graph/resource-resolver.js";
import { resolveRequiredSharePointSite } from "../../sharepoint/site-resolver.js";

type UtilityDependencies = {
  getGraphClient: typeof getGraphClient;
  getAuthInstance: () => Promise<{ isAuthenticated(): Promise<boolean> }>;
};

let utilityDependencies: UtilityDependencies = {
  getGraphClient,
  getAuthInstance: async () =>
    (await import("../../auth/microsoft-graph-auth.js")).getAuthInstance(),
};

export function __setUtilityDependenciesForTests(
  overrides?: Partial<UtilityDependencies>,
): void {
  utilityDependencies = {
    getGraphClient,
    getAuthInstance: async () =>
      (await import("../../auth/microsoft-graph-auth.js")).getAuthInstance(),
    ...overrides,
  };
}

// Tool 1: Health check and authentication status
export const healthCheck: Tool = {
  name: "health_check",
  description: "Check the health status and authentication of the MCP server",
  inputSchema: {
    type: "object",
    properties: {
      includeUserInfo: {
        type: "boolean",
        description: "Include user profile information",
        default: true,
      },
      includeDriveInfo: {
        type: "boolean",
        description: "Include default drive information",
        default: true,
      },
    },
  },
};

export async function handleHealthCheck(args: any) {
  try {
    const { includeUserInfo = true, includeDriveInfo = true } = args;
    const client = utilityDependencies.getGraphClient();
    const auth = await utilityDependencies.getAuthInstance();

    // Check authentication status
    const isAuthenticated = await auth.isAuthenticated();

    const healthStatus: any = {
      server: "MCP OneDrive/SharePoint Server",
      version: "1.0.0",
      status: "healthy",
      timestamp: new Date().toISOString(),
      authentication: {
        isAuthenticated,
        authMethod: "Microsoft Graph Device Code Flow",
      },
    };

    if (!isAuthenticated) {
      healthStatus.status = "authentication_required";
      healthStatus.message = "Please authenticate using the setup-auth script";

      return jsonTextResponse(healthStatus);
    }

    // Test API connectivity
    const apiTest = await client.healthCheck();
    healthStatus.apiConnectivity = {
      status: apiTest.success ? "connected" : "failed",
      graphApiVersion: "v1.0",
      endpoint: "https://graph.microsoft.com/v1.0",
    };

    if (includeUserInfo && apiTest.success && apiTest.data) {
      const userData = apiTest.data;
      healthStatus.user = {
        id: userData.user?.id,
        displayName: userData.user?.displayName,
        mail: userData.user?.mail,
        userPrincipalName: userData.user?.userPrincipalName,
      };
    }

    if (includeDriveInfo && apiTest.success) {
      try {
        const driveResponse = await client.get<Drive>("/me/drive");
        if (driveResponse.success && driveResponse.data) {
          const drive = driveResponse.data;
          healthStatus.defaultDrive = {
            id: drive.id,
            name: drive.name,
            driveType: drive.driveType,
            quota: drive.quota
              ? {
                  total: drive.quota.total,
                  used: drive.quota.used,
                  remaining: drive.quota.remaining,
                  state: drive.quota.state,
                }
              : null,
          };
        }
      } catch (error) {
        healthStatus.defaultDrive = {
          status: "access_error",
          error: "Unable to access default drive",
        };
      }
    }

    // Service capabilities
    healthStatus.capabilities = {
      fileOperations: true,
      excelManipulation: true,
      sharepointLists: true,
      searchFunctionality: true,
      sharingLinks: true,
    };

    return jsonTextResponse(healthStatus);
  } catch (error) {
    return toolErrorResponse("health_check", error);
  }
}

// Tool 2: Get user profile information
export const getUserProfile: Tool = {
  name: "get_user_profile",
  description: "Get detailed information about the authenticated user",
  inputSchema: {
    type: "object",
    properties: {
      includeManager: {
        type: "boolean",
        description: "Include manager information",
        default: false,
      },
      includePhoto: {
        type: "boolean",
        description: "Include profile photo metadata",
        default: false,
      },
    },
  },
};

export async function handleGetUserProfile(args: any) {
  try {
    const client = getGraphClient();
    const { includeManager = false, includePhoto = false } = args;

    // Get user profile
    const userResponse = await client.get<User>("/me");

    if (!userResponse.success || !userResponse.data) {
      throw new Error("Failed to retrieve user profile");
    }

    const user = userResponse.data;
    const profile: any = {
      id: user.id,
      displayName: user.displayName,
      mail: user.mail,
      userPrincipalName: user.userPrincipalName,
      jobTitle: user.jobTitle,
      department: user.department,
      officeLocation: (user as any).officeLocation,
      mobilePhone: (user as any).mobilePhone,
      businessPhones: (user as any).businessPhones,
      preferredLanguage: (user as any).preferredLanguage,
      country: (user as any).country,
      city: (user as any).city,
      companyName: (user as any).companyName,
    };

    // Include manager information if requested
    if (includeManager) {
      try {
        const managerResponse = await client.get<User>("/me/manager");
        if (managerResponse.success && managerResponse.data) {
          profile.manager = {
            id: managerResponse.data.id,
            displayName: managerResponse.data.displayName,
            mail: managerResponse.data.mail,
            jobTitle: managerResponse.data.jobTitle,
          };
        }
      } catch (error) {
        profile.manager = null;
      }
    }

    // Include photo metadata if requested
    if (includePhoto) {
      try {
        const photoResponse = await client.get<any>("/me/photo/$value");
        if (photoResponse.success) {
          profile.photo = {
            hasPhoto: true,
            size: "Available",
          };
        }
      } catch (error) {
        profile.photo = {
          hasPhoto: false,
        };
      }
    }

    return jsonTextResponse(profile);
  } catch (error) {
    return toolErrorResponse("get_user_profile", error);
  }
}

// Tool 3: List all accessible drives
export const listDrives: Tool = {
  name: "list_drives",
  description:
    "List all accessible drives (OneDrive + SharePoint document libraries)",
  inputSchema: {
    type: "object",
    properties: {
      includeQuota: {
        type: "boolean",
        description: "Include quota information for each drive",
        default: true,
      },
      siteId: {
        type: "string",
        description:
          "Optional SharePoint site ID to retrieve the default drive for a specific site",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
      driveType: {
        type: "string",
        enum: ["personal", "business", "documentLibrary", "all"],
        description: "Filter by drive type",
        default: "all",
      },
      limit: {
        type: "number",
        description: "Maximum number of drives to return",
        default: 50,
      },
    },
  },
};

export async function handleListDrives(args: any) {
  try {
    const client = getGraphClient();
    const { includeQuota = true, driveType = "all", limit = 50 } = args;

    if (args.site || args.siteId || args.siteUrl) {
      const resolvedSite = await resolveRequiredSharePointSite(args, client);
      const driveEndpoint = resolvedSite.driveId
        ? `/drives/${resolvedSite.driveId}`
        : `/sites/${resolvedSite.siteId}/drive`;
      const driveResponse = await client.get<Drive>(driveEndpoint);

      if (!driveResponse.success || !driveResponse.data) {
        throw new Error(
          "Failed to retrieve drive for resolved SharePoint site",
        );
      }

      const drive = driveResponse.data;
      const driveInfo: any = {
        id: drive.id,
        name: drive.name,
        driveType: drive.driveType,
        webUrl: drive.webUrl,
        description: drive.description,
        site: resolvedSite,
      };

      if (drive.owner) {
        driveInfo.owner = {
          user: drive.owner.user
            ? {
                displayName: drive.owner.user.displayName,
                id: drive.owner.user.id,
              }
            : null,
          group: drive.owner.group
            ? {
                displayName: drive.owner.group.displayName,
                id: drive.owner.group.id,
              }
            : null,
        };
      }

      if (includeQuota && drive.quota) {
        driveInfo.quota = {
          total: drive.quota.total,
          used: drive.quota.used,
          remaining: drive.quota.remaining,
          state: drive.quota.state,
          usagePercentage: drive.quota.total
            ? Math.round((drive.quota.used / drive.quota.total) * 100)
            : 0,
        };
      }

      return jsonTextResponse({
        driveCount: 1,
        filterType: driveType,
        site: resolvedSite,
        drives: [driveInfo],
      });
    }

    const drives: Drive[] = [];

    // Get personal OneDrive
    if (driveType === "all" || driveType === "personal") {
      try {
        const personalResponse = await client.get<Drive>("/me/drive");
        if (personalResponse.success && personalResponse.data) {
          drives.push(personalResponse.data);
        }
      } catch (error) {
        // Personal drive might not be accessible
      }
    }

    // Get all other drives
    const drivesEndpoint = "/me/drives";
    const params: any = {
      $top: limit.toString(),
    };

    if (driveType !== "all") {
      params["$filter"] = `driveType eq '${driveType}'`;
    }

    const drivesResponse = await client.get<GraphResponse<Drive>>(
      drivesEndpoint,
      params,
    );

    if (drivesResponse.success && drivesResponse.data) {
      const additionalDrives = (drivesResponse.data as any).value || [];
      // Avoid duplicates (personal drive might be included in both calls)
      const existingIds = new Set(drives.map((d) => d.id));
      additionalDrives.forEach((drive: Drive) => {
        if (!existingIds.has(drive.id)) {
          drives.push(drive);
        }
      });
    }

    const result = {
      driveCount: drives.length,
      filterType: driveType,
      drives: drives.map((drive: Drive) => {
        const driveInfo: any = {
          id: drive.id,
          name: drive.name,
          driveType: drive.driveType,
          webUrl: drive.webUrl,
          description: drive.description,
        };

        if (drive.owner) {
          driveInfo.owner = {
            user: drive.owner.user
              ? {
                  displayName: drive.owner.user.displayName,
                  id: drive.owner.user.id,
                }
              : null,
            group: drive.owner.group
              ? {
                  displayName: drive.owner.group.displayName,
                  id: drive.owner.group.id,
                }
              : null,
          };
        }

        if (includeQuota && drive.quota) {
          driveInfo.quota = {
            total: drive.quota.total,
            used: drive.quota.used,
            remaining: drive.quota.remaining,
            state: drive.quota.state,
            usagePercentage: drive.quota.total
              ? Math.round((drive.quota.used / drive.quota.total) * 100)
              : 0,
          };
        }

        return driveInfo;
      }),
    };

    return jsonTextResponse(result);
  } catch (error) {
    return toolErrorResponse("list_drives", error);
  }
}

// Tool 4: Global search across all content
export const globalSearch: Tool = {
  name: "global_search",
  description: "Search across all accessible content (files, lists, sites)",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description: "Search query string",
      },
      entityTypes: {
        type: "array",
        items: {
          type: "string",
          enum: ["driveItem", "site", "list", "listItem", "message", "event"],
        },
        description: "Types of entities to search for",
        default: ["driveItem"],
      },
      limit: {
        type: "number",
        description: "Maximum number of results per entity type",
        default: 20,
      },
      includeSummary: {
        type: "boolean",
        description: "Include content summary in results",
        default: true,
      },
    },
    required: ["query"],
  },
};

export async function handleGlobalSearch(args: any) {
  try {
    const client = getGraphClient();
    const {
      query,
      entityTypes = ["driveItem"],
      limit = 20,
      includeSummary = true,
    } = args;

    // Try Microsoft Search API first
    try {
      const searchRequest = {
        requests: entityTypes.map((entityType: string) => ({
          entityTypes: [entityType],
          query: {
            queryString: query,
          },
          from: 0,
          size: limit,
          fields: includeSummary
            ? ["*"]
            : ["name", "webUrl", "lastModifiedDateTime", "size"],
        })),
      };

      const searchResponse = await client.post<any>(
        "/search/query",
        searchRequest,
      );

      if (searchResponse.success && searchResponse.data) {
        const responses = searchResponse.data.value || [];
        const results: any = {
          query,
          entityTypes,
          totalResults: 0,
          results: {},
        };

        responses.forEach((response: any) => {
          const hits = response.hitsContainers?.[0]?.hits || [];
          const entityType =
            response.hitsContainers?.[0]?.hits?.[0]?.resource?.[
              "@odata.type"
            ] || "unknown";

          results.totalResults += hits.length;
          results.results[entityType] = hits.map((hit: any) => {
            const resource = hit.resource;
            const result: any = {
              id: resource.id,
              name: resource.name || resource.displayName,
              webUrl: resource.webUrl,
              lastModifiedDateTime: resource.lastModifiedDateTime,
              size: resource.size,
            };

            if (resource.file) {
              result.type = "file";
              result.mimeType = resource.file.mimeType;
            } else if (resource.folder) {
              result.type = "folder";
              result.childCount = resource.folder.childCount;
            }

            if (includeSummary && hit.summary) {
              result.summary = hit.summary;
            }

            return result;
          });
        });

        return jsonTextResponse(results);
      }
    } catch (searchError) {
      // Fall back to drive search if Microsoft Search fails
    }

    // Fallback: Search in personal OneDrive
    const fallbackEndpoint = `/me/drive/search(q='${encodeURIComponent(escapeODataString(query))}')`;
    const fallbackResponse = await client.get<GraphResponse<any>>(
      fallbackEndpoint,
      {
        $top: limit.toString(),
      },
    );

    if (fallbackResponse.success && fallbackResponse.data) {
      const items = (fallbackResponse.data as any).value || [];

      const results = {
        query,
        searchType: "fallback_onedrive_only",
        totalResults: items.length,
        results: items.map((item: any) => ({
          id: item.id,
          name: item.name,
          type: item.file ? "file" : "folder",
          webUrl: item.webUrl,
          lastModifiedDateTime: item.lastModifiedDateTime,
          size: item.size,
          mimeType: item.file?.mimeType,
          path: item.parentReference?.path,
        })),
      };

      return jsonTextResponse(results);
    }

    throw new Error("Search failed");
  } catch (error) {
    return toolErrorResponse("global_search", error);
  }
}

// Tool 5: Batch operations utility
export const batchOperations: Tool = {
  name: "batch_operations",
  description:
    "Execute multiple Graph API operations in a single batch request",
  inputSchema: {
    type: "object",
    properties: {
      requests: {
        type: "array",
        items: {
          type: "object",
          properties: {
            id: {
              type: "string",
              description: "Unique identifier for this request",
            },
            method: {
              type: "string",
              enum: ["GET", "POST", "PUT", "PATCH", "DELETE"],
              description: "HTTP method",
            },
            url: {
              type: "string",
              description: "Graph API endpoint (without base URL)",
            },
            body: {
              type: "object",
              description: "Request body (for POST/PUT/PATCH)",
            },
            headers: {
              type: "object",
              description: "Additional headers",
            },
          },
          required: ["id", "method", "url"],
        },
        description: "Array of requests to execute in batch",
        maxItems: 20,
      },
      continueOnError: {
        type: "boolean",
        description: "Continue processing other requests if one fails",
        default: true,
      },
    },
    required: ["requests"],
  },
};

export async function handleBatchOperations(args: any) {
  try {
    const client = getGraphClient();
    const { requests, continueOnError = true } = args;

    if (!requests || requests.length === 0) {
      throw new Error("At least one request is required");
    }

    if (requests.length > 20) {
      throw new Error("Maximum 20 requests allowed per batch");
    }

    const response = await client.batch(requests);

    if (response.success && response.data) {
      const batchResponses = response.data;

      const result: any = {
        batchSize: requests.length,
        continueOnError,
        timestamp: new Date().toISOString(),
        responses: batchResponses.map((batchResponse: any) => ({
          id: batchResponse.id,
          status: batchResponse.status,
          success: batchResponse.status >= 200 && batchResponse.status < 300,
          body: batchResponse.body,
          headers: batchResponse.headers,
        })),
      };

      // Calculate summary statistics
      const successCount = result.responses.filter(
        (r: any) => r.success,
      ).length;
      const failureCount = result.responses.length - successCount;

      result.summary = {
        totalRequests: result.responses.length,
        successful: successCount,
        failed: failureCount,
        successRate: Math.round((successCount / result.responses.length) * 100),
      };

      return jsonTextResponse(result);
    }

    throw new Error("Batch operation failed");
  } catch (error) {
    return toolErrorResponse("batch_operations", error);
  }
}

// Export all tools and handlers
export const utilityTools = [
  healthCheck,
  getUserProfile,
  listDrives,
  globalSearch,
  batchOperations,
];

export const utilityHandlers = {
  health_check: handleHealthCheck,
  get_user_profile: handleGetUserProfile,
  list_drives: handleListDrives,
  global_search: handleGlobalSearch,
  batch_operations: handleBatchOperations,
};
