/**
 * SharePoint lists management tools
 * Comprehensive CRUD operations for SharePoint lists and list items
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../../graph/client.js";
import {
  Site,
  List,
  ListItem,
  ListColumn,
  ContentType,
  GraphResponse,
} from "../../graph/models.js";
import {
  extractPaginatedResult,
  jsonTextResponse,
  toolErrorResponse,
} from "../../graph/contracts.js";
import { SecurityValidator } from "../../utils/security-validator.js";
import {
  getKnownSharePointSites,
  resolveRequiredSharePointSite,
  resolveSharePointSiteReference,
} from "../../sharepoint/site-resolver.js";

// Tool 1: Discover SharePoint sites
export const discoverSites: Tool = {
  name: "discover_sites",
  description: "Discover SharePoint sites accessible by the user",
  inputSchema: {
    type: "object",
    properties: {
      search: {
        type: "string",
        description: "Search term to filter sites by name or description",
      },
      limit: {
        type: "number",
        description: "Maximum number of sites to return",
        default: 50,
      },
      includePersonalSite: {
        type: "boolean",
        description:
          "Also include the tenant root SharePoint site (`/sites/root`) when available",
        default: false,
      },
      pageToken: {
        type: "string",
        description:
          "Opaque pagination token from a previous response (Graph nextLink)",
      },
    },
  },
};

function buildDiscoverSitesEndpoint(search?: string): {
  endpoint: string;
  searchLabel: string;
} {
  if (!search) {
    return {
      endpoint: "/sites?search=*",
      searchLabel: "all sites",
    };
  }

  const validation = SecurityValidator.validateSearchQuery(search);
  if (!validation.isValid || !validation.sanitized) {
    throw new Error(validation.error || "Invalid search query");
  }

  return {
    endpoint: `/sites?search=${encodeURIComponent(validation.sanitized)}`,
    searchLabel: validation.sanitized,
  };
}

export async function handleDiscoverSites(args: any) {
  try {
    const client = getGraphClient();
    const { search, limit = 50, includePersonalSite = false, pageToken } = args;

    let response;
    let searchLabel = "all sites";

    if (pageToken) {
      response = await client.get<GraphResponse<Site>>(pageToken);
    } else {
      const discoveryRequest = buildDiscoverSitesEndpoint(search);
      searchLabel = discoveryRequest.searchLabel;
      response = await client.get<GraphResponse<Site>>(
        discoveryRequest.endpoint,
        {
          $top: limit.toString(),
        },
      );
    }

    if (!response.success || !response.data) {
      throw new Error("Failed to retrieve sites");
    }

    const { items, pagination } = extractPaginatedResult<Site>(
      response.data,
      limit,
    );
    const sites: Site[] = [...items];

    if (includePersonalSite && !pageToken) {
      try {
        const personalSiteResponse = await client.get<Site>("/sites/root");
        if (personalSiteResponse.success && personalSiteResponse.data) {
          const alreadyIncluded = sites.some(
            (site) => site.id === personalSiteResponse.data?.id,
          );
          if (!alreadyIncluded) {
            sites.unshift(personalSiteResponse.data);
          }
        }
      } catch {
        // Personal site access may not be available, continue without it
      }
    }

    return jsonTextResponse({
      search: searchLabel,
      siteCount: sites.length,
      includePersonalSite,
      pagination: {
        ...pagination,
        returned: sites.length,
      },
      sites: sites.map((site: Site) => ({
        id: site.id,
        name: site.name || site.displayName,
        displayName: site.displayName,
        description: site.description,
        webUrl: site.webUrl,
        createdDateTime: site.createdDateTime,
        lastModifiedDateTime: site.lastModifiedDateTime,
        isRoot: !!site.root,
      })),
    });
  } catch (error) {
    return toolErrorResponse("discover_sites", error);
  }
}

export const resolveSite: Tool = {
  name: "resolve_site",
  description:
    "Resolve a SharePoint site from canonical aliases, siteId, or canonical URL without relying on discover_sites",
  inputSchema: {
    type: "object",
    properties: {
      site: {
        type: "string",
        description: "Known alias, site name hint, or canonical SharePoint URL",
      },
      siteId: {
        type: "string",
        description: "SharePoint site ID (passes through if already known)",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
    },
  },
};

export async function handleResolveSite(args: any) {
  try {
    const client = getGraphClient();
    const resolved = await resolveSharePointSiteReference(args, client);

    if (!resolved) {
      return jsonTextResponse({
        resolved: false,
        input: args,
        knownSites: getKnownSharePointSites().map((site) => ({
          key: site.key,
          name: site.name,
          siteId: site.siteId,
          siteUrl: site.siteUrl,
          driveId: site.driveId,
          aliases: site.aliases,
        })),
      });
    }

    return jsonTextResponse({
      resolved: true,
      site: resolved,
      knownSites: getKnownSharePointSites().map((site) => ({
        key: site.key,
        name: site.name,
        siteId: site.siteId,
        siteUrl: site.siteUrl,
        driveId: site.driveId,
        aliases: site.aliases,
      })),
    });
  } catch (error) {
    return toolErrorResponse("resolve_site", error);
  }
}

// Tool 2: List SharePoint lists in a site
export const listSiteLists: Tool = {
  name: "list_site_lists",
  description: "List all SharePoint lists in a specific site",
  inputSchema: {
    type: "object",
    properties: {
      siteId: {
        type: "string",
        description: "SharePoint site ID",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
      includeHidden: {
        type: "boolean",
        description: "Include hidden lists",
        default: false,
      },
      includeSystemLists: {
        type: "boolean",
        description: "Include system lists (like Workflow Tasks)",
        default: false,
      },
      limit: {
        type: "number",
        description: "Maximum number of lists to return",
        default: 100,
      },
      pageToken: {
        type: "string",
        description:
          "Opaque pagination token from a previous response (Graph nextLink)",
      },
    },
    required: [],
  },
};

export async function handleListSiteLists(args: any) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const {
      includeHidden = false,
      includeSystemLists = false,
      limit = 100,
      pageToken,
    } = args;

    const endpoint = pageToken || `/sites/${siteId}/lists`;
    const params: any = {
      $top: limit.toString(),
      $expand: "columns,contentTypes",
      $orderby: "displayName",
    };

    const response = await client.get<GraphResponse<List>>(
      endpoint,
      pageToken ? undefined : params,
    );

    if (response.success && response.data) {
      const { items: allLists, pagination } = extractPaginatedResult(
        response.data,
        limit,
      );

      // Client-side filtering — Graph API /sites/{siteId}/lists has
      // limited $filter support so we filter after fetching.
      const lists = allLists.filter((list: List) => {
        if (!includeHidden && list.list?.hidden) return false;
        if (!includeSystemLists && list.displayName?.startsWith("_"))
          return false;
        return true;
      });

      return jsonTextResponse({
        siteId,
        site: resolvedSite,
        listCount: lists.length,
        includeHidden,
        includeSystemLists,
        pagination,
        lists: lists.map((list: List) => ({
          id: list.id,
          name: list.name,
          displayName: list.displayName,
          description: list.description,
          webUrl: list.webUrl,
          template: list.list?.template,
          hidden: list.list?.hidden,
          contentTypesEnabled: list.list?.contentTypesEnabled,
          columnCount: list.columns?.length || 0,
          contentTypeCount: list.contentTypes?.length || 0,
          createdDateTime: list.createdDateTime,
          lastModifiedDateTime: list.lastModifiedDateTime,
        })),
      });
    }

    throw new Error("Failed to retrieve lists");
  } catch (error) {
    return toolErrorResponse("list_site_lists", error);
  }
}

// Tool 3: Get list schema (columns and content types)
export const getListSchema: Tool = {
  name: "get_list_schema",
  description: "Get detailed schema information for a SharePoint list",
  inputSchema: {
    type: "object",
    properties: {
      siteId: {
        type: "string",
        description: "SharePoint site ID",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
      listId: {
        type: "string",
        description: "SharePoint list ID",
      },
      includeContentTypes: {
        type: "boolean",
        description: "Include content type information",
        default: true,
      },
    },
    required: ["listId"],
  },
};

export async function handleGetListSchema(args: any) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { listId, includeContentTypes = true } = args;

    // Get list details with expanded information
    const listEndpoint = `/sites/${siteId}/lists/${listId}`;
    const expandItems = ["columns"];
    if (includeContentTypes) {
      expandItems.push("contentTypes");
    }

    const response = await client.get<List>(listEndpoint, {
      $expand: expandItems.join(","),
    });

    if (response.success && response.data) {
      const list = response.data;

      return jsonTextResponse({
        siteId,
        site: resolvedSite,
        list: {
          id: list.id,
          name: list.name,
          displayName: list.displayName,
          description: list.description,
          webUrl: list.webUrl,
          template: list.list?.template,
          hidden: list.list?.hidden,
          contentTypesEnabled: list.list?.contentTypesEnabled,
        },
        columns:
          list.columns?.map((column: ListColumn) => ({
            id: column.id,
            name: column.name,
            displayName: column.displayName,
            description: column.description,
            type: column.type,
            required: column.required,
            hidden: column.hidden,
            indexed: column.indexed,
            textSettings: column.text,
            numberSettings: column.number,
            choiceSettings: column.choice,
            lookupSettings: column.lookup,
          })) || [],
        contentTypes: includeContentTypes
          ? list.contentTypes?.map((ct: ContentType) => ({
              id: ct.id,
              name: ct.name,
              description: ct.description,
              group: ct.group,
              hidden: ct.hidden,
              readOnly: ct.readOnly,
              sealed: ct.sealed,
            })) || []
          : undefined,
      });
    }

    throw new Error("Failed to get list schema");
  } catch (error) {
    return toolErrorResponse("get_list_schema", error);
  }
}

// Tool 4: List items from a SharePoint list
export const listItems: Tool = {
  name: "list_items",
  description:
    "List items from a SharePoint list with filtering and pagination",
  inputSchema: {
    type: "object",
    properties: {
      siteId: {
        type: "string",
        description: "SharePoint site ID",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
      listId: {
        type: "string",
        description: "SharePoint list ID",
      },
      filter: {
        type: "string",
        description: "OData filter expression (e.g., \"Title eq 'Example'\")",
      },
      orderBy: {
        type: "string",
        description: 'Sort order (e.g., "Title", "Created desc")',
        default: "Created desc",
      },
      select: {
        type: "string",
        description:
          'Comma-separated list of fields to return (e.g., "Title,Author,Created")',
      },
      expand: {
        type: "string",
        description: "Comma-separated list of lookup fields to expand",
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
    required: ["listId"],
  },
};

export async function handleListItems(args: any) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const {
      listId,
      filter,
      orderBy = "Created desc",
      select,
      expand,
      limit = 100,
      pageToken,
    } = args;

    const endpoint = pageToken || `/sites/${siteId}/lists/${listId}/items`;
    const params: any = {
      $top: limit.toString(),
      $expand: "fields",
      $orderby: orderBy,
    };

    if (filter) {
      params["$filter"] = filter;
    }

    if (select) {
      params["$select"] =
        `id,webUrl,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,contentType,fields(${select})`;
    }

    if (expand) {
      params["$expand"] += `,fields(${expand})`;
    }

    const response = await client.get<GraphResponse<ListItem>>(
      endpoint,
      pageToken ? undefined : params,
    );

    if (response.success && response.data) {
      const { items, pagination } = extractPaginatedResult(
        response.data,
        limit,
      );

      return jsonTextResponse({
        siteId,
        site: resolvedSite,
        listId,
        filter: filter || "none",
        orderBy,
        itemCount: items.length,
        pagination,
        items: items.map((item: ListItem) => ({
          id: item.id,
          webUrl: item.webUrl,
          createdDateTime: item.createdDateTime,
          lastModifiedDateTime: item.lastModifiedDateTime,
          createdBy: item.createdBy?.user?.displayName,
          lastModifiedBy: item.lastModifiedBy?.user?.displayName,
          contentType: item.contentType?.name,
          fields: item.fields,
        })),
      });
    }

    throw new Error("Failed to retrieve list items");
  } catch (error) {
    return toolErrorResponse("list_items", error);
  }
}

// Tool 5: Get a specific list item
export const getListItem: Tool = {
  name: "get_list_item",
  description: "Get a specific item from a SharePoint list",
  inputSchema: {
    type: "object",
    properties: {
      siteId: {
        type: "string",
        description: "SharePoint site ID",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
      listId: {
        type: "string",
        description: "SharePoint list ID",
      },
      itemId: {
        type: "string",
        description: "List item ID",
      },
      expand: {
        type: "string",
        description: "Comma-separated list of lookup fields to expand",
      },
    },
    required: ["listId", "itemId"],
  },
};

export async function handleGetListItem(args: any) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { listId, itemId, expand } = args;

    const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}`;
    const params: any = {
      $expand: "fields",
    };

    if (expand) {
      params["$expand"] += `,fields(${expand})`;
    }

    const response = await client.get<ListItem>(endpoint, params);

    if (response.success && response.data) {
      const item = response.data;

      return jsonTextResponse({
        siteId,
        site: resolvedSite,
        id: item.id,
        webUrl: item.webUrl,
        createdDateTime: item.createdDateTime,
        lastModifiedDateTime: item.lastModifiedDateTime,
        createdBy: item.createdBy?.user?.displayName,
        lastModifiedBy: item.lastModifiedBy?.user?.displayName,
        contentType: item.contentType?.name,
        parentReference: item.parentReference,
        sharepointIds: item.sharepointIds,
        fields: item.fields,
      });
    }

    throw new Error("Failed to get list item");
  } catch (error) {
    return toolErrorResponse("get_list_item", error);
  }
}

// Tool 6: Create a new list item
export const createListItem: Tool = {
  name: "create_list_item",
  description: "Create a new item in a SharePoint list",
  inputSchema: {
    type: "object",
    properties: {
      siteId: {
        type: "string",
        description: "SharePoint site ID",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
      listId: {
        type: "string",
        description: "SharePoint list ID",
      },
      fields: {
        type: "object",
        description: "Field values for the new item (key-value pairs)",
        additionalProperties: true,
      },
      contentType: {
        type: "string",
        description: "Content type ID (optional)",
      },
    },
    required: ["listId", "fields"],
  },
};

export async function handleCreateListItem(args: any) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { listId, fields, contentType } = args;

    const endpoint = `/sites/${siteId}/lists/${listId}/items`;
    const itemData: any = {
      fields,
    };

    if (contentType) {
      itemData.contentType = { id: contentType };
    }

    const response = await client.post<ListItem>(endpoint, itemData);

    if (response.success && response.data) {
      const item = response.data;

      return jsonTextResponse({
        success: true,
        message: "List item created successfully",
        siteId,
        site: resolvedSite,
        item: {
          id: item.id,
          webUrl: item.webUrl,
          createdDateTime: item.createdDateTime,
          contentType: item.contentType?.name,
          fields: item.fields,
        },
      });
    }

    throw new Error("Failed to create list item");
  } catch (error) {
    return toolErrorResponse("create_list_item", error);
  }
}

// Tool 7: Update a list item
export const updateListItem: Tool = {
  name: "update_list_item",
  description: "Update an existing item in a SharePoint list",
  inputSchema: {
    type: "object",
    properties: {
      siteId: {
        type: "string",
        description: "SharePoint site ID",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
      listId: {
        type: "string",
        description: "SharePoint list ID",
      },
      itemId: {
        type: "string",
        description: "List item ID to update",
      },
      fields: {
        type: "object",
        description: "Field values to update (key-value pairs)",
        additionalProperties: true,
      },
    },
    required: ["listId", "itemId", "fields"],
  },
};

export async function handleUpdateListItem(args: any) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { listId, itemId, fields } = args;

    const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`;

    const response = await client.patch<any>(endpoint, fields);

    if (response.success && response.data) {
      return jsonTextResponse({
        success: true,
        message: "List item updated successfully",
        siteId,
        site: resolvedSite,
        itemId,
        updatedFields: fields,
        result: response.data,
      });
    }

    throw new Error("Failed to update list item");
  } catch (error) {
    return toolErrorResponse("update_list_item", error);
  }
}

// Tool 8: Delete a list item
export const deleteListItem: Tool = {
  name: "delete_list_item",
  description: "Delete an item from a SharePoint list",
  inputSchema: {
    type: "object",
    properties: {
      siteId: {
        type: "string",
        description: "SharePoint site ID",
      },
      site: {
        type: "string",
        description: "Known SharePoint site alias or canonical URL",
      },
      siteUrl: {
        type: "string",
        description: "Canonical SharePoint site URL",
      },
      listId: {
        type: "string",
        description: "SharePoint list ID",
      },
      itemId: {
        type: "string",
        description: "List item ID to delete",
      },
    },
    required: ["listId", "itemId"],
  },
};

export async function handleDeleteListItem(args: any) {
  try {
    const client = getGraphClient();
    const resolvedSite = await resolveRequiredSharePointSite(args, client);
    const siteId = resolvedSite.siteId;
    const { listId, itemId } = args;

    const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}`;

    const response = await client.delete(endpoint);

    if (response.success) {
      return jsonTextResponse({
        success: true,
        message: "List item deleted successfully",
        itemId,
        siteId,
        site: resolvedSite,
        listId,
      });
    }

    throw new Error("Failed to delete list item");
  } catch (error) {
    return toolErrorResponse("delete_list_item", error);
  }
}

// Export all tools and handlers
export const sharepointTools = [
  discoverSites,
  resolveSite,
  listSiteLists,
  getListSchema,
  listItems,
  getListItem,
  createListItem,
  updateListItem,
  deleteListItem,
];

export const sharepointHandlers = {
  discover_sites: handleDiscoverSites,
  resolve_site: handleResolveSite,
  list_site_lists: handleListSiteLists,
  get_list_schema: handleGetListSchema,
  list_items: handleListItems,
  get_list_item: handleGetListItem,
  create_list_item: handleCreateListItem,
  update_list_item: handleUpdateListItem,
  delete_list_item: handleDeleteListItem,
};
