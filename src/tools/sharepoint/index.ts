/**
 * SharePoint lists management tools
 * Comprehensive CRUD operations for SharePoint lists and list items
 */

import { Tool } from '@modelcontextprotocol/sdk/types.js';
import { getGraphClient } from '../../graph/client.js';
import { Site, List, ListItem, ListColumn, ContentType, GraphResponse } from '../../graph/models.js';
import { extractPaginatedResult, jsonTextResponse, toolErrorResponse } from '../../graph/contracts.js';
import { createUserFriendlyError } from '../../graph/error-handler.js';

// Tool 1: Discover SharePoint sites
export const discoverSites: Tool = {
  name: 'discover_sites',
  description: 'Discover SharePoint sites accessible by the user',
  inputSchema: {
    type: 'object',
    properties: {
      search: {
        type: 'string',
        description: 'Search term to filter sites by name or description'
      },
      limit: {
        type: 'number',
        description: 'Maximum number of sites to return',
        default: 50
      },
      includePersonalSite: {
        type: 'boolean',
        description: 'Also include the tenant root SharePoint site (`/sites/root`) when available',
        default: false
      },
      pageToken: {
        type: 'string',
        description: 'Opaque pagination token from a previous response (Graph nextLink)'
      }
    }
  }
};

export async function handleDiscoverSites(args: any) {
  try {
    const client = getGraphClient();
    const { search, limit = 50, includePersonalSite = false, pageToken } = args;

    let response;

    if (pageToken) {
      response = await client.get<GraphResponse<Site>>(pageToken);
    } else if (search) {
      response = await client.get<GraphResponse<Site>>('/sites', {
        search,
        '$top': limit.toString()
      });
    } else {
      response = await client.get<GraphResponse<Site>>('/sites', {
        '$filter': 'siteCollection/root ne null',
        '$top': limit.toString(),
        '$orderby': 'displayName'
      });
    }

    if (!response.success || !response.data) {
      throw new Error('Failed to retrieve sites');
    }

    const { items: sites, pagination } = extractPaginatedResult(response.data, limit);

    if (includePersonalSite && !pageToken) {
      try {
        const personalSiteResponse = await client.get<Site>('/sites/root');
        if (personalSiteResponse.success && personalSiteResponse.data) {
          const alreadyIncluded = sites.some((site) => site.id === personalSiteResponse.data?.id);
          if (!alreadyIncluded) {
            sites.unshift(personalSiteResponse.data);
          }
        }
      } catch {
        // Personal site access may not be available, continue without it
      }
    }

    return jsonTextResponse({
      search: search || 'all sites',
      siteCount: sites.length,
      includePersonalSite,
      pagination: {
        ...pagination,
        returned: sites.length
      },
      sites: sites.map((site: Site) => ({
        id: site.id,
        name: site.name || site.displayName,
        displayName: site.displayName,
        description: site.description,
        webUrl: site.webUrl,
        createdDateTime: site.createdDateTime,
        lastModifiedDateTime: site.lastModifiedDateTime,
        isRoot: !!site.root
      }))
    });
  } catch (error) {
    return toolErrorResponse('discover_sites', error);
  }
}

// Tool 2: List SharePoint lists in a site
export const listSiteLists: Tool = {
  name: 'list_site_lists',
  description: 'List all SharePoint lists in a specific site',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID'
      },
      includeHidden: {
        type: 'boolean',
        description: 'Include hidden lists',
        default: false
      },
      includeSystemLists: {
        type: 'boolean',
        description: 'Include system lists (like Workflow Tasks)',
        default: false
      },
      limit: {
        type: 'number',
        description: 'Maximum number of lists to return',
        default: 100
      },
      pageToken: {
        type: 'string',
        description: 'Opaque pagination token from a previous response (Graph nextLink)'
      }
    },
    required: ['siteId']
  }
};

export async function handleListSiteLists(args: any) {
  try {
    const client = getGraphClient();
    const { siteId, includeHidden = false, includeSystemLists = false, limit = 100, pageToken } = args;

    const endpoint = pageToken || `/sites/${siteId}/lists`;
    const params: any = {
      '$top': limit.toString(),
      '$expand': 'columns,contentTypes',
      '$orderby': 'displayName'
    };

    const filters: string[] = [];
    if (!includeHidden) {
      filters.push('list/hidden eq false');
    }
    if (!includeSystemLists) {
      filters.push("not startswith(displayName,'_')");
    }
    if (filters.length > 0) {
      params['$filter'] = filters.join(' and ');
    }

    const response = await client.get<GraphResponse<List>>(endpoint, pageToken ? undefined : params);

    if (response.success && response.data) {
      const { items: lists, pagination } = extractPaginatedResult(response.data, limit);
      
      return jsonTextResponse({
        siteId,
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
          lastModifiedDateTime: list.lastModifiedDateTime
        }))
      });
    }

    throw new Error('Failed to retrieve lists');
  } catch (error) {
    return toolErrorResponse('list_site_lists', error);
  }
}

// Tool 3: Get list schema (columns and content types)
export const getListSchema: Tool = {
  name: 'get_list_schema',
  description: 'Get detailed schema information for a SharePoint list',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID'
      },
      listId: {
        type: 'string',
        description: 'SharePoint list ID'
      },
      includeContentTypes: {
        type: 'boolean',
        description: 'Include content type information',
        default: true
      }
    },
    required: ['siteId', 'listId']
  }
};

export async function handleGetListSchema(args: any) {
  try {
    const client = getGraphClient();
    const { siteId, listId, includeContentTypes = true } = args;

    // Get list details with expanded information
    const listEndpoint = `/sites/${siteId}/lists/${listId}`;
    const expandItems = ['columns'];
    if (includeContentTypes) {
      expandItems.push('contentTypes');
    }

    const response = await client.get<List>(listEndpoint, {
      '$expand': expandItems.join(',')
    });

    if (response.success && response.data) {
      const list = response.data;
      
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            list: {
              id: list.id,
              name: list.name,
              displayName: list.displayName,
              description: list.description,
              webUrl: list.webUrl,
              template: list.list?.template,
              hidden: list.list?.hidden,
              contentTypesEnabled: list.list?.contentTypesEnabled
            },
            columns: list.columns?.map((column: ListColumn) => ({
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
              lookupSettings: column.lookup
            })) || [],
            contentTypes: includeContentTypes ? (list.contentTypes?.map((ct: ContentType) => ({
              id: ct.id,
              name: ct.name,
              description: ct.description,
              group: ct.group,
              hidden: ct.hidden,
              readOnly: ct.readOnly,
              sealed: ct.sealed
            })) || []) : undefined
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to get list schema');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error getting list schema: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 4: List items from a SharePoint list
export const listItems: Tool = {
  name: 'list_items',
  description: 'List items from a SharePoint list with filtering and pagination',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID'
      },
      listId: {
        type: 'string',
        description: 'SharePoint list ID'
      },
      filter: {
        type: 'string',
        description: 'OData filter expression (e.g., "Title eq \'Example\'")'
      },
      orderBy: {
        type: 'string',
        description: 'Sort order (e.g., "Title", "Created desc")',
        default: 'Created desc'
      },
      select: {
        type: 'string',
        description: 'Comma-separated list of fields to return (e.g., "Title,Author,Created")'
      },
      expand: {
        type: 'string',
        description: 'Comma-separated list of lookup fields to expand'
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
    },
    required: ['siteId', 'listId']
  }
};

export async function handleListItems(args: any) {
  try {
    const client = getGraphClient();
    const { siteId, listId, filter, orderBy = 'Created desc', select, expand, limit = 100, pageToken } = args;

    const endpoint = pageToken || `/sites/${siteId}/lists/${listId}/items`;
    const params: any = {
      '$top': limit.toString(),
      '$expand': 'fields',
      '$orderby': orderBy
    };

    if (filter) {
      params['$filter'] = filter;
    }

    if (select) {
      params['$select'] = `id,webUrl,createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,contentType,fields(${select})`;
    }

    if (expand) {
      params['$expand'] += `,fields(${expand})`;
    }

    const response = await client.get<GraphResponse<ListItem>>(endpoint, pageToken ? undefined : params);

    if (response.success && response.data) {
      const { items, pagination } = extractPaginatedResult(response.data, limit);
      
      return jsonTextResponse({
        siteId,
        listId,
        filter: filter || 'none',
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
          fields: item.fields
        }))
      });
    }

    throw new Error('Failed to retrieve list items');
  } catch (error) {
    return toolErrorResponse('list_items', error);
  }
}

// Tool 5: Get a specific list item
export const getListItem: Tool = {
  name: 'get_list_item',
  description: 'Get a specific item from a SharePoint list',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID'
      },
      listId: {
        type: 'string',
        description: 'SharePoint list ID'
      },
      itemId: {
        type: 'string',
        description: 'List item ID'
      },
      expand: {
        type: 'string',
        description: 'Comma-separated list of lookup fields to expand'
      }
    },
    required: ['siteId', 'listId', 'itemId']
  }
};

export async function handleGetListItem(args: any) {
  try {
    const client = getGraphClient();
    const { siteId, listId, itemId, expand } = args;

    const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}`;
    const params: any = {
      '$expand': 'fields'
    };

    if (expand) {
      params['$expand'] += `,fields(${expand})`;
    }

    const response = await client.get<ListItem>(endpoint, params);

    if (response.success && response.data) {
      const item = response.data;
      
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            id: item.id,
            webUrl: item.webUrl,
            createdDateTime: item.createdDateTime,
            lastModifiedDateTime: item.lastModifiedDateTime,
            createdBy: item.createdBy?.user?.displayName,
            lastModifiedBy: item.lastModifiedBy?.user?.displayName,
            contentType: item.contentType?.name,
            parentReference: item.parentReference,
            sharepointIds: item.sharepointIds,
            fields: item.fields
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to get list item');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error getting list item: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 6: Create a new list item
export const createListItem: Tool = {
  name: 'create_list_item',
  description: 'Create a new item in a SharePoint list',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID'
      },
      listId: {
        type: 'string',
        description: 'SharePoint list ID'
      },
      fields: {
        type: 'object',
        description: 'Field values for the new item (key-value pairs)',
        additionalProperties: true
      },
      contentType: {
        type: 'string',
        description: 'Content type ID (optional)'
      }
    },
    required: ['siteId', 'listId', 'fields']
  }
};

export async function handleCreateListItem(args: any) {
  try {
    const client = getGraphClient();
    const { siteId, listId, fields, contentType } = args;

    const endpoint = `/sites/${siteId}/lists/${listId}/items`;
    const itemData: any = {
      fields
    };

    if (contentType) {
      itemData.contentType = { id: contentType };
    }

    const response = await client.post<ListItem>(endpoint, itemData);

    if (response.success && response.data) {
      const item = response.data;
      
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'List item created successfully',
            item: {
              id: item.id,
              webUrl: item.webUrl,
              createdDateTime: item.createdDateTime,
              contentType: item.contentType?.name,
              fields: item.fields
            }
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to create list item');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error creating list item: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 7: Update a list item
export const updateListItem: Tool = {
  name: 'update_list_item',
  description: 'Update an existing item in a SharePoint list',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID'
      },
      listId: {
        type: 'string',
        description: 'SharePoint list ID'
      },
      itemId: {
        type: 'string',
        description: 'List item ID to update'
      },
      fields: {
        type: 'object',
        description: 'Field values to update (key-value pairs)',
        additionalProperties: true
      }
    },
    required: ['siteId', 'listId', 'itemId', 'fields']
  }
};

export async function handleUpdateListItem(args: any) {
  try {
    const client = getGraphClient();
    const { siteId, listId, itemId, fields } = args;

    const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`;
    
    const response = await client.patch<any>(endpoint, fields);

    if (response.success && response.data) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'List item updated successfully',
            itemId,
            updatedFields: fields,
            result: response.data
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to update list item');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error updating list item: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Tool 8: Delete a list item
export const deleteListItem: Tool = {
  name: 'delete_list_item',
  description: 'Delete an item from a SharePoint list',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID'
      },
      listId: {
        type: 'string',
        description: 'SharePoint list ID'
      },
      itemId: {
        type: 'string',
        description: 'List item ID to delete'
      }
    },
    required: ['siteId', 'listId', 'itemId']
  }
};

export async function handleDeleteListItem(args: any) {
  try {
    const client = getGraphClient();
    const { siteId, listId, itemId } = args;

    const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}`;
    
    const response = await client.delete(endpoint);

    if (response.success) {
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'List item deleted successfully',
            itemId,
            siteId,
            listId
          }, null, 2)
        }]
      };
    }

    throw new Error('Failed to delete list item');
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error deleting list item: ${createUserFriendlyError(error)}`
      }],
      isError: true
    };
  }
}

// Export all tools and handlers
export const sharepointTools = [
  discoverSites,
  listSiteLists,
  getListSchema,
  listItems,
  getListItem,
  createListItem,
  updateListItem,
  deleteListItem
];

export const sharepointHandlers = {
  discover_sites: handleDiscoverSites,
  list_site_lists: handleListSiteLists,
  get_list_schema: handleGetListSchema,
  list_items: handleListItems,
  get_list_item: handleGetListItem,
  create_list_item: handleCreateListItem,
  update_list_item: handleUpdateListItem,
  delete_list_item: handleDeleteListItem
};