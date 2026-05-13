/**
 * Advanced collaboration and permissions management tools
 * Enhanced features for team collaboration and access control
 */

import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { getGraphClient } from "../../graph/client.js";
import { jsonTextResponse, toolErrorResponse } from "../../graph/contracts.js";
import { DriveItem, Permission, User } from "../../graph/models.js";
import {
  buildDriveItemEndpoint,
  getDriveRootEndpoint,
} from "../../graph/resource-resolver.js";
import { resolveDriveTargetContext } from "../../sharepoint/site-resolver.js";

// Tool 1: Advanced sharing with email notifications
export const advancedShare: Tool = {
  name: "advanced_share",
  description:
    "Create advanced sharing links with custom permissions and email notifications",
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
      recipients: {
        type: "array",
        items: { type: "string" },
        description: "Email addresses of recipients",
      },
      permission: {
        type: "string",
        enum: ["read", "write", "owner"],
        description: "Permission level",
        default: "read",
      },
      requireSignIn: {
        type: "boolean",
        description: "Require sign-in to access",
        default: true,
      },
      sendInvitation: {
        type: "boolean",
        description: "Send email invitation",
        default: true,
      },
      message: {
        type: "string",
        description: "Custom message for invitation email",
      },
      expirationDateTime: {
        type: "string",
        description: "Link expiration (ISO 8601 format)",
      },
      retainInheritedPermissions: {
        type: "boolean",
        description: "Keep existing inherited permissions",
        default: true,
      },
    },
    required: ["recipients"],
  },
};

export async function handleAdvancedShare(args: any) {
  try {
    const client = getGraphClient();
    const {
      itemId,
      itemPath,
      recipients,
      permission = "read",
      requireSignIn = true,
      sendInvitation = true,
      message,
      expirationDateTime,
      retainInheritedPermissions = true,
    } = args;
    const { siteId, driveId } = await resolveDriveTargetContext(
      { site: args.site, siteId: args.siteId, siteUrl: args.siteUrl, driveId: args.driveId },
      client,
    );

    if (!itemId && !itemPath) {
      throw new Error("Either itemId or itemPath is required");
    }

    const endpoint = buildDriveItemEndpoint(
      { itemId, itemPath, siteId, driveId },
      "/invite",
    );

    const inviteData: any = {
      recipients: recipients.map((email: string) => ({
        email,
      })),
      requireSignIn,
      sendInvitation,
      retainInheritedPermissions,
      roles: permission === "owner" ? ["owner"] : [permission],
    };

    if (message) {
      inviteData.message = message;
    }

    if (expirationDateTime) {
      inviteData.expirationDateTime = expirationDateTime;
    }

    const response = await client.post<any>(endpoint, inviteData);

    if (response.success && response.data) {
      const permissions = response.data.value || [];

      return jsonTextResponse({
        success: true,
        message: "Sharing invitations sent successfully",
        recipientCount: recipients.length,
        sendInvitation,
        permissions: permissions.map((perm: any) => ({
          id: perm.id,
          grantedTo:
            perm.grantedTo?.user?.email || perm.grantedTo?.user?.displayName,
          roles: perm.roles,
          hasPassword: perm.hasPassword,
          expirationDateTime: perm.expirationDateTime,
          shareId: perm.shareId,
        })),
      });
    }

    throw new Error("Failed to share item");
  } catch (error) {
    return toolErrorResponse("advanced_share", error);
  }
}

// Tool 2: Manage permissions
export const managePermissions: Tool = {
  name: "manage_permissions",
  description: "List, update, or revoke permissions for a file or folder",
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
      action: {
        type: "string",
        enum: ["list", "update", "revoke"],
        description: "Action to perform",
        default: "list",
      },
      permissionId: {
        type: "string",
        description: "Permission ID (for update/revoke actions)",
      },
      newRoles: {
        type: "array",
        items: { type: "string" },
        description: "New roles for update action (read, write, owner)",
      },
    },
    required: ["action"],
  },
};

export async function handleManagePermissions(args: any) {
  try {
    const client = getGraphClient();
    const {
      itemId,
      itemPath,
      action = "list",
      permissionId,
      newRoles,
    } = args;
    const { siteId, driveId } = await resolveDriveTargetContext(
      { site: args.site, siteId: args.siteId, siteUrl: args.siteUrl, driveId: args.driveId },
      client,
    );

    if (!itemId && !itemPath) {
      throw new Error("Either itemId or itemPath is required");
    }

    const baseEndpoint = buildDriveItemEndpoint(
      { itemId, itemPath, siteId, driveId },
    );

    switch (action) {
      case "list": {
        const endpoint = `${baseEndpoint}/permissions`;
        const response = await client.get<any>(endpoint);

        if (response.success && response.data) {
          const permissions = (response.data as any).value || [];

          return jsonTextResponse({
            action: "list",
            permissionCount: permissions.length,
            permissions: permissions.map((perm: any) => ({
              id: perm.id,
              roles: perm.roles,
              grantedTo: {
                user: perm.grantedTo?.user
                  ? {
                      email: perm.grantedTo.user.email,
                      displayName: perm.grantedTo.user.displayName,
                      id: perm.grantedTo.user.id,
                    }
                  : null,
                application: perm.grantedTo?.application
                  ? {
                      displayName: perm.grantedTo.application.displayName,
                      id: perm.grantedTo.application.id,
                    }
                  : null,
              },
              link: perm.link
                ? {
                    type: perm.link.type,
                    scope: perm.link.scope,
                    webUrl: perm.link.webUrl,
                  }
                : null,
              inheritedFrom: perm.inheritedFrom,
              expirationDateTime: perm.expirationDateTime,
              hasPassword: perm.hasPassword,
            })),
          });
        }
        break;
      }

      case "update": {
        if (!permissionId || !newRoles) {
          throw new Error(
            "permissionId and newRoles are required for update action",
          );
        }

        const endpoint = `${baseEndpoint}/permissions/${permissionId}`;
        const updateData = {
          roles: newRoles,
        };

        const response = await client.patch<Permission>(endpoint, updateData);

        if (response.success && response.data) {
          return jsonTextResponse({
            action: "update",
            success: true,
            message: "Permission updated successfully",
            permissionId,
            newRoles,
            result: response.data,
          });
        }
        break;
      }

      case "revoke": {
        if (!permissionId) {
          throw new Error("permissionId is required for revoke action");
        }

        const endpoint = `${baseEndpoint}/permissions/${permissionId}`;
        const response = await client.delete(endpoint);

        if (response.success) {
          return jsonTextResponse({
            action: "revoke",
            success: true,
            message: "Permission revoked successfully",
            permissionId,
          });
        }
        break;
      }

      default:
        throw new Error(`Invalid action: ${action}`);
    }

    throw new Error(`Failed to ${action} permissions`);
  } catch (error) {
    return toolErrorResponse("manage_permissions", error);
  }
}

// Tool 3: Check user access
export const checkUserAccess: Tool = {
  name: "check_user_access",
  description: "Check what access a specific user has to a file or folder",
  inputSchema: {
    type: "object",
    properties: {
      itemId: {
        type: "string",
        description: "Item ID to check",
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
      userEmail: {
        type: "string",
        description: "Email address of the user to check",
      },
      includeInherited: {
        type: "boolean",
        description: "Include inherited permissions",
        default: true,
      },
    },
    required: ["userEmail"],
  },
};

export async function handleCheckUserAccess(args: any) {
  try {
    const client = getGraphClient();
    const {
      itemId,
      itemPath,
      userEmail,
      includeInherited = true,
    } = args;
    const { siteId, driveId } = await resolveDriveTargetContext(
      { site: args.site, siteId: args.siteId, siteUrl: args.siteUrl, driveId: args.driveId },
      client,
    );

    if (!itemId && !itemPath) {
      throw new Error("Either itemId or itemPath is required");
    }

    const baseEndpoint = buildDriveItemEndpoint(
      { itemId, itemPath, siteId, driveId },
    );

    // Get all permissions for the item
    const permissionsEndpoint = `${baseEndpoint}/permissions`;
    const response = await client.get<any>(permissionsEndpoint);

    if (response.success && response.data) {
      const allPermissions = (response.data as any).value || [];

      // Filter permissions for the specific user
      const userPermissions = allPermissions.filter((perm: any) => {
        // Check direct user permissions
        if (
          perm.grantedTo?.user?.email?.toLowerCase() === userEmail.toLowerCase()
        ) {
          return true;
        }

        // Check granted to identities (for shared links)
        if (perm.grantedToIdentities) {
          return perm.grantedToIdentities.some(
            (identity: any) =>
              identity.user?.email?.toLowerCase() === userEmail.toLowerCase(),
          );
        }

        return false;
      });

      // Check for inherited permissions if requested
      let inheritedPermissions: any[] = [];
      if (includeInherited) {
        inheritedPermissions = allPermissions.filter(
          (perm: any) => perm.inheritedFrom && !userPermissions.includes(perm),
        );
      }

      const hasAccess =
        userPermissions.length > 0 || inheritedPermissions.length > 0;
      const effectiveRoles = new Set<string>();

      [...userPermissions, ...inheritedPermissions].forEach((perm: any) => {
        if (perm.roles) {
          perm.roles.forEach((role: string) => effectiveRoles.add(role));
        }
      });

      return jsonTextResponse({
        userEmail,
        hasAccess,
        effectiveRoles: Array.from(effectiveRoles),
        directPermissions: userPermissions.map((perm: any) => ({
          id: perm.id,
          roles: perm.roles,
          type: perm.link ? "link" : "direct",
          expirationDateTime: perm.expirationDateTime,
        })),
        inheritedPermissions: inheritedPermissions.map((perm: any) => ({
          id: perm.id,
          roles: perm.roles,
          inheritedFrom: perm.inheritedFrom,
        })),
        summary: {
          canRead:
            effectiveRoles.has("read") ||
            effectiveRoles.has("write") ||
            effectiveRoles.has("owner"),
          canWrite: effectiveRoles.has("write") || effectiveRoles.has("owner"),
          isOwner: effectiveRoles.has("owner"),
        },
      });
    }

    throw new Error("Failed to check user access");
  } catch (error) {
    return toolErrorResponse("check_user_access", error);
  }
}

// Export all collaboration tools and handlers
export const collaborationTools = [
  advancedShare,
  managePermissions,
  checkUserAccess,
];

export const collaborationHandlers = {
  advanced_share: handleAdvancedShare,
  manage_permissions: handleManagePermissions,
  check_user_access: handleCheckUserAccess,
};
