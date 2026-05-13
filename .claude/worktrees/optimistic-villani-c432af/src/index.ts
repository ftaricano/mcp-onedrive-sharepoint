#!/usr/bin/env node

/**
 * MCP OneDrive/SharePoint Server
 * Unified Microsoft Graph API server for OneDrive and SharePoint operations
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from "@modelcontextprotocol/sdk/types.js";
import { Tool } from "@modelcontextprotocol/sdk/types.js";

// Import all tool categories
import { fileTools, fileHandlers } from "./tools/files/index.js";
import {
  sharepointTools,
  sharepointHandlers,
} from "./tools/sharepoint/index.js";
import { utilityTools, utilityHandlers } from "./tools/utils/index.js";

// Import advanced tools
import { advancedTools, advancedHandlers } from "./tools/advanced/index.js";

// Import configuration and initialization
import { bootstrap, prewarmAuth } from "./core/bootstrap.js";
import { resetGraphClient } from "./graph/client.js";
import { createUserFriendlyError } from "./graph/error-handler.js";
import { toolErrorResponse } from "./graph/contracts.js";

class McpOneDriveSharePointServer {
  private server: Server;

  constructor() {
    this.server = new Server(
      {
        name: "mcp-onedrive-sharepoint",
        version: "1.0.0",
        description:
          "Microsoft Graph MCP Server for OneDrive and SharePoint operations",
      },
      {
        capabilities: {
          tools: {},
        },
      },
    );

    this.setupRequestHandlers();
  }

  private setupRequestHandlers(): void {
    // List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: this.getAllTools(),
      };
    });

    // Handle tool execution
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      try {
        await this.ensureInitialized();
        return await this.handleToolCall(
          request.params.name,
          request.params.arguments || {},
        );
      } catch (error) {
        throw new McpError(
          ErrorCode.InternalError,
          `Error executing tool ${request.params.name}: ${createUserFriendlyError(error)}`,
        );
      }
    });
  }

  private getAllTools(): Tool[] {
    return [
      ...fileTools,
      ...sharepointTools,
      ...utilityTools,
      ...advancedTools,
    ];
  }

  private getAllHandlers(): Record<string, Function> {
    return {
      ...fileHandlers,
      ...sharepointHandlers,
      ...utilityHandlers,
      ...advancedHandlers,
    };
  }

  private async ensureInitialized(): Promise<void> {
    try {
      await bootstrap();
    } catch (error) {
      console.error(
        "Failed to initialize server:",
        createUserFriendlyError(error),
      );
      throw error;
    }
  }

  private async handleToolCall(toolName: string, args: any): Promise<any> {
    const handlers = this.getAllHandlers();
    const handler = handlers[toolName];

    if (!handler) {
      throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${toolName}`);
    }

    try {
      const result = await handler(args);

      // Ensure result has the expected MCP response format
      if (!result || typeof result !== "object") {
        throw new Error("Handler returned invalid response format");
      }

      if (!result.content || !Array.isArray(result.content)) {
        throw new Error("Handler response missing required content array");
      }

      return result;
    } catch (error) {
      console.error(`Error in tool ${toolName}:`, error);

      // Return error in MCP format
      return toolErrorResponse(toolName, error);
    }
  }

  async start(): Promise<void> {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("MCP OneDrive/SharePoint Server running on stdio");

    // Prewarm auth in background so the first tool call finds a hot token.
    prewarmAuth();
  }

  async cleanup(): Promise<void> {
    try {
      // Clean up Graph client resources
      resetGraphClient();
      console.error("Server cleanup completed");
    } catch (error) {
      console.error("Error during cleanup:", error);
    }
  }
}

// Handle process signals for graceful shutdown
async function setupSignalHandlers(
  server: McpOneDriveSharePointServer,
): Promise<void> {
  const signals = ["SIGINT", "SIGTERM", "SIGQUIT"];

  for (const signal of signals) {
    process.on(signal, async () => {
      console.error(`Received ${signal}, shutting down gracefully...`);
      await server.cleanup();
      process.exit(0);
    });
  }

  // Handle uncaught exceptions
  process.on("uncaughtException", async (error) => {
    console.error("Uncaught exception:", error);
    await server.cleanup();
    process.exit(1);
  });

  // Handle unhandled promise rejections
  process.on("unhandledRejection", async (reason, promise) => {
    console.error("Unhandled rejection at:", promise, "reason:", reason);
    await server.cleanup();
    process.exit(1);
  });
}

// Main execution
async function main(): Promise<void> {
  try {
    const server = new McpOneDriveSharePointServer();

    // Setup signal handlers for graceful shutdown
    await setupSignalHandlers(server);

    // Start the server
    await server.start();
  } catch (error) {
    console.error("Failed to start MCP OneDrive/SharePoint Server:", error);
    process.exit(1);
  }
}

// Export for testing
export { McpOneDriveSharePointServer };

// Run if this is the main module
if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    console.error("Fatal error:", error);
    process.exit(1);
  });
}
