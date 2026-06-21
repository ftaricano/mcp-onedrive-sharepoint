import { Tool } from "@modelcontextprotocol/sdk/types.js";
import { fileTools, fileHandlers } from "./files/index.js";
import { sharepointTools, sharepointHandlers } from "./sharepoint/index.js";
import { utilityTools, utilityHandlers } from "./utils/index.js";
import { advancedTools, advancedHandlers } from "./advanced/index.js";

export type ToolHandler = (args: any) => Promise<any>;
export type ToolProfile = "core" | "full";

export interface ToolRegistry {
  tools: Tool[];
  handlers: Record<string, ToolHandler>;
  profile: ToolProfile;
  disabledTools: string[];
}

const CORE_TOOL_NAMES = new Set([
  "health_check",
  "list_drives",
  "discover_sites",
  "resolve_site",
  "list_files",
  "search_files",
  "get_file_metadata",
  "download_file",
  "upload_file",
  "create_folder",
]);

const EXPERIMENTAL_TOOL_NAMES = new Set(["batch_operations"]);

const ALL_TOOLS = [
  ...fileTools,
  ...sharepointTools,
  ...utilityTools,
  ...advancedTools,
];
const FULL_HANDLERS = {
  ...fileHandlers,
  ...sharepointHandlers,
  ...utilityHandlers,
  ...advancedHandlers,
} as Record<string, ToolHandler>;

function readProfile(): ToolProfile {
  const raw = (process.env.MCP_TOOL_PROFILE || "core").trim().toLowerCase();

  if (raw === "core" || raw === "full") {
    return raw;
  }

  throw new Error(
    `Invalid MCP_TOOL_PROFILE="${raw}". Expected "core" or "full".`,
  );
}

function readDisabledTools(): string[] {
  return (process.env.MCP_DISABLED_TOOLS || "")
    .split(",")
    .map((name) => name.trim())
    .filter(Boolean);
}

function experimentalToolsEnabled(): boolean {
  return process.env.MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH === "true";
}

export function getToolRegistry(): ToolRegistry {
  const profile = readProfile();
  const disabledTools = readDisabledTools();
  const disabled = new Set(disabledTools);

  const enabledToolNames =
    profile === "core"
      ? CORE_TOOL_NAMES
      : new Set(
          ALL_TOOLS.map((tool) => tool.name).filter(
            (name) =>
              !EXPERIMENTAL_TOOL_NAMES.has(name) || experimentalToolsEnabled(),
          ),
        );

  return {
    profile,
    disabledTools,
    tools: ALL_TOOLS.filter(
      (tool) => enabledToolNames.has(tool.name) && !disabled.has(tool.name),
    ),
    handlers: Object.fromEntries(
      Object.entries(FULL_HANDLERS).filter(
        ([name]) => enabledToolNames.has(name) && !disabled.has(name),
      ),
    ),
  };
}
