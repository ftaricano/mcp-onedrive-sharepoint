#!/usr/bin/env node

/**
 * CLI adapter — exposes the same MCP tools as plain subcommands.
 * Usage:
 *   ods <tool-name> --arg=value [--json '{...}']
 *   ods list
 *   ods schema <tool-name>
 *   ods auth
 */

import { bootstrap } from "./core/bootstrap.js";
import { fileTools, fileHandlers } from "./tools/files/index.js";
import {
  sharepointTools,
  sharepointHandlers,
} from "./tools/sharepoint/index.js";
import { utilityTools, utilityHandlers } from "./tools/utils/index.js";
import { advancedTools, advancedHandlers } from "./tools/advanced/index.js";
import { createUserFriendlyError } from "./graph/error-handler.js";
import { parseArgs, buildArgs, extractText } from "./cli/args.js";
import { runAuthSetup } from "./cli/auth-command.js";

const allTools = [
  ...fileTools,
  ...sharepointTools,
  ...utilityTools,
  ...advancedTools,
];
const allHandlers: Record<string, (args: any) => Promise<any>> = {
  ...fileHandlers,
  ...sharepointHandlers,
  ...utilityHandlers,
  ...advancedHandlers,
} as Record<string, (args: any) => Promise<any>>;

function printUsage(): void {
  process.stderr.write(
    [
      "Usage:",
      "  ods <tool-name> --key=value [--key value] [--json '<payload>']",
      "  ods list              List all available tools",
      "  ods schema <tool>     Print JSON schema for a tool",
      "  ods auth              Run interactive Microsoft Graph auth setup",
      "  ods help              Print this message",
      "",
    ].join("\n"),
  );
}

async function main(): Promise<void> {
  const [, , command, ...rest] = process.argv;

  if (!command || command === "help" || command === "--help" || command === "-h") {
    printUsage();
    process.exit(command ? 0 : 1);
  }

  if (command === "list") {
    for (const tool of allTools) {
      process.stdout.write(`${tool.name}\t${tool.description ?? ""}\n`);
    }
    return;
  }

  if (command === "schema") {
    const name = rest[0];
    const tool = allTools.find((t) => t.name === name);
    if (!tool) {
      process.stderr.write(`Unknown tool: ${name}\n`);
      process.exit(1);
    }
    process.stdout.write(JSON.stringify(tool.inputSchema, null, 2) + "\n");
    return;
  }

  if (command === "auth") {
    await runAuthSetup();
    return;
  }

  const handler = allHandlers[command];
  if (!handler) {
    process.stderr.write(`Unknown tool: ${command}\nTry: ods list\n`);
    process.exit(1);
  }

  const parsed = parseArgs(rest);
  const args = buildArgs(parsed);

  await bootstrap();
  const result = await handler(args);

  const text = extractText(result);
  process.stdout.write(text + "\n");

  if ((result as { isError?: boolean } | undefined)?.isError) process.exit(2);
}

main().catch((error) => {
  process.stderr.write(`Error: ${createUserFriendlyError(error)}\n`);
  process.exit(1);
});
