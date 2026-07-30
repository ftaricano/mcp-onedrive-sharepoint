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
import { getToolRegistry } from "./tools/registry.js";
import { createUserFriendlyError } from "./graph/error-handler.js";
import { parseArgs, buildArgs, extractText } from "./cli/args.js";
import { runAuthSetup } from "./cli/auth-command.js";

function printUsage(): void {
  process.stderr.write(
    [
      "Usage:",
      "  ods <tool-name> --key=value [--key value] [--json '<payload>']",
      "  ods list              List all available tools",
      "  ods schema <tool>     Print JSON schema for a tool",
      "  ods auth              (disabled — client credentials come from 1Password)",
      "  ods help              Print this message",
      "",
    ].join("\n"),
  );
}

async function main(): Promise<void> {
  const [, , command, ...rest] = process.argv;
  const registry = getToolRegistry();

  if (
    !command ||
    command === "help" ||
    command === "--help" ||
    command === "-h"
  ) {
    printUsage();
    process.exit(command ? 0 : 1);
  }

  if (command === "list") {
    for (const tool of registry.tools) {
      process.stdout.write(`${tool.name}\t${tool.description ?? ""}\n`);
    }
    return;
  }

  if (command === "schema") {
    const name = rest[0];
    const tool = registry.tools.find((t) => t.name === name);
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

  const handler = registry.handlers[command];
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
