import test from "node:test";
import assert from "node:assert/strict";
import { getToolRegistry } from "../tools/registry.js";

function withEnv<T>(
  updates: Record<string, string | undefined>,
  fn: () => T,
): T {
  const previous = Object.fromEntries(
    Object.keys(updates).map((key) => [key, process.env[key]]),
  );

  try {
    for (const [key, value] of Object.entries(updates)) {
      if (value === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = value;
      }
    }

    return fn();
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = value;
      }
    }
  }
}

test("tool registry defaults to the safe core public tool surface", () => {
  withEnv(
    {
      MCP_TOOL_PROFILE: undefined,
      MCP_DISABLED_TOOLS: undefined,
      MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH: undefined,
    },
    () => {
      const registry = getToolRegistry();
      const names = registry.tools.map((tool) => tool.name);

      assert.equal(registry.profile, "core");
      assert.equal(names.length, 10);
      assert.ok(names.includes("list_files"));
      assert.ok(names.includes("upload_file"));
      assert.ok(!names.includes("delete_item"));
      assert.ok(!names.includes("excel_analysis"));
      assert.equal(registry.handlers.excel_analysis, undefined);
    },
  );
});

test("full profile exposes advanced tools but keeps raw Graph batch experimental", () => {
  withEnv(
    {
      MCP_TOOL_PROFILE: "full",
      MCP_DISABLED_TOOLS: undefined,
      MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH: undefined,
    },
    () => {
      const registry = getToolRegistry();
      const names = registry.tools.map((tool) => tool.name);

      assert.equal(registry.profile, "full");
      assert.equal(names.length, 32);
      assert.ok(names.includes("delete_item"));
      assert.ok(names.includes("excel_analysis"));
      assert.ok(!names.includes("batch_operations"));
      assert.equal(registry.handlers.batch_operations, undefined);
      assert.equal(typeof registry.handlers.excel_analysis, "function");
    },
  );
});

test("experimental Graph batch tool requires explicit opt-in", () => {
  withEnv(
    {
      MCP_TOOL_PROFILE: "full",
      MCP_DISABLED_TOOLS: undefined,
      MCP_ENABLE_EXPERIMENTAL_GRAPH_BATCH: "true",
    },
    () => {
      const registry = getToolRegistry();
      const names = registry.tools.map((tool) => tool.name);

      assert.equal(names.length, 33);
      assert.ok(names.includes("batch_operations"));
      assert.equal(typeof registry.handlers.batch_operations, "function");
    },
  );
});

test("core profile hides advanced and destructive tools without removing core handlers", () => {
  withEnv({ MCP_TOOL_PROFILE: "core", MCP_DISABLED_TOOLS: undefined }, () => {
    const registry = getToolRegistry();
    const names = registry.tools.map((tool) => tool.name);

    assert.equal(registry.profile, "core");
    assert.equal(names.length, 10);
    assert.ok(names.includes("list_files"));
    assert.ok(!names.includes("list_items"));
    assert.ok(!names.includes("delete_item"));
    assert.ok(!names.includes("excel_analysis"));
    assert.equal(registry.handlers.excel_analysis, undefined);
    assert.equal(typeof registry.handlers.list_files, "function");
  });
});

test("disabled tools are removed from tool list and handlers", () => {
  withEnv(
    {
      MCP_TOOL_PROFILE: "full",
      MCP_DISABLED_TOOLS: "sync_folder, delete_item",
    },
    () => {
      const registry = getToolRegistry();
      const names = registry.tools.map((tool) => tool.name);

      assert.ok(!names.includes("sync_folder"));
      assert.ok(!names.includes("delete_item"));
      assert.equal(registry.handlers.sync_folder, undefined);
      assert.equal(registry.handlers.delete_item, undefined);
      assert.equal(typeof registry.handlers.list_files, "function");
    },
  );
});

test("invalid tool profile is rejected early", () => {
  withEnv({ MCP_TOOL_PROFILE: "everything" }, () => {
    assert.throws(() => getToolRegistry(), /Invalid MCP_TOOL_PROFILE/);
  });
});
