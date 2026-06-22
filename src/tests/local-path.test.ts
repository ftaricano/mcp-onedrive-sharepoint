import test from "node:test";
import assert from "node:assert/strict";
import {
  mkdirSync,
  mkdtempSync,
  realpathSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { resolveLocalPath } from "../utils/local-path.js";

function withLocalRoot<T>(root: string, fn: () => T): T {
  const previous = process.env.MCP_LOCAL_FILE_ROOT;
  process.env.MCP_LOCAL_FILE_ROOT = root;

  try {
    return fn();
  } finally {
    if (previous === undefined) {
      delete process.env.MCP_LOCAL_FILE_ROOT;
    } else {
      process.env.MCP_LOCAL_FILE_ROOT = previous;
    }
  }
}

test("resolveLocalPath keeps relative paths inside MCP_LOCAL_FILE_ROOT", () => {
  const root = mkdtempSync(path.join(tmpdir(), "mcp-local-root-"));

  try {
    withLocalRoot(root, () => {
      assert.equal(
        resolveLocalPath("downloads/report.txt"),
        path.join(realpathSync(root), "downloads/report.txt"),
      );
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("resolveLocalPath rejects absolute paths outside MCP_LOCAL_FILE_ROOT", () => {
  const root = mkdtempSync(path.join(tmpdir(), "mcp-local-root-"));

  try {
    withLocalRoot(root, () => {
      assert.throws(
        () => resolveLocalPath("/etc/passwd"),
        /outside MCP_LOCAL_FILE_ROOT/,
      );
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test("resolveLocalPath rejects symlink escapes outside MCP_LOCAL_FILE_ROOT", () => {
  const root = mkdtempSync(path.join(tmpdir(), "mcp-local-root-"));
  const outside = mkdtempSync(path.join(tmpdir(), "mcp-outside-root-"));

  try {
    mkdirSync(path.join(root, "safe"));
    writeFileSync(path.join(outside, "secret.txt"), "secret");
    symlinkSync(outside, path.join(root, "safe", "link-out"));

    withLocalRoot(root, () => {
      assert.throws(
        () => resolveLocalPath("safe/link-out/secret.txt", { mustExist: true }),
        /resolves outside MCP_LOCAL_FILE_ROOT/,
      );
      assert.throws(
        () => resolveLocalPath("safe/link-out/new.txt"),
        /resolves outside MCP_LOCAL_FILE_ROOT/,
      );
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
    rmSync(outside, { recursive: true, force: true });
  }
});

test("resolveLocalPath validates existence when requested", () => {
  const root = mkdtempSync(path.join(tmpdir(), "mcp-local-root-"));

  try {
    const filePath = path.join(root, "input.txt");
    writeFileSync(filePath, "hello");

    withLocalRoot(root, () => {
      assert.equal(
        resolveLocalPath("input.txt", { mustExist: true }),
        path.join(realpathSync(root), "input.txt"),
      );
      assert.throws(
        () => resolveLocalPath("missing.txt", { mustExist: true }),
        /does not exist/,
      );
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
