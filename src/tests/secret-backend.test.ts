import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

test("operational wrappers use the canonical 1Password-only helper", () => {
  const source = ["scripts/ods.sh", "scripts/run-stdio.sh", "scripts/onepassword-graph-env.sh"]
    .map((relativePath) => fs.readFileSync(path.join(root, relativePath), "utf8"))
    .join("\n");

  assert.match(source, /from cpz_keychain import get/);
  assert.doesNotMatch(source, /find-generic-password/);
  assert.doesNotMatch(source, /unlock-keychain/);
});

test("dotenv loader discards persisted Graph client secrets", () => {
  const source = fs.readFileSync(
    path.join(root, "scripts/exec-with-env.mjs"),
    "utf8",
  );

  assert.match(source, /delete fileEnv\.MICROSOFT_GRAPH_CLIENT_SECRET/);
  assert.match(source, /delete fileEnv\.SP_CLIENT_SECRET/);
});
