import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

test("operational wrappers use the canonical 1Password-only helper", () => {
  const source = ["scripts/ods.sh", "scripts/run-stdio.sh", "scripts/spcall.sh", "scripts/onepassword-graph-env.sh", "scripts/with-onepassword-graph-env.sh"]
    .map((relativePath) => fs.readFileSync(path.join(root, relativePath), "utf8"))
    .join("\n");

  assert.match(source, /from cpz_keychain import get_item/);
  assert.doesNotMatch(source, /find-generic-password/);
  assert.doesNotMatch(source, /unlock-keychain/);
  assert.doesNotMatch(source, /exec-with-env/);
});

test("runtime source has no dotenv, Keychain, or file credential backend", () => {
  const source = [
    "src/auth/microsoft-graph-auth.ts",
    "src/config/index.ts",
    "src/auth/setup-auth.ts",
  ]
    .map((relativePath) => fs.readFileSync(path.join(root, relativePath), "utf8"))
    .join("\n");

  assert.doesNotMatch(source, /dotenv/);
  assert.doesNotMatch(source, /keytar/);
  assert.doesNotMatch(source, /FileFallbackStore/);
  assert.doesNotMatch(source, /readFile|writeFile|mkdir/);
});
