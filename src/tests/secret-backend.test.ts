import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

test("legacy launchers read secrets without a disk cache or stale fallback", () => {
  const source = ["scripts/ods.sh", "scripts/run-stdio.sh", "scripts/spcall.sh", "scripts/onepassword-graph-env.sh", "scripts/with-onepassword-graph-env.sh"]
    .map((relativePath) => fs.readFileSync(path.join(root, relativePath), "utf8"))
    .join("\n");

  // `_op_get` (not `get`/`get_item`): a direct password-manager read with no
  // plaintext cache on disk and no stale-value fallback.
  assert.match(source, /from \w+ import _op_get/);
  assert.doesNotMatch(source, /import get\b|get_item/);
  // Fixed, isolated interpreters on the path that handles the secret.
  assert.match(source, /\/usr\/bin\/python3 -I /);
  assert.doesNotMatch(source, /^ *python3 -c/m);
  assert.doesNotMatch(source, /#!\/usr\/bin\/env bash/);
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
