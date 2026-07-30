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

  // `_op_get` (e não `get`/`get_item`): leitura 1P pura com contrato idêntico
  // na main e na branch do hub scripts — sem cache plaintext em disco e sem
  // stale-fallback (o `get()` da main tem os dois), e sem depender da ordem
  // de merge do JAR-424.
  assert.match(source, /from cpz_keychain import _op_get/);
  assert.doesNotMatch(source, /import get\b|get_item/);
  // Interpretadores fixos e isolados no caminho que manipula o segredo.
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
