import test from "node:test";
import assert from "node:assert/strict";
import { pathToFileURL } from "node:url";
import { resolve } from "node:path";

import { isMainModule } from "../index.js";

test("isMainModule matches resolved filesystem paths", () => {
  const entryPath = resolve("build/index.js");
  assert.equal(isMainModule(pathToFileURL(entryPath).href, entryPath), true);
});

test("isMainModule does not match a different entry path", () => {
  const modulePath = resolve("build/index.js");
  const entryPath = resolve("build/cli.js");
  assert.equal(isMainModule(pathToFileURL(modulePath).href, entryPath), false);
});

test("isMainModule handles missing argv path", () => {
  const modulePath = resolve("build/index.js");
  assert.equal(isMainModule(pathToFileURL(modulePath).href, undefined), false);
});
