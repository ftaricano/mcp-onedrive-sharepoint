import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);

// Supply-chain floor guard: axios >= 1.18 carries security patches for URL
// parsing / redirect handling, which matter here because the Graph client
// builds request URLs via buildUrl(). Keeping this assertion prevents a future
// lockfile regression from silently dropping back below the patched line.
test("axios resolves to a version that includes the >=1.18 security patches", () => {
  const { version } = require("axios/package.json") as { version: string };
  const [major, minor] = version.split(".").map((part) => Number.parseInt(part, 10));

  assert.ok(
    major > 1 || (major === 1 && minor >= 18),
    `axios must be >= 1.18.0 (URL/redirect security patches), got ${version}`,
  );
});
