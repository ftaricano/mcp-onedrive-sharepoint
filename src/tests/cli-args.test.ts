import test from "node:test";
import assert from "node:assert/strict";

import { coerce, parseArgs, buildArgs, extractText } from "../cli/args.js";

test("coerce returns true for undefined value (bare flag)", () => {
  assert.equal(coerce(undefined), true);
});

test("coerce handles boolean and null literals", () => {
  assert.equal(coerce("true"), true);
  assert.equal(coerce("false"), false);
  assert.equal(coerce("null"), null);
});

test("coerce parses integers and floats", () => {
  assert.equal(coerce("42"), 42);
  assert.equal(coerce("-7"), -7);
  assert.equal(coerce("3.14"), 3.14);
  assert.equal(coerce("-0.5"), -0.5);
});

test("coerce leaves non-numeric strings untouched", () => {
  assert.equal(coerce("financeiro"), "financeiro");
  assert.equal(coerce("/Shared Documents"), "/Shared Documents");
  assert.equal(coerce(""), "");
});

test("coerce does not treat version-like strings as numbers", () => {
  assert.equal(coerce("1.2.3"), "1.2.3");
});

test("parseArgs supports --key=value", () => {
  const result = parseArgs(["--site=financeiro", "--path=/x"]);
  assert.deepEqual(result, { site: "financeiro", path: "/x" });
});

test("parseArgs supports --key value", () => {
  const result = parseArgs(["--site", "financeiro", "--limit", "50"]);
  assert.deepEqual(result, { site: "financeiro", limit: 50 });
});

test("parseArgs treats trailing bare flag as true", () => {
  const result = parseArgs(["--path", "/x", "--recurse"]);
  assert.deepEqual(result, { path: "/x", recurse: true });
});

test("parseArgs treats flag followed by flag as boolean true", () => {
  const result = parseArgs(["--recurse", "--path", "/x"]);
  assert.deepEqual(result, { recurse: true, path: "/x" });
});

test("parseArgs ignores non-flag tokens", () => {
  const result = parseArgs(["positional", "--site", "financeiro"]);
  assert.deepEqual(result, { site: "financeiro" });
});

test("parseArgs handles empty argv", () => {
  assert.deepEqual(parseArgs([]), {});
});

test("buildArgs returns parsed as-is when no --json given", () => {
  const parsed = { site: "financeiro", limit: 50 };
  assert.deepEqual(buildArgs(parsed), parsed);
});

test("buildArgs merges --json payload with CLI flags taking precedence", () => {
  const parsed = {
    json: '{"driveId":"abc","path":"/old"}',
    path: "/new",
  };
  assert.deepEqual(buildArgs(parsed), {
    driveId: "abc",
    path: "/new",
  });
});

test("buildArgs throws on invalid JSON payload", () => {
  assert.throws(
    () => buildArgs({ json: "{not-json" }),
    /Invalid --json payload/,
  );
});

test("buildArgs rejects non-object JSON payloads", () => {
  assert.throws(
    () => buildArgs({ json: "[1,2,3]" }),
    /must be a JSON object/,
  );
  assert.throws(() => buildArgs({ json: "42" }), /must be a JSON object/);
  assert.throws(() => buildArgs({ json: "null" }), /must be a JSON object/);
});

test("extractText pulls text out of MCP content array", () => {
  const result = {
    content: [{ type: "text", text: '{"ok":true}' }],
  };
  assert.equal(extractText(result), '{"ok":true}');
});

test("extractText joins multi-part content with newlines", () => {
  const result = {
    content: [
      { type: "text", text: "first" },
      { type: "text", text: "second" },
    ],
  };
  assert.equal(extractText(result), "first\nsecond");
});

test("extractText falls back to JSON when result has no content array", () => {
  assert.equal(extractText({ foo: 1 }), '{\n  "foo": 1\n}');
  assert.equal(extractText(null), "null");
  assert.equal(extractText("plain"), '"plain"');
});
