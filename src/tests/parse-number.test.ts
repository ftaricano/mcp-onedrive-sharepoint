import test from "node:test";
import assert from "node:assert/strict";

import { parsePositiveInt } from "../utils/parse-number.js";

test("parsePositiveInt parses a valid positive integer string", () => {
  assert.equal(parsePositiveInt("30000", 5000, 60000), 30000);
});

test("parsePositiveInt parses a valid numeric value", () => {
  assert.equal(parsePositiveInt(42, 5000, 60000), 42);
});

test("parsePositiveInt returns fallback on NaN (malformed string)", () => {
  // Real-world case: a hostile/garbled Retry-After header.
  assert.equal(parsePositiveInt("not-a-number", 1000, 60000), 1000);
  assert.equal(parsePositiveInt("", 1000, 60000), 1000);
  assert.equal(parsePositiveInt(undefined, 1000, 60000), 1000);
  assert.equal(parsePositiveInt(NaN, 1000, 60000), 1000);
});

test("parsePositiveInt returns fallback on zero or negative values", () => {
  assert.equal(parsePositiveInt("0", 1000, 60000), 1000);
  assert.equal(parsePositiveInt("-5", 1000, 60000), 1000);
  assert.equal(parsePositiveInt(-1, 1000, 60000), 1000);
});

test("parsePositiveInt clamps finite values above the max to the max", () => {
  // The real hostile path: a giant *finite* Retry-After / config value.
  // Parses fine, then gets clamped so it cannot hang the process.
  assert.equal(parsePositiveInt("999999999", 1000, 60000), 60000);
  assert.equal(parsePositiveInt(999999999, 1000, 60000), 60000);
});

test("parsePositiveInt returns fallback on non-finite Infinity input", () => {
  // Infinity is not a meaningful integer; fall back rather than clamp.
  assert.equal(parsePositiveInt(Infinity, 1000, 60000), 1000);
  assert.equal(parsePositiveInt(-Infinity, 1000, 60000), 1000);
});

test("parsePositiveInt accepts the boundary value equal to max", () => {
  assert.equal(parsePositiveInt("60000", 1000, 60000), 60000);
});

test("parsePositiveInt truncates fractional input via base-10 parse", () => {
  assert.equal(parsePositiveInt("12.9", 1000, 60000), 12);
});
