import test from "node:test";
import assert from "node:assert/strict";

import { GraphClient, MAX_RETRY_AFTER_SECONDS } from "../graph/client.js";

/**
 * Regression guard for the `Retry-After` DoS vector.
 *
 * The response interceptor reads the server-provided `Retry-After` header and
 * turns it into `rateLimitDelay` (ms), which is later fed to `setTimeout`. A
 * hostile or buggy server returning a huge value ("999999999") would otherwise
 * schedule a multi-year timer and hang the process. The delay must be parsed
 * safely, reject NaN/negative values, and clamp to MAX_RETRY_AFTER_SECONDS.
 */

// `updateRateLimitInfo` and `rateLimitDelay` are private; exercise them via a
// typed cast, matching the existing `(client as any)` pattern in
// graph-contracts.test.ts. Construction is inert (no network) so this is safe.
function delayForRetryAfter(retryAfter: unknown): number {
  const client = new GraphClient();
  (client as any).updateRateLimitInfo({ headers: { "retry-after": retryAfter } });
  return (client as any).rateLimitDelay as number;
}

const MAX_DELAY_MS = MAX_RETRY_AFTER_SECONDS * 1000;

test("MAX_RETRY_AFTER_SECONDS caps the delay at 300_000 ms", () => {
  assert.equal(MAX_RETRY_AFTER_SECONDS, 300);
  assert.equal(MAX_DELAY_MS, 300_000);
});

test("giant Retry-After ('999999999') is clamped to <= 300_000 ms", () => {
  const delay = delayForRetryAfter("999999999");
  assert.ok(
    delay <= MAX_DELAY_MS,
    `expected delay <= ${MAX_DELAY_MS} ms, got ${delay}`,
  );
  // Specifically pinned to the cap, not merely "small".
  assert.equal(delay, MAX_DELAY_MS);
});

test("a sane Retry-After is honoured verbatim (seconds -> ms)", () => {
  assert.equal(delayForRetryAfter("5"), 5_000);
  // Boundary value equal to the cap stays put.
  assert.equal(delayForRetryAfter(String(MAX_RETRY_AFTER_SECONDS)), MAX_DELAY_MS);
});

test("malformed / NaN Retry-After does not become an absurd delay", () => {
  assert.equal(delayForRetryAfter("not-a-number"), 0);
  // An HTTP-date (only seconds are supported here) is non-numeric -> no delay,
  // never a multi-year hang.
  assert.equal(delayForRetryAfter("Wed, 21 Oct 2099 07:28:00 GMT"), 0);
});

test("negative or zero Retry-After does not become an absurd delay", () => {
  assert.equal(delayForRetryAfter("-5"), 0);
  assert.equal(delayForRetryAfter("0"), 0);
});

test("absent Retry-After header clears any pending rate-limit delay", () => {
  const client = new GraphClient();
  (client as any).rateLimitDelay = 42_000;
  (client as any).updateRateLimitInfo({ headers: {} });
  assert.equal((client as any).rateLimitDelay, 0);
});
