/**
 * Numeric parsing helpers with safe bounds.
 */

/**
 * Parse a value into a positive integer, guarding against NaN and runaway
 * magnitudes. Used for untrusted/optional numeric inputs such as the
 * `Retry-After` HTTP header and numeric environment configuration.
 *
 * Returns `fallback` when the value is missing, non-numeric, zero or negative.
 * Clamps any value above `max` down to `max` so a hostile/huge input cannot
 * make the process hang on an unbounded timer.
 */
export function parsePositiveInt(
  value: string | number | undefined | null,
  fallback: number,
  max: number,
): number {
  const parsed =
    typeof value === "number" ? Math.trunc(value) : parseInt(value ?? "", 10);

  if (!Number.isFinite(parsed) || parsed <= 0) {
    return fallback;
  }

  return Math.min(parsed, max);
}
