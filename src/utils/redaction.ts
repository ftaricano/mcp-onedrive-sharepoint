/** Redact output only; never pass these values back into Graph requests. */
// Keep token field names identical across objects and embedded JSON diagnostics.
// Do not include "code": ordinary Graph error codes must remain readable.
const tokenKeyPattern = String.raw`[\w-]*token[\w-]*`;
const sensitiveKeyPattern = String.raw`(?:@microsoft\.graph\.downloadurl|downloadurl|tempauth|${tokenKeyPattern}|authorization|proxy-authorization)`;
const sensitiveKey = new RegExp(`^${sensitiveKeyPattern}$`, "i");
const sensitiveTextField = new RegExp(
  String.raw`(["']${sensitiveKeyPattern}["']\s*:\s*)(["'])(.*?)\2`,
  "gi",
);
const signedQuery = /[?&](?:tempauth|sig|se|[^=&]*token[^=&]*|code)=/i;

function isSignedUrl(value: string): boolean {
  if (signedQuery.test(value)) return true;
  try {
    return [...new URL(value).searchParams.keys()].some((key) =>
      /^(?:tempauth|sig|se|code)$|token/i.test(key),
    );
  } catch {
    return false;
  }
}

function safeUrl(value: string): string {
  try {
    const url = new URL(value);
    return `${url.protocol}//${url.host}${url.pathname}?***`;
  } catch {
    return "***";
  }
}

function redactPlainText(value: string): string {
  return value
    .replace(/https?:\/\/[^\s"'<>\\]+/gi, (url) =>
      isSignedUrl(url) ? safeUrl(url) : url,
    )
    .replace(
      /\b(Authorization|Proxy-Authorization)\s*:\s*[^\r\n]+/gi,
      "$1: ***",
    )
    .replace(/\bBearer\s+[^\s"'<>\\,;]+/gi, "Bearer ***")
    .replace(
      /\b(tempauth|sig|se|[\w-]*token[\w-]*|code)\s*=\s*[^\s&"'<>\\,;]+/gi,
      "$1=***",
    );
}

// Decode JSON string literals inside diagnostics before matching sensitive fields.
// Re-encode only changed literals, preserving prefixes and ordinary metadata.
const maxJsonStringLayers = 8;

function redactString(value: string, depth = 0): string {
  // Bound hostile nesting and fail closed instead of returning uninspected text.
  if (depth >= maxJsonStringLayers) return "***";

  const fieldsMasked = value.replace(
    sensitiveTextField,
    (_match, prefix, quote, entry: string) =>
      `${prefix}${quote}${/^https?:\/\//i.test(entry) ? safeUrl(entry) : "***"}${quote}`,
  );
  // Plain header matching must not consume the rest of an encoded JSON document.
  return fieldsMasked
    .split(/("(?:\\.|[^"\\])*")/g)
    .map((part, index) => {
      if (index % 2 === 0) return redactPlainText(part);
      let decoded: string;
      try {
        decoded = JSON.parse(part) as string;
      } catch {
        return "***";
      }
      const redacted = redactString(decoded, depth + 1);
      return redacted === decoded ? part : JSON.stringify(redacted);
    })
    .join("");
}

export function redactSensitive(value: unknown): unknown {
  if (typeof value === "string") return redactString(value);
  // Do not serialize HTTP client internals, which may contain request credentials.
  if (value instanceof Error) {
    return { name: value.name, message: redactString(value.message) };
  }
  if (Array.isArray(value)) return value.map(redactSensitive);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [
        key,
        sensitiveKey.test(key)
          ? typeof entry === "string" && /^https?:\/\//i.test(entry)
            ? safeUrl(entry)
            : "***"
          : redactSensitive(entry),
      ]),
    );
  }
  return value;
}

/** Covers both plain diagnostics and JSON text already serialized by a handler. */
export function redactOutput(text: string): string {
  try {
    const parsed: unknown = JSON.parse(text);
    return JSON.stringify(redactSensitive(parsed), null, 2);
  } catch {
    return redactString(text);
  }
}
