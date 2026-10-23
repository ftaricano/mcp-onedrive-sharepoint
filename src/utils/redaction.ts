/** Redact output only; never pass these values back into Graph requests. */
const sensitiveKey =
  /^(?:@microsoft\.graph\.downloadurl|downloadurl|tempauth|token|access_token|authorization|proxy-authorization)$/i;
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

function redactString(value: string): string {
  return value
    .replace(
      /(["'](?:@microsoft\.graph\.downloadurl|downloadurl|tempauth|access_token|token|authorization|proxy-authorization)["']\s*:\s*)(["'])(.*?)\2/gi,
      (_match, prefix, quote, entry: string) =>
        `${prefix}${quote}${/^https?:\/\//i.test(entry) ? safeUrl(entry) : "***"}${quote}`,
    )
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
