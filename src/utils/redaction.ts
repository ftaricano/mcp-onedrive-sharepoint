/** Redact output only; never pass these values back into Graph requests. */
// One field policy for structured objects and all textual assignment syntaxes.
const sensitiveKey =
  /^(?:[\w-]*token[\w-]*|client_secret|password|sig|se|code|tempauth|(?:@microsoft\.graph\.)?downloadurl|(?:proxy-)?authorization|www-authenticate)$/i;
const authenticationHeader = /^(?:(?:proxy-)?authorization|www-authenticate)$/i;
// Only known diagnostic codes are public. An arbitrary value in "code" is secret.
// Query parameters and = assignments never use this exception.
const publicErrorCodes = new Set(
  `InvalidAuthenticationToken AuthenticationFailure ExpiredAuthenticationToken
MalformedAuthenticationToken InvalidTokenType TokenNotFound Unauthenticated Forbidden
InsufficientScope AccessDenied Unauthorized InsufficientPrivileges InvalidPermission ConsentRequired
UserNotInTenant NotFound ResourceNotFound ItemNotFound FileNotFound DriveItemNotFound SiteNotFound
ListNotFound WorkbookNotFound WorksheetNotFound TooManyRequests RateLimitExceeded RequestThrottled
ThrottledRequest ServiceUnavailable TemporarilyUnavailable InsufficientStorage QuotaExceeded
StorageQuotaExceeded DiskQuotaExceeded FileSizeLimitExceeded ItemCountLimitExceeded BadRequest
InvalidRequest InvalidParameter MalformedRequest InvalidRange InvalidFilter InvalidOrderBy
ValidationError InvalidFileName InvalidPath UnsupportedMediaType Conflict ResourceModified
ConcurrencyFailure NameAlreadyExists ItemAlreadyExists DuplicateName VersionConflict
InternalServerError BadGateway GatewayTimeout ServiceError TemporaryFailure UnexpectedError
NetworkError ConnectionError TimeoutError DNSError ConnectionRefused ConnectionTimeout ECONNRESET
ECONNREFUSED ECONNABORTED ETIMEDOUT EAI_AGAIN ENOTFOUND EHOSTUNREACH ENETUNREACH EPIPE UnknownError
generalException invalidRequest itemNotFound notSupported accessDenied unauthenticated
activityLimitReached nameAlreadyExists resourceModified resyncRequired ErrorAccessDenied
ErrorInvalidIdMalformed ErrorItemNotFound ErrorInvalidRequest ErrorInvalidUser UserNotFound
Request_ResourceNotFound Authorization_RequestDenied invalid_client invalid_grant`
    .toLowerCase()
    .split(/\s+/),
);
const jwt =
  /(?<![\w.-])[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+(?![\w.-])/g;
const maxJsonStringLayers = 8;

function safeUrl(value: string): string {
  try {
    const url = new URL(value);
    return `${url.protocol}//${url.host}${url.pathname.replace(jwt, "***")}?***`;
  } catch {
    return "***";
  }
}

function redactUrl(value: string): string {
  try {
    const url = new URL(value);
    // Unknown SharePoint signing parameters are private too: redact its whole query.
    const sharepointQuery =
      /(?:^|\.)sharepoint(?:-df|-mil)?\.(?:com|us|de|cn)$/i.test(
        url.hostname,
      ) &&
      (url.search !== "" || url.hash !== "");
    const signed = [...url.searchParams].some(
      ([key, entry]) => sensitiveKey.test(key) || entry.match(jwt) !== null,
    );
    if (sharepointQuery || signed) return safeUrl(value);
    // Keep the URL authority readable; a hostname is not a JWT.
    return value.replace(
      /^(https?:\/\/[^/?#]+)(.*)$/i,
      (_match, authority: string, rest: string) =>
        authority + rest.replace(jwt, "***"),
    );
  } catch {
    return "***";
  }
}

function maskAuthenticationHeader(value: unknown): string {
  if (typeof value !== "string") return "***";
  // The first HTTP token is the scheme; every parameter/credential is private.
  // A lone opaque value has no independently identifiable scheme: fail closed.
  const scheme = /^([!#$%&'*+.^_`|~0-9A-Za-z-]+)[ \t]+\S/.exec(value.trim());
  return scheme ? `${scheme[1].replace(jwt, "***")} ***` : "***";
}

function readAuthenticationHeader(
  text: string,
  start: number,
  jsonField: boolean,
): string {
  const tail = text.slice(start);
  // JSON/shell quotes bound the complete value, not individual parameters.
  if (jsonField && (tail[0] === '"' || tail[0] === "'"))
    return readValue(text, start);
  // On a diagnostic header line, commas and k=v pairs are part of the header.
  return /^[^\r\n]*/.exec(tail)![0];
}

function maskField(
  key: string,
  value: unknown,
  diagnosticCode: boolean,
): unknown {
  if (!sensitiveKey.test(key)) return value;
  if (authenticationHeader.test(key)) return maskAuthenticationHeader(value);
  if (
    diagnosticCode &&
    key.toLowerCase() === "code" &&
    typeof value === "string" &&
    publicErrorCodes.has(value.toLowerCase())
  )
    return value;
  return typeof value === "string" && /^https?:\/\//i.test(value)
    ? safeUrl(value)
    : "***";
}

// Quoted tokens are indivisible: field identity and its value stay together.
// This same lexer handles JSON, shell-like assignments, queries and headers.
const lexemePattern =
  /https?:\/\/[^\s"'<>\\]+|"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|[@\w.-]+/gi;

function decodeQuoted(value: string): string {
  return value[0] === '"'
    ? (JSON.parse(value) as string)
    : value.slice(1, -1).replace(/\\(['\\])/g, "$1");
}

function encodeQuoted(value: string, quote: string): string {
  if (quote === '"') return JSON.stringify(value);
  return (
    "'" +
    value
      .replace(/\\/g, "\\\\")
      .replace(/'/g, "\\'")
      .replace(/\r/g, "\\r")
      .replace(/\n/g, "\\n") +
    "'"
  );
}

function readValue(text: string, start: number): string {
  const tail = text.slice(start);
  if (tail[0] === '"' || tail[0] === "'") {
    const quoted = /^(?:"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*')/.exec(tail);
    // An incomplete literal consumes the remainder and fails closed.
    return quoted?.[0] ?? tail;
  }
  // Read a compound scalar as one value, including scheme + quoted credential.
  // Delimiters inside quotes do not terminate it; adjacent assignments do.
  let offset = 0;
  while (offset < tail.length) {
    const rest = tail.slice(offset);
    if (/^[\r\n,;&<>}\]]/.test(rest)) break;
    if (
      /^\s+(?:[@\w.-]+|"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*')\s*[:=]/.test(rest)
    )
      break;
    if (rest[0] === '"' || rest[0] === "'") {
      const quoted = /^(?:"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*')/.exec(rest);
      if (!quoted) return tail;
      offset += quoted[0].length;
    } else {
      offset++;
    }
  }
  return tail.slice(0, offset).trimEnd();
}

/** Locate an authentication production, independent of the scheme name. */
function standaloneAuthenticationEnd(
  text: string,
  start: number,
  depth: number,
): number | null {
  const whitespace = /^[ \t]+/.exec(text.slice(start));
  if (!whitespace) return null;
  const valueStart = start + whitespace[0].length;
  const tail = text.slice(valueStart);
  if (tail[0] === '"' || tail[0] === "'") {
    const raw = readValue(text, valueStart);
    let decoded: string;
    try {
      decoded = decodeQuoted(raw);
    } catch {
      return valueStart + /^[^\r\n]*/.exec(tail)![0].length;
    }
    // A serialized diagnostic is an envelope, not an opaque credential.
    // Recurse into it before interpreting its leading label as an auth scheme.
    if (
      /^[\[{]/.test(decoded.trim()) ||
      redactString(decoded, depth + 1) !== decoded
    )
      return null;
    return valueStart + /^[^\r\n]*/.exec(tail)![0].length;
  }
  // RFC token68 credentials or a parameter list. Do not split parameters by name.
  const parameters = /^[!#$%&'*+.^_`|~0-9A-Za-z-]+[ \t]*=/.test(tail);
  if (parameters) return valueStart + /^[^\r\n]*/.exec(tail)![0].length;
  const opaque = /^[A-Za-z0-9._~+\/-]+={0,}(?=[ \t]*(?:$|[\r\n,;}\]]))/.exec(
    tail,
  );
  // Public diagnostic error codes are not opaque authentication credentials.
  if (!opaque || publicErrorCodes.has(opaque[0].toLowerCase())) return null;
  return valueStart + opaque[0].length;
}

function redactString(value: string, depth = 0): string {
  if (depth >= maxJsonStringLayers) return "***";
  // A raw escaped fragment need not have its own surrounding JSON quotes.
  // Decode one complete representation before lexing; preserve its escape level.
  if (value.includes('\\"')) {
    try {
      const decoded = JSON.parse('"' + value + '"') as string;
      if (decoded !== value) {
        const redacted = redactString(decoded, depth + 1);
        return redacted === decoded
          ? value
          : JSON.stringify(redacted).slice(1, -1);
      }
    } catch {
      /* Ordinary JSON literals are decoded individually below. */
    }
  }
  const lexer = new RegExp(lexemePattern);
  let result = "";
  let end = 0;
  let token: RegExpExecArray | null;
  while ((token = lexer.exec(value)) !== null) {
    const part = token[0];
    const start = token.index;
    result += value.slice(end, start);
    if (/^https?:\/\//i.test(part)) {
      result += redactUrl(part);
      end = lexer.lastIndex;
      continue;
    }
    const quoted = part[0] === '"' || part[0] === "'";
    // Mixed malformed escape representations must never expose unparsed fragments.
    if (quoted && value[start - 1] === "\\") return "***";
    let decoded: string;
    try {
      decoded = quoted ? decodeQuoted(part) : part;
    } catch {
      return "***";
    }
    const separator = /^\s*[:=]\s*/.exec(value.slice(lexer.lastIndex));
    if (separator && sensitiveKey.test(decoded)) {
      const valueStart = lexer.lastIndex + separator[0].length;
      const raw = authenticationHeader.test(decoded)
        ? readAuthenticationHeader(
            value,
            valueStart,
            quoted || /[\[{,]\s*$/.test(value.slice(0, start)),
          )
        : readValue(value, valueStart);
      const quote = raw[0] === '"' || raw[0] === "'" ? raw[0] : "";
      let entry: string;
      try {
        entry = quote ? decodeQuoted(raw.trim()) : raw;
      } catch {
        entry = "***";
      }
      const masked = maskField(
        decoded,
        entry,
        separator[0].trim() === ":",
      ) as string;
      result +=
        part + separator[0] + (quote ? encodeQuoted(masked, quote) : masked);
      end = valueStart + raw.length;
      lexer.lastIndex = end;
      continue;
    }
    // All standalone schemes use the same value grammar, including quoted values.
    const authenticationEnd =
      !quoted &&
      value[start - 1] !== "=" &&
      /^[A-Za-z][!#$%&'*+.^_`|~0-9A-Za-z-]*$/.test(part)
        ? standaloneAuthenticationEnd(value, lexer.lastIndex, depth)
        : null;
    if (authenticationEnd !== null) {
      result += part + " ***";
      end = authenticationEnd;
      lexer.lastIndex = end;
      continue;
    }
    const redacted = quoted
      ? redactString(decoded, depth + 1)
      : part.replace(jwt, "***");
    result +=
      redacted === decoded
        ? part
        : quoted
          ? encodeQuoted(redacted, part[0])
          : redacted;
    end = lexer.lastIndex;
  }
  return result + value.slice(end);
}

export function redactSensitive(value: unknown): unknown {
  if (typeof value === "string") return redactString(value);
  // Never serialize HTTP client internals, which may contain request credentials.
  if (value instanceof Error)
    return { name: value.name, message: redactString(value.message) };
  if (Array.isArray(value)) return value.map(redactSensitive);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [
        key,
        sensitiveKey.test(key)
          ? maskField(key, entry, true)
          : redactSensitive(entry),
      ]),
    );
  }
  return value;
}

/** Covers plain diagnostics and JSON already serialized by a handler. */
export function redactOutput(text: string): string {
  try {
    return JSON.stringify(redactSensitive(JSON.parse(text)), null, 2);
  } catch {
    return redactString(text);
  }
}
