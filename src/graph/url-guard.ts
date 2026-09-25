/**
 * Destination guard for authenticated Microsoft Graph requests.
 *
 * Several tools accept a caller-supplied `pageToken` (a Graph `@odata.nextLink`)
 * and pass it to the Graph client as the request endpoint. axios ignores
 * `baseURL` when the URL is absolute, so without this guard an absolute value
 * would receive the app-only bearer token on any host. The client calls
 * `assertGraphRequestUrl` on the effective URL of every request, before a token
 * is fetched or attached.
 */
import { GraphApiError } from "./error-handler.js";

/** Hosts that may receive the Graph bearer token. */
export const ALLOWED_GRAPH_HOSTS: ReadonlySet<string> = new Set([
  "graph.microsoft.com",
]);

/** Graph API versions the client talks to. */
const ALLOWED_GRAPH_PATH_PREFIXES = ["/v1.0/", "/beta/"];

function rejectGraphRequestUrl(reason: string): never {
  throw new GraphApiError(
    {
      error: {
        code: "InvalidRequest",
        message: `Refusing to send an authenticated request: ${reason}. Pagination tokens must be Microsoft Graph nextLink values (https://graph.microsoft.com/v1.0 or /beta).`,
      },
    },
    "Request URL validation",
  );
}

/**
 * Throws unless `url` is an absolute `https://graph.microsoft.com/{v1.0|beta}/...`
 * URL on the default port with no userinfo. Returns the parsed URL.
 */
export function assertGraphRequestUrl(url: string): URL {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return rejectGraphRequestUrl("the request URL is not a valid URL");
  }

  if (parsed.protocol !== "https:") {
    rejectGraphRequestUrl("the request URL must use https");
  }

  if (parsed.username !== "" || parsed.password !== "") {
    rejectGraphRequestUrl("the request URL must not contain credentials");
  }

  if (!ALLOWED_GRAPH_HOSTS.has(parsed.hostname) || parsed.port !== "") {
    rejectGraphRequestUrl("the request URL host is not Microsoft Graph");
  }

  if (
    !ALLOWED_GRAPH_PATH_PREFIXES.some((prefix) =>
      parsed.pathname.startsWith(prefix),
    )
  ) {
    rejectGraphRequestUrl("the request URL has an unsupported Graph API path");
  }

  return parsed;
}
