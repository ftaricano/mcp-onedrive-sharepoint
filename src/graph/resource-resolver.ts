export interface DriveTarget {
  driveId?: string;
  siteId?: string;
}

export interface DriveItemTarget extends DriveTarget {
  itemId?: string;
  itemPath?: string;
}

function normalizePathSegment(path?: string): string {
  if (!path) {
    return "";
  }

  return path.trim().replace(/^\/+/, "").replace(/\/+$/, "");
}

export function getDriveRootEndpoint(target: DriveTarget): string {
  if (target.driveId) {
    return `/drives/${target.driveId}`;
  }

  if (target.siteId) {
    return `/sites/${target.siteId}/drive`;
  }

  return "/me/drive";
}

export function buildDriveChildrenEndpoint(
  target: DriveTarget & { path?: string },
): string {
  const base = getDriveRootEndpoint(target);
  const normalizedPath = normalizePathSegment(target.path);

  if (!normalizedPath) {
    return `${base}/root/children`;
  }

  return `${base}/root:/${normalizedPath}:/children`;
}

export function buildDriveItemEndpoint(
  target: DriveItemTarget,
  suffix = "",
): string {
  const base = getDriveRootEndpoint(target);

  if (target.itemId) {
    return `${base}/items/${target.itemId}${suffix}`;
  }

  const normalizedPath = normalizePathSegment(target.itemPath);

  if (!normalizedPath) {
    throw new Error("Either itemId or itemPath must be provided");
  }

  return `${base}/root:/${normalizedPath}:${suffix}`;
}

/**
 * OData function parameters delimited by single quotes require the single
 * quote itself to be doubled to escape it (OData ABNF `quoted-string`).
 * Any caller constructing a `search(q='...')` / `range(address='...')` /
 * similar OData-quoted URL MUST pass the user-supplied value through this
 * first. `encodeURIComponent` does not escape `'` (it is RFC 3986 unreserved
 * when used as a sub-delim), so a literal quote would otherwise terminate
 * the OData string and allow query corruption.
 */
export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

export function buildDriveSearchEndpoint(
  target: DriveTarget,
  query: string,
): string {
  const base = getDriveRootEndpoint(target);
  const escaped = encodeURIComponent(escapeODataString(query));
  return `${base}/root/search(q='${escaped}')`;
}

export function describeDriveTarget(target: DriveTarget): string {
  if (target.driveId) {
    return `drive:${target.driveId}`;
  }

  if (target.siteId) {
    return `site:${target.siteId}`;
  }

  return "me";
}
