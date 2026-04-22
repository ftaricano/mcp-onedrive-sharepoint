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

export function buildDriveSearchEndpoint(
  target: DriveTarget,
  query: string,
): string {
  const base = getDriveRootEndpoint(target);
  return `${base}/root/search(q='${encodeURIComponent(query)}')`;
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
