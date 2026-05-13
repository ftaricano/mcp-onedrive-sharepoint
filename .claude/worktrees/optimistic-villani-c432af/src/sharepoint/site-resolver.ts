import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { Site } from "../graph/models.js";

export interface KnownSharePointSite {
  key: string;
  name: string;
  siteId: string;
  siteUrl: string;
  driveId?: string;
  aliases: string[];
}

export interface ResolvedSharePointSite {
  key?: string;
  name?: string;
  siteId: string;
  siteUrl?: string;
  driveId?: string;
  aliases?: string[];
  resolutionSource: "canonical_registry" | "graph_url_lookup" | "site_id";
  input: {
    site?: string;
    siteId?: string;
    siteUrl?: string;
  };
}

export interface SharePointSiteLookupInput {
  site?: string;
  siteId?: string;
  siteUrl?: string;
}

// Sites registry is loaded from an external JSON file so tenant-specific
// site/drive ids never live in the repo. Override the location with
// MCP_SITES_CONFIG_PATH; otherwise the resolver looks for
// <repo>/config/sites.local.json. If the file is missing, the registry stays
// empty and resolution falls back to explicit siteId/siteUrl/Graph lookup.
function defaultSitesConfigPath(): string {
  try {
    const moduleDir = dirname(fileURLToPath(import.meta.url));
    return resolve(moduleDir, "../../config/sites.local.json");
  } catch {
    return resolve(process.cwd(), "config/sites.local.json");
  }
}

function loadKnownSitesFromDisk(): KnownSharePointSite[] {
  const configPath = process.env.MCP_SITES_CONFIG_PATH || defaultSitesConfigPath();
  if (!existsSync(configPath)) {
    return [];
  }
  try {
    const raw = readFileSync(configPath, "utf-8");
    const parsed = JSON.parse(raw) as unknown;
    const sites = Array.isArray(parsed)
      ? parsed
      : (parsed as { sites?: unknown })?.sites;
    if (!Array.isArray(sites)) return [];
    return sites.filter(
      (s): s is KnownSharePointSite =>
        !!s &&
        typeof (s as KnownSharePointSite).siteId === "string" &&
        Array.isArray((s as KnownSharePointSite).aliases),
    );
  } catch {
    return [];
  }
}

let KNOWN_SHAREPOINT_SITES: KnownSharePointSite[] = loadKnownSitesFromDisk();

function normalizeAlias(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim()
    .replace(/^https?:\/\//, "")
    .replace(/\?.*$/, "")
    .replace(/#.*$/, "")
    .replace(/\/$/, "")
    .replace(/[^a-z0-9]+/g, "");
}

function normalizeUrl(value: string): string {
  try {
    const parsed = new URL(value);
    return `${parsed.protocol}//${parsed.host}${parsed.pathname}`
      .replace(/\/$/, "")
      .toLowerCase();
  } catch {
    return value.trim().replace(/\/$/, "").toLowerCase();
  }
}

function toResolvedSite(
  site: KnownSharePointSite,
  input: SharePointSiteLookupInput,
  resolutionSource: ResolvedSharePointSite["resolutionSource"],
): ResolvedSharePointSite {
  return {
    key: site.key,
    name: site.name,
    siteId: site.siteId,
    siteUrl: site.siteUrl,
    driveId: site.driveId,
    aliases: site.aliases,
    resolutionSource,
    input,
  };
}

function matchKnownSite(reference: string): KnownSharePointSite | undefined {
  const normalizedReference = normalizeAlias(reference);
  const normalizedUrlReference = normalizeUrl(reference);

  return KNOWN_SHAREPOINT_SITES.find((site) => {
    if (site.siteId.toLowerCase() === reference.trim().toLowerCase()) {
      return true;
    }

    if (normalizeUrl(site.siteUrl) === normalizedUrlReference) {
      return true;
    }

    if (
      normalizeAlias(site.name) === normalizedReference ||
      normalizeAlias(site.key) === normalizedReference
    ) {
      return true;
    }

    return site.aliases.some(
      (alias) => normalizeAlias(alias) === normalizedReference,
    );
  });
}

function buildGraphSiteLookupEndpoint(siteUrl: string): string {
  const parsed = new URL(siteUrl);
  return `/sites/${parsed.hostname}:${parsed.pathname.replace(/\/$/, "")}`;
}

export function getKnownSharePointSites(): KnownSharePointSite[] {
  return KNOWN_SHAREPOINT_SITES.map((site) => ({
    ...site,
    aliases: [...site.aliases],
  }));
}

export function __setKnownSitesForTests(sites: KnownSharePointSite[]): void {
  KNOWN_SHAREPOINT_SITES = sites.map((site) => ({
    ...site,
    aliases: [...site.aliases],
  }));
}

export function __resetKnownSitesForTests(): void {
  KNOWN_SHAREPOINT_SITES = loadKnownSitesFromDisk();
}

export async function resolveSharePointSiteReference(
  input: SharePointSiteLookupInput,
  client?: {
    get<T>(
      endpoint: string,
      params?: Record<string, string>,
    ): Promise<{ success: boolean; data?: T }>;
  },
): Promise<ResolvedSharePointSite | null> {
  const reference = input.siteId || input.siteUrl || input.site;
  if (!reference) {
    return null;
  }

  const knownSite = matchKnownSite(reference);
  if (knownSite) {
    return toResolvedSite(knownSite, input, "canonical_registry");
  }

  if (input.siteId) {
    return {
      siteId: input.siteId,
      resolutionSource: "site_id",
      input,
    };
  }

  const candidateUrl =
    input.siteUrl ||
    (input.site && /^https?:\/\//i.test(input.site) ? input.site : undefined);
  if (candidateUrl && client) {
    const response = await client.get<Site>(
      buildGraphSiteLookupEndpoint(candidateUrl),
    );
    if (response.success && response.data?.id) {
      return {
        siteId: response.data.id,
        siteUrl: response.data.webUrl || candidateUrl,
        name: response.data.displayName || response.data.name,
        resolutionSource: "graph_url_lookup",
        input,
      };
    }
  }

  return null;
}

export async function resolveRequiredSharePointSite(
  input: SharePointSiteLookupInput,
  client: {
    get<T>(
      endpoint: string,
      params?: Record<string, string>,
    ): Promise<{ success: boolean; data?: T }>;
  },
): Promise<ResolvedSharePointSite> {
  const resolved = await resolveSharePointSiteReference(input, client);
  if (!resolved?.siteId) {
    throw new Error(
      "A valid SharePoint site reference is required. Provide siteId, site alias, or canonical siteUrl.",
    );
  }

  return resolved;
}

export async function resolveDriveTargetContext(
  input: SharePointSiteLookupInput & { driveId?: string },
  client?: {
    get<T>(
      endpoint: string,
      params?: Record<string, string>,
    ): Promise<{ success: boolean; data?: T }>;
  },
): Promise<{
  siteId?: string;
  driveId?: string;
  resolvedSite?: ResolvedSharePointSite | null;
}> {
  const resolvedSite = await resolveSharePointSiteReference(input, client);
  const siteId = resolvedSite?.siteId ?? input.siteId;
  const driveId = input.driveId ?? resolvedSite?.driveId;

  if ((input.site || input.siteUrl) && !resolvedSite && !siteId && !driveId) {
    throw new Error(
      "Unable to resolve the requested SharePoint site. Refusing to fall back to personal OneDrive.",
    );
  }

  return {
    siteId,
    driveId,
    resolvedSite,
  };
}
