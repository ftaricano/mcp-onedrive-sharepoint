export function coerce(value: string | undefined): unknown {
  if (value === undefined) return true;
  if (value === "true") return true;
  if (value === "false") return false;
  if (value === "null") return null;
  if (/^-?\d+$/.test(value)) return Number(value);
  if (/^-?\d+\.\d+$/.test(value)) return Number(value);
  return value;
}

export function parseArgs(argv: string[]): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i];
    if (!token.startsWith("--")) continue;

    const eqIdx = token.indexOf("=");
    let key: string;
    let rawValue: string | undefined;

    if (eqIdx !== -1) {
      key = token.slice(2, eqIdx);
      rawValue = token.slice(eqIdx + 1);
    } else {
      key = token.slice(2);
      const next = argv[i + 1];
      if (next !== undefined && !next.startsWith("--")) {
        rawValue = next;
        i++;
      } else {
        rawValue = "true";
      }
    }

    out[key] = coerce(rawValue);
  }
  return out;
}

export function buildArgs(
  parsed: Record<string, unknown>,
): Record<string, unknown> {
  const { json, ...rest } = parsed;
  if (typeof json === "string") {
    let payload: unknown;
    try {
      payload = JSON.parse(json);
    } catch (err) {
      throw new Error(`Invalid --json payload: ${(err as Error).message}`);
    }
    if (payload === null || typeof payload !== "object" || Array.isArray(payload)) {
      throw new Error("--json payload must be a JSON object");
    }
    return { ...(payload as Record<string, unknown>), ...rest };
  }
  return rest;
}

export function extractText(result: unknown): string {
  if (!result || typeof result !== "object") {
    return JSON.stringify(result, null, 2);
  }
  const r = result as { content?: unknown };
  if (Array.isArray(r.content)) {
    return r.content
      .map((c: any) =>
        typeof c?.text === "string" ? c.text : JSON.stringify(c),
      )
      .join("\n");
  }
  return JSON.stringify(result, null, 2);
}
