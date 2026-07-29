import { loadConfig } from "../config/index.js";
import {
  getAuthInstance,
  initializeAuth,
} from "../auth/microsoft-graph-auth.js";
import { getGraphClient } from "../graph/client.js";

let initialized = false;
let initPromise: Promise<void> | null = null;

export async function bootstrap(): Promise<void> {
  if (initialized) return;
  if (initPromise) return initPromise;

  initPromise = (async () => {
    const config = loadConfig();
    initializeAuth(config.auth);

    const auth = getAuthInstance();
    await auth.getAccessToken();

    getGraphClient();
    initialized = true;
  })();

  try {
    await initPromise;
  } catch (err) {
    initPromise = null;
    throw err;
  }
}

export function prewarmAuth(): void {
  try {
    const config = loadConfig();
    initializeAuth(config.auth);
    const auth = getAuthInstance() as unknown as {
      prewarm?: () => void;
    };
    if (typeof auth.prewarm === "function") auth.prewarm();
  } catch (err) {
    // Non-fatal: bootstrap() will raise the real failure on first tool call.
    // Still log so missing config / broken 1Password resolution does not vanish.
    const message = err instanceof Error ? err.message : String(err);
    console.error(`[prewarmAuth] skipped: ${message}`);
  }
}
