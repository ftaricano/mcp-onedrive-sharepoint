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
    const cachedUser = await auth.getCurrentUser();
    if (!cachedUser) {
      throw new Error(
        "Authentication required. Run `npm run setup-auth` before calling tools.",
      );
    }

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
  } catch {
    // Config missing or auth not yet primed — surfaces on first tool call.
  }
}
