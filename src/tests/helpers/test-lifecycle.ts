import { after, afterEach } from "node:test";

import { __setGraphClientInstanceForTests } from "../../graph/client.js";
import { cleanupAllCaches } from "../../utils/cache-manager.js";

export function registerGraphClientTestLifecycle(): void {
  afterEach(() => {
    __setGraphClientInstanceForTests(null);
    cleanupAllCaches();
  });

  after(() => {
    cleanupAllCaches();
  });
}
