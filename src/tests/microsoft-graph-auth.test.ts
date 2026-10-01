import assert from "node:assert/strict";
import test from "node:test";

import { MicrosoftGraphAuth } from "../auth/microsoft-graph-auth.js";

test("fails high when no client secret is in the environment", async () => {
  const auth = new MicrosoftGraphAuth({
    clientId: "11111111-1111-4111-8111-111111111111",
    tenantId: "22222222-2222-4222-822222222222",
  });

  await assert.rejects(
    auth.getAccessToken(),
    /Missing Microsoft Graph client secret.*MICROSOFT_GRAPH_CLIENT_SECRET/i,
  );
});

test("does not offer delegated authentication that would persist a token", async () => {
  const auth = new MicrosoftGraphAuth({
    clientId: "11111111-1111-4111-8111-111111111111",
    clientSecret: "test-secret",
  });

  await assert.rejects(auth.authenticate(), /does not persist tokens/i);
});

test("signOut only clears the in-memory session state", async () => {
  const auth = new MicrosoftGraphAuth({
    clientId: "11111111-1111-4111-8111-111111111111",
    clientSecret: "test-secret",
  });

  await auth.signOut();
  assert.equal(await auth.getCurrentUser(), null);
});
