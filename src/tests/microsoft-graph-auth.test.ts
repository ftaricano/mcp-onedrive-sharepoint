import assert from "node:assert/strict";
import test from "node:test";

import { MicrosoftGraphAuth } from "../auth/microsoft-graph-auth.js";

test("fails high when the 1Password launcher did not inject a client secret", async () => {
  const auth = new MicrosoftGraphAuth({
    clientId: "11111111-1111-4111-8111-111111111111",
    tenantId: "22222222-2222-4222-822222222222",
  });

  await assert.rejects(
    auth.getAccessToken(),
    /1Password launcher.*SP_CLIENT_SECRET/i,
  );
});

test("does not offer delegated authentication that would persist a token", async () => {
  const auth = new MicrosoftGraphAuth({
    clientId: "11111111-1111-4111-8111-111111111111",
    clientSecret: "test-secret",
  });

  await assert.rejects(auth.authenticate(), /cannot persist tokens outside 1Password/i);
});

test("signOut only clears the in-memory session state", async () => {
  const auth = new MicrosoftGraphAuth({
    clientId: "11111111-1111-4111-8111-111111111111",
    clientSecret: "test-secret",
  });

  await auth.signOut();
  assert.equal(await auth.getCurrentUser(), null);
});
