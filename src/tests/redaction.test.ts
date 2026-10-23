import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { redactOutput, redactSensitive } from "../utils/redaction.js";
import { jsonTextResponse, toolErrorResponse } from "../graph/contracts.js";
import { SecurityValidator } from "../utils/security-validator.js";

const canary = "fictional-sensitive-canary";
const base = "https://contoso.sharepoint.com/documents/report.xlsx";
for (const key of [
  "tempauth",
  "sig",
  "se",
  "token",
  "access_token",
  "code",
  "refresh_token",
  "%73ig",
]) {
  test(`removes signed query ${key} from every output boundary`, () => {
    const input = `${base}?public=also-private&${key}=${canary}#fragment-private`;
    for (const output of [
      redactOutput(input),
      jsonTextResponse({ nested: [{ url: input }] }).content[0].text,
      SecurityValidator.sanitizeForLogging({ url: input }),
      JSON.stringify(toolErrorResponse("upload_file", new Error(input))),
    ]) {
      assert.ok(!output.includes(canary));
      assert.ok(!output.includes("also-private"));
      assert.ok(!output.includes("fragment-private"));
      assert.ok(output.includes(base));
      assert.ok(output.includes("***"));
    }
  });
}
for (const key of [
  "@microsoft.graph.downloadUrl",
  "downloadUrl",
  "tempauth",
  "Authorization",
  "authorization",
  "Proxy-Authorization",
]) {
  test(`redacts sensitive field ${key} before JSON serialization`, () => {
    const output = jsonTextResponse({ nested: [{ [key]: canary }] }).content[0]
      .text;
    assert.ok(!output.includes(canary));
    assert.equal(JSON.parse(output).nested[0][key], "***");
    const text = `diagnostic ${JSON.stringify({ [key]: canary })}`;
    assert.ok(!redactOutput(text).includes(canary));
  });
}
// Match the token field family already supported by assignment diagnostics.
for (const key of [
  "token",
  "access_token",
  "refresh_token",
  "id_token",
  "custom_token",
  "sessionToken",
  "refresh-token",
  "REFRESH_TOKEN",
]) {
  test(`redacts token field ${key} across structured and textual boundaries`, () => {
    const input = {
      nested: [{ [key]: canary }],
      code: "InvalidAuthenticationToken",
    };
    const structured = jsonTextResponse(input).content[0].text;
    const serialized = redactOutput(JSON.stringify(input));
    const diagnostic = redactOutput(`diagnostic ${JSON.stringify(input)}`);
    const error = toolErrorResponse("upload_file", {
      error: {
        code: "InvalidAuthenticationToken",
        message: `diagnostic ${JSON.stringify(input)}`,
      },
    });

    assert.equal(JSON.parse(structured).nested[0][key], "***");
    assert.equal(JSON.parse(serialized).nested[0][key], "***");
    for (const output of [
      structured,
      serialized,
      diagnostic,
      JSON.stringify(error),
    ]) {
      assert.ok(!output.includes(canary));
      assert.ok(output.includes("***"));
      assert.ok(output.includes("InvalidAuthenticationToken"));
    }
    assert.equal(
      JSON.parse(error.content[1].text).error.code,
      "InvalidAuthenticationToken",
    );
    assert.equal(input.nested[0][key], canary);
  });
}
for (const input of [
  `Authorization: Basic ${canary}`,
  `Authorization: Bearer ${canary}==`,
  `Bearer ${canary}==`,
  `tempauth=${canary}`,
  `access_token=${canary}`,
  `code=${canary}`,
]) {
  test(`redacts diagnostic ${input.split(/[ :=]/)[0]}`, () => {
    assert.ok(!redactOutput(input).includes(canary));
  });
}
test("download URL fields retain only host and path even for unknown query names", () => {
  const input = {
    "@microsoft.graph.downloadUrl": `${base}?opaque=${canary}#private`,
    eTag: '"version-1"',
    cTag: '"content-1"',
  };
  const output = redactSensitive(input) as typeof input;
  assert.equal(output["@microsoft.graph.downloadUrl"], `${base}?***`);
  assert.equal(output.eTag, input.eTag);
  assert.equal(output.cTag, input.cTag);
  assert.equal(
    input["@microsoft.graph.downloadUrl"],
    `${base}?opaque=${canary}#private`,
  );
  assert.equal(
    redactOutput(JSON.stringify(input)),
    JSON.stringify(output, null, 2),
  );
});
test("redaction is idempotent and keeps ordinary metadata and unsigned URLs", () => {
  const input = {
    url: `${base}?view=1`,
    size: 10,
    code: "NotFound",
    eTag: '"v1"',
    cTag: '"c1"',
  };
  assert.deepEqual(redactSensitive(input), input);
  const output = redactOutput(`${base}?sig=${canary}`);
  assert.equal(redactOutput(output), output);
});
test("does not serialize HTTP request internals from errors", () => {
  const error = Object.assign(new Error(`Bearer ${canary}==`), {
    config: { headers: { Authorization: canary } },
  });
  assert.ok(!JSON.stringify(redactSensitive(error)).includes(canary));
});

// Unknown commands stop before bootstrap: no identity, credentials or Graph call.
test("ods stderr masks sensitive command text before writing", () => {
  const cli = fileURLToPath(new URL("../cli.js", import.meta.url));
  for (const input of [
    `${base}?sig=${canary}`,
    `Authorization: Bearer ${canary}==`,
  ]) {
    const child = spawnSync(process.execPath, [cli, input], {
      encoding: "utf8",
      env: { MCP_TOOL_PROFILE: "core" },
    });
    assert.equal(child.status, 1);
    assert.equal(child.stdout, "");
    assert.ok(!child.stderr.includes(canary));
    assert.ok(child.stderr.includes("***"));
    if (input.startsWith("https:")) assert.ok(child.stderr.includes(base));
  }
});
