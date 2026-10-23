import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { redactOutput, redactSensitive } from "../utils/redaction.js";
import { jsonTextResponse, toolErrorResponse } from "../graph/contracts.js";
import { AuditLogger, SecurityValidator } from "../utils/security-validator.js";
import { createUserFriendlyError } from "../graph/error-handler.js";

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

function escapedDiagnostic(
  fields: Record<string, string>,
  layers: number,
): string {
  let body = JSON.stringify(fields);
  for (let layer = 0; layer < layers; layer++) {
    body = JSON.stringify({ body });
  }
  return `diagnostic ${body}`;
}

for (const layers of [1, 2]) {
  for (const key of [
    "access_token",
    "refresh_token",
    "id_token",
    "Authorization",
    "downloadUrl",
    "@microsoft.graph.downloadUrl",
    "tempauth",
  ]) {
    test(`redacts ${key} escaped in ${layers} JSON layers at every output boundary`, () => {
      const diagnostic = escapedDiagnostic(
        {
          [key]: canary,
          code: "InvalidAuthenticationToken",
          name: "report.txt",
        },
        layers,
      );
      const error = new Error(diagnostic);
      const envelope = jsonTextResponse({ diagnostic });
      const outputs = [
        redactOutput(diagnostic),
        JSON.stringify(redactSensitive({ diagnostic })),
        envelope.content[0].text,
        redactOutput(envelope.content[0].text),
        JSON.stringify(redactSensitive(error)),
        JSON.stringify(toolErrorResponse("upload_file", error)),
        createUserFriendlyError(error),
        SecurityValidator.sanitizeForLogging(diagnostic),
        SecurityValidator.sanitizeForLogging({ diagnostic }),
      ];
      for (const output of outputs) {
        assert.ok(!output.includes(canary));
        assert.ok(output.includes("***"));
        assert.ok(output.includes("InvalidAuthenticationToken"));
        assert.ok(output.includes("report.txt"));
      }
      // Every wrapper remains valid JSON with the original number of layers.
      let decoded = redactOutput(diagnostic).slice("diagnostic ".length);
      for (let layer = 0; layer < layers; layer++)
        decoded = JSON.parse(decoded).body;
      assert.equal(JSON.parse(decoded)[key], "***");
      assert.ok(diagnostic.includes(canary));
    });
  }

  for (const [label, value] of [
    ["bearer", `Bearer ${canary}==`],
    ["header", `Authorization: Basic ${canary}`],
    ["signed-url", `${base}?sig=${canary}`],
  ]) {
    test(`redacts escaped ${label} in ${layers} JSON layers and actual audit logs`, (t) => {
      const diagnostic = escapedDiagnostic(
        { message: value, code: "NotFound" },
        layers,
      );
      const consoleLog = t.mock.method(console, "log", () => {});
      AuditLogger.log("upload", "fictional-user", diagnostic, "failure", {
        diagnostic,
      });
      const outputs = [
        redactOutput(diagnostic),
        JSON.stringify(toolErrorResponse("upload_file", new Error(diagnostic))),
        JSON.stringify(AuditLogger.getRecentLogs(1)),
        JSON.stringify(consoleLog.mock.calls),
      ];
      for (const output of outputs) {
        assert.ok(!output.includes(canary));
        assert.ok(output.includes("NotFound"));
      }
      let decoded = redactOutput(diagnostic).slice("diagnostic ".length);
      for (let layer = 0; layer < layers; layer++)
        decoded = JSON.parse(decoded).body;
      assert.equal(JSON.parse(decoded).code, "NotFound");
      if (label === "signed-url") {
        assert.ok(redactOutput(diagnostic).includes(base));
      }
    });
  }

  test(`ods stderr redacts diagnostics with ${layers} escaped JSON layers`, () => {
    const diagnostic = escapedDiagnostic(
      { refresh_token: canary, id_token: canary },
      layers,
    );
    const child = spawnSync(
      process.execPath,
      [fileURLToPath(new URL("../cli.js", import.meta.url)), diagnostic],
      {
        encoding: "utf8",
        env: { MCP_TOOL_PROFILE: "core" },
      },
    );
    assert.equal(child.status, 1);
    assert.equal(child.stdout, "");
    assert.ok(!child.stderr.includes(canary));
    assert.ok(child.stderr.includes("***"));
  });
}

test("ordinary escaped JSON metadata is preserved verbatim", () => {
  const diagnostic = escapedDiagnostic(
    {
      name: "report.txt",
      code: "InvalidAuthenticationToken",
      eTag: '"version-1"',
      cTag: '"content-1"',
    },
    2,
  );
  assert.equal(redactOutput(diagnostic), diagnostic);
});

test("excessive escaped nesting fails closed", () => {
  assert.ok(
    !redactOutput(escapedDiagnostic({ refresh_token: canary }, 10)).includes(
      canary,
    ),
  );
});
