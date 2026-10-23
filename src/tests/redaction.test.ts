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
    url: "https://example.com/documents/report.xlsx?view=1",
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

for (const header of ["Authorization", "Proxy-Authorization"]) {
  for (const scheme of ["Basic", "Bearer"]) {
    for (const layers of [0, 1, 2]) {
      test(`redacts quoted ${header} ${scheme} through ${layers} JSON layers`, (t) => {
        const message = `${header}: "${scheme} ${canary}"`;
        const diagnostic =
          layers === 0
            ? message
            : escapedDiagnostic(
                { message, code: "NotFound", name: "report.txt" },
                layers,
              );
        const envelope = jsonTextResponse({ diagnostic, code: "NotFound" });
        const consoleLog = t.mock.method(console, "log", () => {});
        AuditLogger.log("upload", "fictional-user", diagnostic, "failure", {
          diagnostic,
        });
        const outputs = [
          redactOutput(diagnostic),
          envelope.content[0].text,
          redactOutput(envelope.content[0].text), // CLI's second pass
          JSON.stringify(
            toolErrorResponse("upload_file", new Error(diagnostic)),
          ),
          createUserFriendlyError(new Error(diagnostic)),
          SecurityValidator.sanitizeForLogging(diagnostic),
          SecurityValidator.sanitizeForLogging({ diagnostic }),
          JSON.stringify(AuditLogger.getRecentLogs(1)),
          JSON.stringify(consoleLog.mock.calls),
        ];
        for (const output of outputs) {
          assert.ok(!output.includes(canary));
          assert.ok(output.includes("***"));
        }
        assert.equal(JSON.parse(envelope.content[0].text).code, "NotFound");
        if (layers > 0) {
          let decoded = redactOutput(diagnostic).slice("diagnostic ".length);
          for (let layer = 0; layer < layers; layer++)
            decoded = JSON.parse(decoded).body;
          const fields = JSON.parse(decoded);
          assert.equal(fields.code, "NotFound");
          assert.equal(fields.name, "report.txt");
          assert.ok(!fields.message.includes(canary));
        }
        const child = spawnSync(
          process.execPath,
          [fileURLToPath(new URL("../cli.js", import.meta.url)), diagnostic],
          { encoding: "utf8", env: { MCP_TOOL_PROFILE: "core" } },
        );
        assert.equal(child.status, 1);
        assert.equal(child.stdout, "");
        assert.ok(!child.stderr.includes(canary));
        assert.ok(child.stderr.includes("***"));
      });
    }
    test(`quoted ${header} ${scheme} preserves adjacent JSON fields`, () => {
      const diagnostic = `diagnostic {${header}: "${scheme} ${canary}", "code":"NotFound", "name":"report.txt"}`;
      const output = redactOutput(diagnostic);
      assert.ok(!output.includes(canary));
      assert.ok(output.includes('"code":"NotFound"'));
      assert.ok(output.includes('"name":"report.txt"'));
    });
  }
}

// Full field policy × syntax × representation × case matrix. No real credentials.
const matrixKeys = [
  "access_token",
  "refresh_token",
  "id_token",
  "client_secret",
  "password",
  "token",
  "sig",
  "se",
  "code",
  "tempauth",
  "downloadUrl",
  "@microsoft.graph.downloadUrl",
  "Authorization",
  "Proxy-Authorization",
  "x-ms-access-token",
  "x-ms-custom-token",
];
const assignmentForms: Array<[string, (key: string) => string]> = [
  ["equals", (key) => `${key}=${canary}`],
  ["colon", (key) => `${key}: ${canary}`],
  ["json", (key) => JSON.stringify({ [key]: canary })],
  ["single-quotes", (key) => `{'${key}':'${canary}'}`],
  ["quoted-equals", (key) => `${key}="${canary}"`],
  ["single-quoted-equals", (key) => `${key}='${canary}'`],
  ["bearer", (key) => `${key}: Bearer ${canary}`],
  ["basic", (key) => `${key}: Basic ${canary}`],
  ["other-scheme", (key) => `${key}: FictionalScheme ${canary}`],
  ["query", (key) => `${base}?view=public&${key}=${canary}`],
];
function encodeDiagnostic(value: string, layers: number, raw: boolean): string {
  for (let layer = 0; layer < layers; layer++) {
    value = JSON.stringify(value);
    if (raw) value = value.slice(1, -1);
  }
  return value;
}
const representations: Array<[string, number, boolean]> = [
  ["plain", 0, false],
  ["json-1", 1, false],
  ["json-2", 2, false],
  ["raw-escape-1", 1, true],
  ["raw-escape-2", 2, true],
];
for (const field of matrixKeys) {
  for (const key of [field, field.toUpperCase()]) {
    test(`matrix structured field ${key}`, () => {
      const input = { nested: [{ [key]: canary }], name: "report.txt" };
      const output = JSON.parse(jsonTextResponse(input).content[0].text);
      assert.equal(output.nested[0][key], "***");
      assert.equal(output.name, "report.txt");
      assert.equal(input.nested[0][key], canary);
    });
    for (const [syntax, form] of assignmentForms) {
      for (const [representation, layers, raw] of representations) {
        test(`matrix ${key} / ${syntax} / ${representation}`, () => {
          const diagnostic = `diagnostic ${encodeDiagnostic(form(key), layers, raw)}`;
          const envelope = jsonTextResponse({
            diagnostic,
            eTag: '"v1"',
            cTag: '"c1"',
          });
          const outputs = [
            redactOutput(diagnostic),
            JSON.stringify(redactSensitive({ diagnostic })),
            envelope.content[0].text,
            redactOutput(envelope.content[0].text),
            JSON.stringify(
              toolErrorResponse("upload_file", new Error(diagnostic)),
            ),
            SecurityValidator.sanitizeForLogging(diagnostic),
            SecurityValidator.sanitizeForLogging({ diagnostic }),
          ];
          for (const output of outputs) {
            assert.ok(!output.includes(canary));
            assert.ok(output.includes("***"));
          }
          const parsed = JSON.parse(envelope.content[0].text);
          assert.equal(parsed.eTag, '"v1"');
          assert.equal(parsed.cTag, '"c1"');
          if (syntax === "query") assert.ok(outputs[0].includes(base));
        });
      }
    }
  }
}

for (const layers of [0, 1, 2]) {
  test(`public metadata remains readable through ${layers} layers`, () => {
    const fields = {
      name: "report.txt",
      code: "InvalidAuthenticationToken",
      eTag: '"v1"',
      cTag: '"c1"',
      url: base,
    };
    const text = `diagnostic ${encodeDiagnostic(JSON.stringify(fields), layers, false)}`;
    assert.equal(redactOutput(text), text);
    const output = redactSensitive(fields);
    assert.deepEqual(output, fields);
  });
  test(`JWT and unknown SharePoint signing key through ${layers} layers`, () => {
    const fictionalJwt = "ZmFrZS1oZWFkZXI.ZmFrZS1wYXlsb2Fk.ZmFrZS1zaWduYXR1cmU";
    for (const input of [
      fictionalJwt,
      `unrecognized=${fictionalJwt}`,
      `${base}?opaque=${canary}#${canary}`,
      `https://example.com/file?opaque=${fictionalJwt}`,
      `https://example.com/${fictionalJwt}`,
      `downloadUrl=https://example.com/${fictionalJwt}?opaque=${canary}`,
    ]) {
      const text = `diagnostic ${encodeDiagnostic(input, layers, false)}`;
      for (const output of [
        redactOutput(text),
        jsonTextResponse({ text }).content[0].text,
        JSON.stringify(toolErrorResponse("upload_file", new Error(text))),
        SecurityValidator.sanitizeForLogging(text),
      ]) {
        assert.ok(!output.includes(canary));
        assert.ok(!output.includes(fictionalJwt));
        assert.ok(output.includes("***"));
        if (input.startsWith(base)) assert.ok(output.includes(base));
      }
    }
  });
}

test("known error codes stay public only in diagnostic fields", () => {
  for (const code of [
    "NotFound",
    "InvalidAuthenticationToken",
    "accessDenied",
    "ECONNRESET",
    "UnknownError",
  ]) {
    const input = { code, eTag: '"v1"' };
    assert.deepEqual(redactSensitive(input), input);
    assert.ok(
      redactOutput(`diagnostic ${JSON.stringify(input)}`).includes(code),
    );
    assert.ok(!redactOutput(`code=${code}`).includes(code));
    assert.ok(!redactOutput(`${base}?code=${code}`).includes(code));
  }
});

test("generic assignments preserve adjacent fields and quoted delimiters", () => {
  for (const key of matrixKeys) {
    for (const value of [
      canary,
      `${canary}, extra; & detail`,
      `prefix "${canary}" suffix`,
      `prefix '${canary}' suffix`,
    ]) {
      const text = `diagnostic {${JSON.stringify(key)}:${JSON.stringify(value)},"name":"report.txt","eTag":"v1","cTag":"c1","code":"NotFound"}`;
      const output = redactOutput(text).slice("diagnostic ".length);
      const parsed = JSON.parse(output);
      const isHeader = /^(?:(?:proxy-)?authorization|www-authenticate)$/i.test(
        key,
      );
      assert.equal(
        parsed[key],
        key === "code"
          ? "NotFound"
          : isHeader && value.startsWith("prefix ")
            ? "prefix ***"
            : "***",
      );
      assert.equal(parsed.name, "report.txt");
      assert.equal(parsed.eTag, "v1");
      assert.equal(parsed.cTag, "c1");
      assert.equal(parsed.code, "NotFound");
      assert.ok(!output.includes(canary));
    }
    const isHeader = /^(?:(?:proxy-)?authorization|www-authenticate)$/i.test(
      key,
    );
    const output = redactOutput(
      isHeader
        ? JSON.stringify({ [key]: canary, diagnostic: "eTag=v1 cTag=c1" })
        : `${key}=${canary} eTag=v1 cTag=c1`,
    );
    assert.ok(!output.includes(canary));
    assert.ok(output.includes("eTag=v1 cTag=c1"));
  }
});

test("CLI masks generic quoted assignments before bootstrap", () => {
  for (const key of [
    "access_token",
    "client_secret",
    "password",
    "x-ms-access-token",
  ]) {
    const diagnostic = `${key}="${canary}"`;
    const child = spawnSync(
      process.execPath,
      [fileURLToPath(new URL("../cli.js", import.meta.url)), diagnostic],
      { encoding: "utf8", env: { MCP_TOOL_PROFILE: "core" } },
    );
    assert.equal(child.status, 1);
    assert.equal(child.stdout, "");
    assert.ok(!child.stderr.includes(canary));
    assert.ok(child.stderr.includes("***"));
  }
});

test("compound scalar values keep quotes bound to their field", () => {
  for (const key of matrixKeys) {
    for (const value of [
      `Basic "${canary}"`,
      `FictionalScheme '${canary}'`,
      `prefix"${canary}"suffix`,
      `prefix'${canary}'suffix`,
    ]) {
      const diagnostic =
        /^(?:(?:proxy-)?authorization|www-authenticate)$/i.test(key)
          ? JSON.stringify({ [key]: value, diagnostic: "eTag=v1" })
          : `${key}=${value} eTag=v1`;
      const output = redactOutput(diagnostic);
      assert.ok(!output.includes(canary));
      assert.ok(output.includes("eTag=v1"));
    }
  }
});

test("field/value association survives whitespace around either separator", () => {
  for (const key of matrixKeys) {
    for (const delimiter of ["=", ":"]) {
      const input = `diagnostic ${JSON.stringify(key)} \r\n ${delimiter} \r\n "${canary}"`;
      assert.ok(!redactOutput(input).includes(canary));
      assert.ok(!jsonTextResponse({ input }).content[0].text.includes(canary));
    }
  }
});

const authenticationHeaders = [
  "Authorization",
  "Proxy-Authorization",
  "WWW-Authenticate",
];
const authenticationSchemes = [
  "Bearer",
  "Basic",
  "Digest",
  "NTLM",
  "Negotiate",
  "AWS4-HMAC-SHA256",
  "FictionalScheme",
];
const parameterCanaries = [
  "fictional-username-canary",
  "fictional-response-canary",
  "fictional-nonce-canary",
  "fictional-credential-canary",
  "fictional-signature-canary",
  "fictional-extra-canary",
];
// Every scheme gets parameters, including unknown names and mixed delimiters:
// no parameter-name allowlist can accidentally make this matrix pass.
function parameterizedHeader(scheme: string): string {
  return `${scheme} username="${parameterCanaries[0]}", response="${parameterCanaries[1]}", nonce="${parameterCanaries[2]}", Credential=${parameterCanaries[3]}, Signature=${parameterCanaries[4]}; opaque=${parameterCanaries[5]}`;
}
for (const header of authenticationHeaders) {
  for (const key of [header, header.toLowerCase()]) {
    for (const scheme of authenticationSchemes) {
      const value = parameterizedHeader(scheme);
      const line = `${key}: ${value}`;
      const formats: Array<[string, string]> = [
        ["line", line],
        [
          "json-field",
          JSON.stringify({
            [key]: value,
            code: "NotFound",
            name: "report.txt",
          }),
        ],
        [
          "json-message",
          `diagnostic ${JSON.stringify({ message: line, code: "NotFound", name: "report.txt" })}`,
        ],
        ["serialized-1", encodeDiagnostic(line, 1, false)],
        ["serialized-2", encodeDiagnostic(line, 2, false)],
        ["raw-escape-1", encodeDiagnostic(line, 1, true)],
        ["raw-escape-2", encodeDiagnostic(line, 2, true)],
      ];
      for (const [format, text] of formats) {
        test(`authentication header ${key} / ${scheme} / ${format}`, (t) => {
          const diagnostic =
            format === "json-field" ? text : `diagnostic ${text}`;
          const envelope = jsonTextResponse({
            diagnostic,
            eTag: '"v1"',
            cTag: '"c1"',
            code: "NotFound",
          });
          const consoleLog = t.mock.method(console, "log", () => {});
          AuditLogger.log("upload", "fictional-user", diagnostic, "failure", {
            diagnostic,
          });
          const outputs = [
            redactOutput(diagnostic),
            envelope.content[0].text,
            redactOutput(envelope.content[0].text),
            JSON.stringify(
              toolErrorResponse("upload_file", new Error(diagnostic)),
            ),
            createUserFriendlyError(new Error(diagnostic)),
            SecurityValidator.sanitizeForLogging(diagnostic),
            SecurityValidator.sanitizeForLogging({ diagnostic }),
            JSON.stringify(AuditLogger.getRecentLogs(1)),
            JSON.stringify(consoleLog.mock.calls),
          ];
          for (const output of outputs) {
            for (const sentinel of parameterCanaries)
              assert.ok(!output.includes(sentinel));
            assert.ok(output.includes(`${scheme} ***`));
            for (const parameter of [
              "username=",
              "response=",
              "nonce=",
              "Credential=",
              "Signature=",
              "opaque=",
            ]) {
              assert.ok(!output.includes(parameter));
            }
          }
          const parsed = JSON.parse(envelope.content[0].text);
          assert.equal(parsed.eTag, '"v1"');
          assert.equal(parsed.cTag, '"c1"');
          assert.equal(parsed.code, "NotFound");
          if (format === "json-field") {
            const fields = JSON.parse(outputs[0]);
            assert.equal(fields[key], `${scheme} ***`);
            assert.equal(fields.name, "report.txt");
            assert.equal(fields.code, "NotFound");
          }
          assert.equal(redactOutput(line), `${key}: ${scheme} ***`);
          assert.equal(
            redactOutput(`${key}: ${scheme} ***`),
            `${key}: ${scheme} ***`,
          );
          const structured = JSON.parse(
            jsonTextResponse({ [key]: value, eTag: '"v1"' }).content[0].text,
          );
          assert.equal(structured[key], `${scheme} ***`);
          assert.equal(structured.eTag, '"v1"');
        });
      }
      test(`CLI header ${key} / ${scheme} without bootstrap`, () => {
        const child = spawnSync(
          process.execPath,
          [fileURLToPath(new URL("../cli.js", import.meta.url)), line],
          { encoding: "utf8", env: { MCP_TOOL_PROFILE: "core" } },
        );
        assert.equal(child.status, 1);
        assert.equal(child.stdout, "");
        for (const sentinel of parameterCanaries)
          assert.ok(!child.stderr.includes(sentinel));
        assert.ok(child.stderr.includes(`${scheme} ***`));
      });
    }
  }
}

test("authentication line consumes every parameter through EOL, preserving following lines", () => {
  for (const header of authenticationHeaders) {
    const text = `${header}: ${parameterizedHeader("Digest")}\r\nname: report.txt`;
    const output = redactOutput(text);
    assert.equal(output, `${header}: Digest ***\r\nname: report.txt`);
  }
});

test("plain authentication lines never resume after a quoted component", () => {
  for (const header of authenticationHeaders) {
    for (const text of [
      `${header}: "Digest" username="${parameterCanaries[0]}", response="${parameterCanaries[1]}"`,
      `${header}: 'Digest' nonce='${parameterCanaries[2]}'`,
      `${header}: "Basic ${parameterCanaries[0]}" extra=${parameterCanaries[1]}`,
    ]) {
      const output = redactOutput(text);
      for (const sentinel of parameterCanaries)
        assert.ok(!output.includes(sentinel));
      assert.ok(output.includes("***"));
    }
  }
});

const standaloneValueForms: Array<[string, string]> = [
  ["bare", canary],
  ["double-quoted", JSON.stringify(canary)],
  ["single-quoted", `'${canary}'`],
  ["parameters", `username="${canary}", nonce='${canary}', response=${canary}`],
];
for (const originalScheme of authenticationSchemes) {
  for (const scheme of [originalScheme, originalScheme.toLowerCase()]) {
    for (const prefix of [
      "",
      ...authenticationHeaders.map((header) => `${header}: `),
    ]) {
      for (const [form, credential] of standaloneValueForms) {
        for (const [representation, layers, raw] of representations) {
          test(`standalone scheme ${scheme} / ${prefix || "none"} / ${form} / ${representation}`, (t) => {
            const line = `${prefix}${scheme} ${credential}`;
            const diagnostic = `diagnostic ${encodeDiagnostic(line, layers, raw)}`;
            const envelope = jsonTextResponse({
              diagnostic,
              eTag: '"v1"',
              cTag: '"c1"',
              code: "NotFound",
              url: base,
            });
            const consoleLog = t.mock.method(console, "log", () => {});
            AuditLogger.log("upload", "fictional-user", diagnostic, "failure", {
              diagnostic,
            });
            for (const output of [
              redactOutput(line),
              redactOutput(diagnostic),
              envelope.content[0].text,
              redactOutput(envelope.content[0].text),
              JSON.stringify(
                toolErrorResponse("upload_file", new Error(diagnostic)),
              ),
              createUserFriendlyError(new Error(diagnostic)),
              SecurityValidator.sanitizeForLogging(diagnostic),
              SecurityValidator.sanitizeForLogging({ diagnostic }),
              JSON.stringify(AuditLogger.getRecentLogs(1)),
              JSON.stringify(consoleLog.mock.calls),
            ]) {
              assert.ok(!output.includes(canary));
              assert.ok(output.includes(`${scheme} ***`));
            }
            const fields = JSON.parse(envelope.content[0].text);
            assert.equal(fields.eTag, '"v1"');
            assert.equal(fields.cTag, '"c1"');
            assert.equal(fields.code, "NotFound");
            assert.equal(fields.url, base);
            assert.equal(redactOutput(line), `${prefix}${scheme} ***`);
          });
        }
      }
    }
  }
  for (const [form, credential] of standaloneValueForms.slice(0, 3)) {
    test(`CLI standalone ${originalScheme} / ${form} without bootstrap`, () => {
      const child = spawnSync(
        process.execPath,
        [
          fileURLToPath(new URL("../cli.js", import.meta.url)),
          `${originalScheme} ${credential}`,
        ],
        { encoding: "utf8", env: { MCP_TOOL_PROFILE: "core" } },
      );
      assert.equal(child.status, 1);
      assert.equal(child.stdout, "");
      assert.ok(!child.stderr.includes(canary));
      assert.ok(child.stderr.includes(`${originalScheme} ***`));
    });
  }
}

test("standalone authentication preserves outer JSON metadata and public diagnostics", () => {
  const input = {
    message: `Bearer "${canary}"`,
    code: "InvalidAuthenticationToken",
    eTag: '"v1"',
    cTag: '"c1"',
    url: base,
  };
  for (const layers of [0, 1, 2]) {
    const encoded = encodeDiagnostic(JSON.stringify(input), layers, false);
    const output = redactOutput(`diagnostic ${encoded}`);
    assert.ok(!output.includes(canary));
    assert.ok(output.includes("InvalidAuthenticationToken"));
    assert.ok(output.includes(base));
    assert.ok(output.includes("Bearer ***"));
  }
  assert.equal(redactOutput("code: NotFound"), "code: NotFound");
  assert.equal(
    redactOutput("Authentication Error: expired"),
    "Authentication Error: expired",
  );
});

// Round 8: a credential belongs to its scheme, irrespective of following prose.
const proseSchemes = [...authenticationSchemes, "Token"];
const proseCredentialForms: Array<[string, string]> = [
  ["bare", canary],
  ["padded-token68", `${canary}==`],
  ["double-quoted", JSON.stringify(canary)],
  ["single-quoted", `'${canary}'`],
];
for (const scheme of proseSchemes) {
  for (const [form, credential] of proseCredentialForms) {
    for (const suffix of [" expired", ", expired", "! expired"]) {
      for (const [representation, layers, raw] of representations) {
        test(`credential before prose ${scheme} / ${form} / ${suffix} / ${representation}`, (t) => {
          const line = `${scheme} ${credential}${suffix}`;
          const metadata = {
            message: line,
            name: "Monthly Report.xlsx",
            code: "InvalidAuthenticationToken",
            eTag: '"v1"',
            cTag: '"c1"',
            url: base,
          };
          const diagnostic = `diagnostic ${encodeDiagnostic(JSON.stringify(metadata), layers, raw)}`;
          const envelope = jsonTextResponse({ ...metadata, diagnostic });
          const consoleLog = t.mock.method(console, "log", () => {});
          AuditLogger.log("upload", "fictional-user", diagnostic, "failure", {
            diagnostic,
          });
          for (const output of [
            redactOutput(line),
            redactOutput(diagnostic),
            envelope.content[0].text,
            redactOutput(envelope.content[0].text),
            JSON.stringify(
              toolErrorResponse("upload_file", new Error(diagnostic)),
            ),
            createUserFriendlyError(new Error(diagnostic)),
            SecurityValidator.sanitizeForLogging(diagnostic),
            SecurityValidator.sanitizeForLogging({ diagnostic }),
            JSON.stringify(AuditLogger.getRecentLogs(1)),
            JSON.stringify(consoleLog.mock.calls),
          ]) {
            assert.ok(!output.includes(canary));
            assert.ok(output.includes(`${scheme} ***${suffix}`));
          }
          const fields = JSON.parse(envelope.content[0].text);
          assert.deepEqual(fields, {
            ...metadata,
            message: `${scheme} ***${suffix}`,
            diagnostic: redactOutput(diagnostic),
          });
          assert.equal(redactOutput(line), `${scheme} ***${suffix}`);
          assert.equal(
            redactOutput(redactOutput(line)),
            `${scheme} ***${suffix}`,
          );
          // Decode preserved envelopes, proving commas/quotes/braces still delimit metadata.
          let recovered = redactOutput(diagnostic).slice("diagnostic ".length);
          for (let layer = 0; layer < layers; layer++) {
            recovered = JSON.parse(raw ? '"' + recovered + '"' : recovered);
          }
          assert.deepEqual(JSON.parse(recovered), {
            ...metadata,
            message: `${scheme} ***${suffix}`,
          });
        });
      }
    }
  }
  for (const [representation, layers, raw] of representations) {
    test(`CLI credential before prose ${scheme} / ${representation}`, () => {
      const diagnostic = encodeDiagnostic(
        `${scheme} ${canary} expired`,
        layers,
        raw,
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
      assert.ok(child.stderr.includes(`${scheme} *** expired`));
    });
  }
}

test("review example and successive authentication credentials retain their suffixes", () => {
  const diagnostic = "Bearer fictional-sensitive-canary expired";
  assert.equal(redactOutput(diagnostic), "Bearer *** expired");
  assert.equal(
    redactOutput(
      `Bearer ${canary} expired; Basic '${canary}' invalid; refresh_token=fictional-followup-canary`,
    ),
    "Bearer *** expired; Basic *** invalid; refresh_token=***",
  );
});

test("unknown schemes require authentication headers; ordinary multiword metadata survives", () => {
  const ordinary = {
    name: "Monthly Report.xlsx",
    message: "Monthly Report.xlsx expired",
    code: "NotFound",
    eTag: '"v1"',
    cTag: '"c1"',
    url: base,
  };
  assert.deepEqual(redactSensitive(ordinary), ordinary);
  assert.deepEqual(
    JSON.parse(redactOutput(JSON.stringify(ordinary))),
    ordinary,
  );
  for (const header of authenticationHeaders) {
    for (const [name, layers, raw] of representations) {
      const input = {
        ...ordinary,
        [header]: `UnregisteredAuth ${canary} expired`,
      };
      const text = encodeDiagnostic(JSON.stringify(input), layers, raw);
      const output = redactOutput(text);
      assert.ok(!output.includes(canary), name);
      assert.ok(output.includes("UnregisteredAuth ***"), name);
      assert.ok(output.includes(ordinary.name), name);
    }
  }
});
