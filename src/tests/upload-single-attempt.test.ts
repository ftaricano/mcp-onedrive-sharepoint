import test from "node:test";
import assert from "node:assert/strict";
import { createServer, Agent } from "node:https";
import { createServer as createHttpServer } from "node:http";
import { connect, type ConnectionOptions } from "node:tls";
import { execFileSync } from "node:child_process";
import { mkdtemp, writeFile, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { AxiosInstance } from "axios";
import {
  GraphClient,
  __setGraphClientInstanceForTests,
  type UploadOptions,
} from "../graph/client.js";
import {
  __setAuthInstanceForTests,
  type MicrosoftGraphAuth,
} from "../auth/microsoft-graph-auth.js";
import { getToolRegistry } from "../tools/registry.js";
import { buildArgs, parseArgs } from "../cli/args.js";

const endpoint = "/drives/acme-drive/root:/report.txt:/content";
const graphUrl = `https://graph.microsoft.com/v1.0${endpoint}`;
const limit = 4194304;
const single = { retryPolicy: "single-attempt", retries: 2, timeout: 80 } as UploadOptions;

type Mode = number | "success" | "lost-response" | "timeout" | "reset" | "EPIPE" | "session" | "missing-folder";

// The real Axios HTTP adapter retains the validated Graph URL. A test-owned
// HTTPS agent connects only to the loopback TLS fixture and trusts its generated
// certificate for graph.microsoft.com. No adapter replacement or DNS/Graph I/O.
async function fixture(t: test.TestContext, mode: Mode = "success") {
  const directory = await mkdtemp(join(tmpdir(), "acme-upload-"));
  const file = join(directory, "report.txt");
  await writeFile(file, "fictional report");
  const keyPath = join(directory, "fixture.key");
  const certPath = join(directory, "fixture.crt");
  execFileSync("openssl", ["req", "-x509", "-newkey", "rsa:2048", "-nodes",
    "-keyout", keyPath, "-out", certPath, "-days", "1", "-subj", "/CN=graph.microsoft.com",
    "-addext", "subjectAltName=DNS:graph.microsoft.com"], { stdio: "ignore" });
  const key = await readFile(keyPath);
  const cert = await readFile(certPath);
  let fixturePort = 0;
  class LoopbackAgent extends Agent {
    createConnection(options: ConnectionOptions) {
      return connect({ ...options, host: "127.0.0.1", port: fixturePort,
        servername: "graph.microsoft.com", ca: cert });
    }
  }
  const agent = new LoopbackAgent({ keepAlive: false, ca: cert });
  const requests: Array<{ method: string; url: string; bytes: number }> = [];
  let applied = 0;
  let injected = false;
  let chunkBytes = 0;
  const chunkServer = createHttpServer((req, res) => {
    const chunks: Buffer[] = [];
    req.on("data", (chunk: Buffer) => chunks.push(chunk));
    req.on("end", () => {
      const bytes = Buffer.concat(chunks).length;
      requests.push({ method: req.method || "", url: req.url || "", bytes });
      applied++;
      chunkBytes += bytes;
      res.writeHead(chunkBytes < limit ? 202 : 200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ id: "acme-file", name: "report.txt", size: chunkBytes }));
    });
  });
  await new Promise<void>((resolve) => chunkServer.listen(0, "127.0.0.1", resolve));
  const chunkAddress = chunkServer.address();
  assert.ok(chunkAddress && typeof chunkAddress !== "string");
  const chunkUrl = `http://127.0.0.1:${chunkAddress.port}/chunk`;
  const oldNoProxy = process.env.NO_PROXY;
  process.env.NO_PROXY = `${oldNoProxy || ""},127.0.0.1`;
  const server = createServer({ key, cert }, (req, res) => {
    const chunks: Buffer[] = [];
    req.on("data", (chunk: Buffer) => chunks.push(chunk));
    req.on("end", () => {
      const method = req.method || "";
      const url = req.url || "";
      const bytes = Buffer.concat(chunks).length;
      requests.push({ method, url, bytes });
      assert.equal(req.headers.host, "graph.microsoft.com");
      assert.ok(url.startsWith("/v1.0/"));
      if (method === "GET") {
        if (mode === "missing-folder") {
          res.writeHead(404, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ error: { code: "itemNotFound", message: "fictional missing folder" } }));
          return;
        }
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ id: "folder", folder: {} }));
        return;
      }
      applied++;
      // Subsequent requests succeed: on the unfixed base this exposes a replay
      // instead of stalling on repeated RetryHelper backoffs.
      if (applied === 1) {
        if (mode === "lost-response" || mode === "reset") {
          injected = true;
          req.socket.destroy();
          return;
        }
        if (mode === "timeout") {
          injected = true;
          return;
        }
        if (mode === "EPIPE") {
          // Deterministic local EPIPE fault on the real client socket after the
          // fixture accepted the body; not an adapter mock or OS timing race.
          injected = true;
          for (const sockets of Object.values(agent.sockets)) {
            for (const socket of sockets || []) {
              socket.destroy(Object.assign(new Error("fictional broken pipe"), { code: "EPIPE" }));
            }
          }
          return;
        }
        if (typeof mode === "number") {
          res.writeHead(mode, mode < 400 ? { Location: graphUrl + "?redirected=1" } : {});
          res.end(JSON.stringify({ error: { code: "fictionalFailure", message: "fictional error" } }));
          return;
        }
      }
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify(mode === "session" && method === "POST"
        ? { uploadUrl: chunkUrl }
        : { id: "acme-file", name: "report.txt", size: bytes }));
    });
  });
  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  assert.ok(address && typeof address !== "string");
  __setAuthInstanceForTests({ getAccessToken: async () => "fictional-token" } as unknown as MicrosoftGraphAuth);
  const client = new GraphClient();
  const instance = (client as unknown as { axios: AxiosInstance }).axios;
  fixturePort = address.port;
  instance.defaults.proxy = false;
  instance.defaults.httpsAgent = agent;
  __setGraphClientInstanceForTests(client);
  const oldRoot = process.env.MCP_LOCAL_FILE_ROOT;
  process.env.MCP_LOCAL_FILE_ROOT = directory;
  t.after(async () => {
    t.diagnostic(JSON.stringify({ mode, applied, injected, requests }));
    await client.cleanup();
    __setGraphClientInstanceForTests(null);
    __setAuthInstanceForTests(null);
    if (oldRoot === undefined) delete process.env.MCP_LOCAL_FILE_ROOT;
    else process.env.MCP_LOCAL_FILE_ROOT = oldRoot;
    if (oldNoProxy === undefined) delete process.env.NO_PROXY;
    else process.env.NO_PROXY = oldNoProxy;
    agent.destroy();
    chunkServer.closeAllConnections();
    await new Promise<void>((resolve) => chunkServer.close(() => resolve()));
    server.closeAllConnections();
    await new Promise<void>((resolve) => server.close(() => resolve()));
    await rm(directory, { recursive: true, force: true });
  });
  return { client, instance, file, requests, get applied() { return applied; }, get injected() { return injected; } };
}

function payload(result: { content: Array<{ text: string }> }) {
  return JSON.parse(result.content[0].text) as { success: boolean };
}

for (const via of ["cli", "mcp"] as const) {
  test(`schema and shared ${via} handler forward single-attempt to real HTTP`, async (t) => {
    const f = await fixture(t, 503);
    const registry = getToolRegistry();
    const schema = registry.tools.find((tool) => tool.name === "upload_file")!.inputSchema;
    assert.deepEqual(schema.properties?.retryPolicy, {
      type: "string", enum: ["default", "single-attempt"], default: "default",
      description: "Upload retry policy; single-attempt requires a file smaller than 4 MiB and an existing destination folder",
    });
    const args = via === "cli"
      ? buildArgs(parseArgs([`--localPath=${f.file}`, "--remotePath=existing/report.txt", "--driveId=acme-drive", "--retryPolicy=single-attempt"]))
      : { localPath: f.file, remotePath: "existing/report.txt", driveId: "acme-drive", retryPolicy: "single-attempt" };
    const result = await registry.handlers.upload_file(args);
    assert.equal(result.isError, true);
    assert.equal(f.applied, 1);
    assert.deepEqual(f.requests.map((r) => r.method), ["PUT"]);
  });
}

test("single-attempt success sends exactly one complete body", async (t) => {
  const f = await fixture(t);
  assert.equal((await f.client.uploadFile(endpoint, f.file, undefined, single)).success, true);
  assert.equal(f.applied, 1);
  assert.equal(f.requests[0].bytes, Buffer.byteLength("fictional report"));
});

for (const mode of ["lost-response", "timeout", "reset", "EPIPE", 400, 401, 403, 404, 408, 409, 429, 500, 502, 503, 504, 507] as const) {
  test(`single-attempt returns ${mode} without replay or reconciliation request`, async (t) => {
    const f = await fixture(t, mode);
    await assert.rejects(f.client.uploadFile(endpoint, f.file, undefined, single), (error: unknown) => {
      assert.ok(error instanceof Error && "code" in error);
      t.diagnostic(JSON.stringify({ errorCode: error.code }));
      if (typeof mode === "string") {
        assert.equal(error.code, mode === "timeout" ? "ECONNABORTED" : mode === "EPIPE" ? "EPIPE" : "ECONNRESET");
      }
      return true;
    });
    assert.equal(f.applied, 1);
    assert.equal(f.requests.length, 1);
    if (typeof mode === "string") assert.equal(f.injected, true);
  });
}

for (const status of [301, 302, 303, 307, 308]) {
  test(`single-attempt refuses HTTP ${status}; default actually follows it`, async (t) => {
    const f = await fixture(t, status);
    await assert.rejects(f.client.uploadFile(endpoint, f.file, undefined, single));
    assert.equal(f.requests.length, 1);
    assert.equal(f.applied, 1);
    assert.ok(f.requests[0].bytes > 0);
  });
  test(`default redirect positive control HTTP ${status}`, async (t) => {
    const f = await fixture(t, status);
    assert.equal((await f.client.uploadFile(endpoint, f.file)).success, true);
    assert.equal(f.requests.length, 2);
    assert.equal(f.requests[1].method, status === 303 ? "GET" : "PUT");
    assert.equal(f.requests[1].bytes, status === 303 ? 0 : f.requests[0].bytes);
  });
}

for (const size of [limit - 1, limit, limit + 1]) {
  for (const via of ["programmatic", "handler"] as const) {
    test(`single-attempt ${via} size ${size}: boundary before any Graph write/session`, async (t) => {
      const f = await fixture(t);
      await writeFile(f.file, Buffer.alloc(size, 65));
      if (via === "programmatic") {
        const run = f.client.uploadFile(endpoint, f.file, undefined, single);
        if (size < limit) assert.equal((await run).success, true);
        else await assert.rejects(run, /single-attempt/);
      } else {
        const result = await getToolRegistry().handlers.upload_file({ localPath: f.file, remotePath: "missing/report.txt", driveId: "acme-drive", retryPolicy: "single-attempt" });
        if (size < limit) assert.equal(payload(result).success, true);
        else assert.equal(result.isError, true);
      }
      assert.equal(f.applied, size < limit ? 1 : 0);
      assert.equal(f.requests.length, size < limit ? 1 : 0);
      if (size < limit) assert.equal(f.requests[0].bytes, size);
    });
  }
}

for (const policy of ["unknown", "", null, false, true, 0, 1, {}, []]) {
  test(`invalid retryPolicy ${JSON.stringify(policy)} fails before HTTP in both consumers`, async (t) => {
    const f = await fixture(t);
    await assert.rejects(f.client.uploadFile(endpoint, f.file, undefined, { retryPolicy: policy } as UploadOptions), /retryPolicy/);
    const args = buildArgs(parseArgs([`--json=${JSON.stringify({ localPath: f.file, remotePath: "missing/report.txt", retryPolicy: policy })}`]));
    const result = await getToolRegistry().handlers.upload_file(args);
    assert.equal(result.isError, true);
    assert.equal(f.requests.length, 0);
    assert.equal(f.applied, 0);
  });
}

for (const retryPolicy of [undefined, "default"] as const) {
  test(`default compatibility: ${retryPolicy} retains retry after uncertain write`, async (t) => {
    const f = await fixture(t, "lost-response");
    const result = await f.client.uploadFile(endpoint, f.file, undefined, { retryPolicy, retries: 2 } as UploadOptions);
    assert.equal(result.success, true);
    assert.equal(f.applied, 2);
  });
  test(`default compatibility: ${retryPolicy} retains large-file session/chunks`, async (t) => {
    const f = await fixture(t, "session");
    await writeFile(f.file, Buffer.alloc(limit, 65));
    assert.equal((await f.client.uploadFile(endpoint, f.file, undefined, { retryPolicy } as UploadOptions)).success, true);
    assert.deepEqual(f.requests.map((r) => r.method), ["POST", "PUT", "PUT"]);
    assert.ok(f.requests[0].url.endsWith("/createUploadSession"));
    assert.equal(f.requests[1].bytes + f.requests[2].bytes, limit);
  });
}

for (const status of [404, 409, 429, 503, 307]) {
  test(`shared handler HTTP ${status}: no folder/session POST, replay or fallback`, async (t) => {
    const f = await fixture(t, status);
    const result = await getToolRegistry().handlers.upload_file({
      localPath: f.file, remotePath: "missing/report.txt", driveId: "acme-drive", retryPolicy: "single-attempt",
    });
    assert.equal(result.isError, true);
    assert.deepEqual(f.requests.map((r) => r.method), ["PUT"]);
    assert.equal(f.applied, 1);
  });
}

test("default shared handler still prepares and creates missing folders", async (t) => {
  const f = await fixture(t, "missing-folder");
  const result = await getToolRegistry().handlers.upload_file({
    localPath: f.file, remotePath: "missing/report.txt", driveId: "acme-drive",
  });
  assert.equal(payload(result).success, true);
  assert.deepEqual(f.requests.map((r) => r.method), ["GET", "GET", "POST", "PUT"]);
  assert.equal(f.applied, 2);
});

test("single-attempt rejects error statuses despite caller validateStatus/retries", async (t) => {
  const f = await fixture(t, 503);
  await assert.rejects(f.client.uploadFile(endpoint, f.file, undefined, {
    ...single, retries: 0, validateStatus: () => true,
  }));
  assert.equal(f.applied, 1);
});
