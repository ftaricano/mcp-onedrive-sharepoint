import { mkdtempSync, mkdirSync, rmSync, writeFileSync, utimesSync } from 'node:fs';
import { tmpdir } from 'node:os';
import * as path from 'node:path';

export type ToolEnvelope = {
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
};

type MockMethod = 'get' | 'post' | 'patch' | 'delete' | 'uploadFile' | 'downloadFile';

type MockHandlers = Partial<Record<MockMethod, (...args: any[]) => any>>;

export function parsePayload<T = any>(response: ToolEnvelope): T {
  return JSON.parse(response.content[0].text) as T;
}

export function createMockGraphClient(handlers: MockHandlers = {}) {
  const calls: Array<{ method: MockMethod; args: any[] }> = [];

  const makeMethod = (method: MockMethod) => {
    return async (...args: any[]) => {
      calls.push({ method, args });
      const handler = handlers[method];
      if (handler) {
        return await handler(...args);
      }

      return { success: true };
    };
  };

  return {
    calls,
    client: {
      get: makeMethod('get'),
      post: makeMethod('post'),
      patch: makeMethod('patch'),
      delete: makeMethod('delete'),
      uploadFile: makeMethod('uploadFile'),
      downloadFile: makeMethod('downloadFile'),
      cleanup: () => {},
    },
    methodCalls(method: MockMethod) {
      return calls.filter((call) => call.method === method);
    },
  };
}

export function createTempDir(prefix: string): string {
  return mkdtempSync(path.join(tmpdir(), prefix));
}

export function cleanupTempDir(dirPath: string): void {
  rmSync(dirPath, { recursive: true, force: true });
}

export function writeFileWithMtime(filePath: string, contents: string | Buffer, mtime: Date): void {
  mkdirSync(path.dirname(filePath), { recursive: true });
  writeFileSync(filePath, contents);
  utimesSync(filePath, mtime, mtime);
}
