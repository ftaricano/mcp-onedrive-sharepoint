import path from "node:path";
import fs from "node:fs";

export function getLocalFileRoot(): string {
  const root = path.resolve(process.env.MCP_LOCAL_FILE_ROOT || process.cwd());
  return fs.existsSync(root) ? fs.realpathSync(root) : root;
}

function isWithinRoot(candidate: string, root: string): boolean {
  const relative = path.relative(root, candidate);
  return (
    relative === "" ||
    (!!relative && !relative.startsWith("..") && !path.isAbsolute(relative))
  );
}

function nearestExistingAncestor(targetPath: string): string {
  let current = targetPath;

  while (!fs.existsSync(current)) {
    const parent = path.dirname(current);
    if (parent === current) {
      throw new Error(
        `No existing ancestor found for local path: ${targetPath}`,
      );
    }
    current = parent;
  }

  return current;
}

export function resolveLocalPath(
  inputPath: string,
  options: { mustExist?: boolean } = {},
): string {
  if (!inputPath || typeof inputPath !== "string") {
    throw new Error("Local path must be a non-empty string");
  }

  const root = getLocalFileRoot();
  const lexicalPath = path.isAbsolute(inputPath)
    ? path.resolve(inputPath)
    : path.resolve(root, inputPath);

  if (!isWithinRoot(lexicalPath, root)) {
    throw new Error(
      `Local path is outside MCP_LOCAL_FILE_ROOT (${root}). Set MCP_LOCAL_FILE_ROOT explicitly for trusted file access.`,
    );
  }

  if (options.mustExist && !fs.existsSync(lexicalPath)) {
    throw new Error(`Local path does not exist: ${lexicalPath}`);
  }

  const realPath = fs.existsSync(lexicalPath)
    ? fs.realpathSync(lexicalPath)
    : fs.realpathSync(nearestExistingAncestor(path.dirname(lexicalPath)));

  if (!isWithinRoot(realPath, root)) {
    throw new Error(
      `Local path resolves outside MCP_LOCAL_FILE_ROOT (${root}). Refusing symlink escape.`,
    );
  }

  return lexicalPath;
}

export function ensureParentDirectory(filePath: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
}
