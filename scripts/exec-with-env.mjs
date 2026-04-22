#!/usr/bin/env node

import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const envPath = path.join(repoRoot, ".env");
const argv = process.argv.slice(2);

if (argv.length === 0) {
  console.error("Usage: exec-with-env.mjs <command> [args...]");
  process.exit(64);
}

let fileEnv = {};

if (fs.existsSync(envPath)) {
  try {
    fileEnv = dotenv.parse(fs.readFileSync(envPath));
  } catch (error) {
    console.error(`Failed to parse ${envPath}:`, error);
    process.exit(1);
  }
}

const child = spawn(argv[0], argv.slice(1), {
  cwd: repoRoot,
  env: {
    ...fileEnv,
    ...process.env,
  },
  stdio: "inherit",
});

child.on("error", (error) => {
  console.error(`Failed to start ${argv[0]}:`, error.message);
  process.exit(1);
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
    return;
  }

  process.exit(code ?? 1);
});
