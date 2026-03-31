#!/usr/bin/env node

import { existsSync, readdirSync, statSync } from "node:fs";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const binDir = path.dirname(fileURLToPath(import.meta.url));
const packageDir = path.resolve(binDir, "..");
const distCli = path.join(packageDir, "dist", "cli.js");
const srcDir = path.join(packageDir, "src");

function needsBuild() {
  if (!existsSync(distCli)) return true;

  const distMtime = statSync(distCli).mtimeMs;
  const queue = [srcDir];
  while (queue.length > 0) {
    const dir = queue.pop();
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        queue.push(full);
      } else if (entry.name.endsWith(".ts") && statSync(full).mtimeMs > distMtime) {
        return true;
      }
    }
  }
  return false;
}

if (needsBuild()) {
  const command = process.platform === "win32" ? "pnpm.cmd" : "pnpm";
  const result = spawnSync(command, ["build"], {
    cwd: packageDir,
    stdio: "inherit",
  });

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

await import(pathToFileURL(distCli).href);
