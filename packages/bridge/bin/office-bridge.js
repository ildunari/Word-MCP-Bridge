#!/usr/bin/env node

import { existsSync } from "node:fs";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const binDir = path.dirname(fileURLToPath(import.meta.url));
const packageDir = path.resolve(binDir, "..");
const distCli = path.join(packageDir, "dist", "cli.js");
const packageJson = path.join(packageDir, "package.json");
const srcDir = path.join(packageDir, "src");

if (!existsSync(distCli) && existsSync(packageJson) && existsSync(srcDir)) {
  const command = process.platform === "win32" ? "pnpm.cmd" : "pnpm";
  const result = spawnSync(command, ["build"], {
    cwd: packageDir,
    stdio: "inherit",
  });

  if (result.status !== 0) {
    process.exit(result.status ?? 1);
  }
}

if (!existsSync(distCli)) {
  console.error(
    "office-bridge could not find a built CLI. Reinstall the package or run `pnpm build` in the repo checkout.",
  );
  process.exit(1);
}

await import(pathToFileURL(distCli).href);
