import { cp, mkdir, rm } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { spawn } from "node:child_process";

const repoRoot = process.cwd();
const addinDir = path.join(repoRoot, "packages", "word-addin");
const distDir = path.join(addinDir, "dist");
const releaseRoot = path.join(repoRoot, "release");
const bundleDir = path.join(releaseRoot, "word-addin-bundle");
const zipPath = path.join(releaseRoot, "word-addin-bundle.zip");

async function main() {
  await mkdir(releaseRoot, { recursive: true });
  await rm(bundleDir, { recursive: true, force: true });
  await rm(zipPath, { force: true });
  await mkdir(bundleDir, { recursive: true });

  await cp(path.join(addinDir, "manifest.xml"), path.join(bundleDir, "manifest.xml"));
  await cp(
    path.join(addinDir, "manifest.prod.xml"),
    path.join(bundleDir, "manifest.prod.xml"),
  );
  await cp(distDir, path.join(bundleDir, "dist"), { recursive: true });

  await zipBundle(bundleDir, zipPath);

  console.log(`Created add-in bundle at ${zipPath}`);
  console.log(`Bundle directory: ${bundleDir}`);
}

function zipBundle(bundleDirectory, destinationZip) {
  return new Promise((resolve, reject) => {
    const child = spawn("ditto", ["-c", "-k", "--sequesterRsrc", "--keepParent", bundleDirectory, destinationZip], {
      stdio: "inherit",
    });

    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(`ditto exited with code ${code ?? 1}`));
    });
    child.on("error", reject);
  });
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
