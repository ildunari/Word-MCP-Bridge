import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

const manifestDir = path.resolve(import.meta.dirname, "..");

function readManifest(name: string): string {
  return readFileSync(path.join(manifestDir, name), "utf8");
}

describe("production manifest", () => {
  it("uses a local-only production taskpane origin", () => {
    const manifest = readManifest("manifest.prod.xml");

    expect(manifest).toContain("https://localhost:3014/taskpane.html");
    expect(manifest).toContain("https://localhost:3014/commands.html");
    expect(manifest).toContain("<AppDomain>https://localhost:3014</AppDomain>");
    expect(manifest).not.toContain("word-mcp-bridge.pages.dev");
  });

  it("keeps the dev manifest on the separate dev-server origin", () => {
    const manifest = readManifest("manifest.xml");

    expect(manifest).toContain("https://localhost:3013/taskpane.html");
    expect(manifest).not.toContain("https://localhost:3014/taskpane.html");
  });
});
