import { svelte } from "@sveltejs/vite-plugin-svelte";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import { viteStaticCopy } from "vite-plugin-static-copy";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

async function getHttpsOptions() {
  try {
    const devCerts = await import("office-addin-dev-certs");
    const certs = await devCerts.getHttpsServerOptions();
    return { ca: certs.ca, key: certs.key, cert: certs.cert };
  } catch {
    console.warn("Could not load office-addin-dev-certs, HTTPS disabled");
    return undefined;
  }
}

export default defineConfig(async () => ({
  root: "src",
  publicDir: "../public",
  build: {
    outDir: "../dist",
    emptyOutDir: true,
    sourcemap: true,
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, "src/taskpane.html"),
        commands: path.resolve(__dirname, "src/commands.html"),
      },
    },
  },
  define: {
    __APP_VERSION__: JSON.stringify("0.0.1"),
  },
  plugins: [
    svelte(),
    viteStaticCopy({
      targets: [
        {
          src: "../manifest*.xml",
          dest: ".",
        },
      ],
    }),
  ],
  server: {
    https: await getHttpsOptions(),
    port: 3013,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
  },
}));
