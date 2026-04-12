#!/usr/bin/env node

import { createReadStream, existsSync, readFileSync, statSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { createSecureServer } from "node:http2";

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (!arg.startsWith("--")) {
      throw new Error(`Unknown argument: ${arg}`);
    }
    const key = arg.slice(2);
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) {
      throw new Error(`Missing value for ${arg}`);
    }
    result[key] = value;
    index += 1;
  }
  return result;
}

function contentType(pathname) {
  switch (extname(pathname).toLowerCase()) {
    case ".html":
      return "text/html; charset=utf-8";
    case ".js":
      return "text/javascript; charset=utf-8";
    case ".css":
      return "text/css; charset=utf-8";
    case ".json":
      return "application/json; charset=utf-8";
    case ".svg":
      return "image/svg+xml";
    case ".png":
      return "image/png";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".ico":
      return "image/x-icon";
    default:
      return "application/octet-stream";
  }
}

function safeAssetPath(root, requestPath) {
  const requested = requestPath === "/" ? "/taskpane.html" : requestPath;
  const normalizedPath = normalize(requested).replace(/^(\.\.[/\\])+/, "");
  return join(root, normalizedPath);
}

const args = parseArgs(process.argv.slice(2));
const root = args.root;
const port = Number.parseInt(args.port ?? "3014", 10);
const certPath = args.cert;
const keyPath = args.key;

if (!root || !certPath || !keyPath || Number.isNaN(port)) {
  throw new Error("Expected --root, --port, --cert, and --key.");
}
if (!existsSync(join(root, "taskpane.html"))) {
  throw new Error(`taskpane root missing taskpane.html: ${root}`);
}
if (!existsSync(certPath)) {
  throw new Error(`missing cert: ${certPath}`);
}
if (!existsSync(keyPath)) {
  throw new Error(`missing key: ${keyPath}`);
}

const server = createSecureServer(
  {
    allowHTTP1: true,
    cert: readFileSync(certPath),
    key: readFileSync(keyPath),
  },
  (req, res) => {
    const path = req.url?.split("?")[0] ?? "/";
    res.setHeader("Access-Control-Allow-Origin", "*");

    if (path === "/healthz") {
      const body = Buffer.from("ok", "utf8");
      res.writeHead(200, {
        "Content-Type": "text/plain; charset=utf-8",
        "Content-Length": body.byteLength,
      });
      res.end(body);
      return;
    }

    const assetPath = safeAssetPath(root, path);
    if (!assetPath.startsWith(root) || !existsSync(assetPath) || !statSync(assetPath).isFile()) {
      const body = Buffer.from("not found", "utf8");
      res.writeHead(404, {
        "Content-Type": "text/plain; charset=utf-8",
        "Content-Length": body.byteLength,
      });
      res.end(body);
      return;
    }

    res.writeHead(200, { "Content-Type": contentType(assetPath) });
    createReadStream(assetPath).pipe(res);
  },
);

server.listen(port, "127.0.0.1");
