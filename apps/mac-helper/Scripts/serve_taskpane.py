#!/usr/bin/env python3
import argparse
import datetime
import functools
import http.server
import os
import socketserver
import ssl


class TaskpaneRequestHandler(http.server.SimpleHTTPRequestHandler):
    log_path = os.environ.get("WORD_MCP_BRIDGE_TASKPANE_LOG", "/tmp/word-mcp-bridge-taskpane.log")

    def __init__(self, *args, directory=None, **kwargs):
        super().__init__(*args, directory=directory, **kwargs)

    @classmethod
    def append_log(cls, line: str) -> None:
        try:
            with open(cls.log_path, "a", encoding="utf-8") as handle:
                handle.write(line + "\n")
        except OSError:
            pass

    def log_message(self, format: str, *args) -> None:
        timestamp = datetime.datetime.now(datetime.timezone.utc).isoformat()
        message = format % args
        self.append_log(f"{timestamp} {self.command} {self.path} {message}")

    def do_GET(self):
        if self.path == "/healthz":
            body = b"ok"
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        return super().do_GET()

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()


class ThreadingHTTPServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True)
    parser.add_argument("--port", type=int, required=True)
    parser.add_argument("--cert", required=True)
    parser.add_argument("--key", required=True)
    args = parser.parse_args()

    if not os.path.exists(os.path.join(args.root, "taskpane.html")):
        raise SystemExit(f"taskpane root missing taskpane.html: {args.root}")
    if not os.path.exists(args.cert):
        raise SystemExit(f"missing cert: {args.cert}")
    if not os.path.exists(args.key):
        raise SystemExit(f"missing key: {args.key}")

    handler = functools.partial(TaskpaneRequestHandler, directory=args.root)
    server = ThreadingHTTPServer(("127.0.0.1", args.port), handler)
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(certfile=args.cert, keyfile=args.key)
    server.socket = context.wrap_socket(server.socket, server_side=True)
    server.serve_forever()


if __name__ == "__main__":
    main()
