#!/usr/bin/env node
// Remote entry: one authenticated Streamable HTTP MCP endpoint for mail, calendars and Teams. Backend credentials stay on the
// server; the MCP client holds only MCP_HTTP_TOKEN. One principal per deployment, like TeamsRelay's MCP_TOKEN.
//   POST/GET/DELETE /mcp   Authorization: Bearer <MCP_HTTP_TOKEN>, Host in MCP_HTTP_HOSTS, no Origin (no browser callers)
//   GET /healthz           {"ok":true}, no auth, no backend call: for the container healthcheck

import crypto from "node:crypto";
import http from "node:http";
import { createMcpHandler, type McpServer } from "@modelcontextprotocol/server";
import { outlookFromEnv, unifiedServer } from "./main.ts";
import { connectTeams } from "./teams.ts";

export const MAX_BODY = 4 * 1024 * 1024;
export const MIN_TOKEN = 32;

export interface HttpOptions {
  /** Bearer token clients present; at least MIN_TOKEN characters */
  token: string;
  /** Exact Host header values accepted on /mcp (public name behind the proxy, plus host:port for direct tests) */
  hosts: string[];
  /** One fresh server per HTTP request */
  factory: () => McpServer | Promise<McpServer>;
  maxBody?: number;
  /** One line per /mcp request: method, status, duration. Never headers or bodies. Default stderr */
  log?: (line: string) => void;
}

function bearerMatches(header: string | undefined, token: string): boolean {
  const got = Buffer.from(header ?? "");
  const want = Buffer.from(`Bearer ${token}`);
  return got.length === want.length && crypto.timingSafeEqual(got, want);
}

export function createHttpServer(
  o: HttpOptions,
): http.Server & { closeMcp(): Promise<void> } {
  if (o.token.length < MIN_TOKEN)
    throw new Error(
      `MCP_HTTP_TOKEN must contain at least ${MIN_TOKEN} characters`,
    );
  if (!o.hosts.length)
    throw new Error("MCP_HTTP_HOSTS must list the accepted Host headers");
  const maxBody = o.maxBody ?? MAX_BODY;
  const hosts = new Set(o.hosts.map((h) => h.toLowerCase()));
  const log = o.log ?? ((line: string) => console.error(line));
  const handler = createMcpHandler(() => o.factory(), {
    maxRequestBodySize: maxBody,
    onerror: (e) => log(`mcp: ${e.name}`),
  });

  const serve = async (
    req: http.IncomingMessage,
    res: http.ServerResponse,
  ): Promise<number> => {
    const path = new URL(req.url ?? "/", "http://local").pathname;
    if (path === "/healthz" && req.method === "GET") {
      res
        .writeHead(200, { "Content-Type": "application/json" })
        .end('{"ok":true}');
      return 200;
    }
    if (path !== "/mcp") return (void res.writeHead(404).end(), 404);
    const host = (req.headers.host ?? "").toLowerCase();
    if (!hosts.has(host)) return (void res.writeHead(421).end(), 421);
    if (req.headers.origin !== undefined)
      return (void res.writeHead(403).end(), 403);
    if (!bearerMatches(req.headers.authorization, o.token)) {
      res
        .writeHead(401, {
          "WWW-Authenticate": 'Bearer realm="communication-mcp"',
        })
        .end();
      return 401;
    }
    if (!["POST", "GET", "DELETE"].includes(req.method ?? "")) {
      res.writeHead(405, { Allow: "POST, GET, DELETE" }).end();
      return 405;
    }
    const declared = Number(req.headers["content-length"] ?? 0);
    if (declared > maxBody)
      return (void res.writeHead(413, { Connection: "close" }).end(), 413);
    const chunks: Buffer[] = [];
    let bytes = 0;
    for await (const c of req) {
      bytes += (c as Buffer).length;
      if (bytes > maxBody)
        return (void res.writeHead(413, { Connection: "close" }).end(), 413);
      chunks.push(c as Buffer);
    }

    // client gone: the abort reaches the running tool call (Teams forwards it upstream)
    const abort = new AbortController();
    res.on("close", () => {
      if (!res.writableFinished) abort.abort();
    });
    const headers = new Headers();
    for (const [k, v] of Object.entries(req.headers)) {
      if (v !== undefined && k !== "authorization")
        headers.set(k, Array.isArray(v) ? v.join(", ") : v);
    }
    const response = await handler.fetch(
      new Request(`http://${host}/mcp`, {
        method: req.method,
        headers,
        body:
          req.method === "POST"
            ? new Uint8Array(Buffer.concat(chunks))
            : undefined,
        signal: abort.signal,
      }),
    );
    res.writeHead(response.status, Object.fromEntries(response.headers));
    if (response.body) {
      const reader = response.body.getReader();
      res.on("close", () => void reader.cancel().catch(() => {}));
      try {
        for (;;) {
          const r = await reader.read();
          if (r.done) break;
          if (!res.write(r.value))
            await new Promise((ok) => res.once("drain", ok).once("close", ok));
        }
      } catch {
        // client closed mid-stream: nothing left to send to
      }
    }
    res.end();
    return response.status;
  };

  const server = http.createServer((req, res) => {
    const started = performance.now();
    const path = new URL(req.url ?? "/", "http://local").pathname.slice(0, 100);
    serve(req, res).then(
      (status) => {
        if (path !== "/healthz")
          log(
            `${req.method} ${path} ${status} ${Math.round(performance.now() - started)}ms`,
          );
      },
      (e: unknown) => {
        log(
          `${req.method} ${path} 500 ${e instanceof Error ? e.name : "error"}`,
        );
        if (!res.headersSent) res.writeHead(500);
        res.end();
      },
    );
  });
  server.requestTimeout = 60_000;
  server.headersTimeout = 15_000;
  return Object.assign(server, { closeMcp: () => handler.close() });
}

/** MCP_HTTP_HOSTS: comma-separated Host header values. */
export function hostsFromEnv(raw: string | undefined): string[] {
  return (raw ?? "")
    .split(",")
    .map((h) => h.trim().toLowerCase())
    .filter(Boolean);
}

if (import.meta.main) {
  const port = Number(process.env.PORT ?? 8190);
  if (!Number.isInteger(port) || port < 1 || port > 65535)
    throw new Error("PORT must be 1-65535");
  const outlook = outlookFromEnv();
  const teams = await connectTeams();
  const server = createHttpServer({
    token: process.env.MCP_HTTP_TOKEN ?? "",
    hosts: hostsFromEnv(process.env.MCP_HTTP_HOSTS),
    factory: () => unifiedServer(outlook, teams, true),
  });
  const bind = process.env.BIND_HOST ?? "127.0.0.1";
  server.listen(port, bind, () =>
    console.error(
      `outlook-mcp: remote MCP on ${bind}:${port}/mcp, Teams tools: ${teams.tools.length}${teams.error ? ` (${teams.error})` : ""}`,
    ),
  );
  const stop = () => {
    server.close();
    server.closeIdleConnections();
    void Promise.allSettled([server.closeMcp(), teams.close()]).then(() =>
      process.exit(0),
    );
    setTimeout(() => process.exit(0), 10_000).unref();
  };
  process.once("SIGTERM", stop);
  process.once("SIGINT", stop);
}
