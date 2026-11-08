// Remote entry end to end: the official MCP client over Streamable HTTP -> http.ts -> unified server -> OE Copilot gateway
// fixture (mail, calendars) and TeamsRelay /mcp fixture (chats), all real HTTP on 127.0.0.1.
import assert from "node:assert/strict";
import http from "node:http";
import net from "node:net";
import { test } from "node:test";
import {
  Client,
  StreamableHTTPClientTransport,
} from "@modelcontextprotocol/client";
import { createMcpHandler, McpServer } from "@modelcontextprotocol/server";
import { createHttpServer } from "../src/http.ts";
import { unifiedServer } from "../src/main.ts";
import { ServiceOutlook, httpGateway } from "../src/service.ts";
import { connectTeams } from "../src/teams.ts";

const TOKEN = "t".repeat(40);
const GATEWAY_KEY = "gateway-key";
const RELAY_TOKEN = "relay-token-0123456789";

async function freePort(): Promise<number> {
  const s = net.createServer();
  await new Promise<void>((ok) => s.listen(0, "127.0.0.1", ok));
  const { port } = s.address() as net.AddressInfo;
  await new Promise<void>((ok) => s.close(() => ok()));
  return port;
}

async function listen(server: http.Server, port = 0): Promise<number> {
  await new Promise<void>((ok) => server.listen(port, "127.0.0.1", ok));
  return (server.address() as net.AddressInfo).port;
}

async function stop(server: http.Server) {
  server.closeAllConnections();
  await new Promise<void>((ok) => server.close(() => ok()));
}

/** OE Copilot communication gateway: one mailbox, one unread message, key checked. */
function gatewayFixture(reads: unknown[]) {
  return http.createServer(async (req, res) => {
    res.setHeader("Content-Type", "application/json");
    if (req.headers.authorization !== `Bearer ${GATEWAY_KEY}`)
      return void res.writeHead(401).end('{"errore":"non autorizzato"}');
    if (req.url === "/api/communication/accounts")
      return void res.end(
        JSON.stringify({
          accounts: [{ upn: "a@msc.example", status: "ok", message: "" }],
        }),
      );
    let body = "";
    for await (const c of req) body += c;
    reads.push(JSON.parse(body));
    res.end(
      JSON.stringify({
        value: [
          {
            Id: "mail-1",
            Subject: "fixture mail",
            IsRead: false,
            ReceivedDateTime: "2026-11-08T08:00:00Z",
          },
        ],
      }),
    );
  });
}

/** TeamsRelay /mcp: one read-only tool; counts tools/list calls. */
function relayFixture(stats: { lists: number }) {
  const mcp = createMcpHandler(() => {
    const s = new McpServer({ name: "teamsrelay", version: "1" });
    s.registerTool(
      "list_chats",
      { description: "chats", annotations: { readOnlyHint: true } },
      async () => ({
        content: [{ type: "text", text: "chat-fixture" }],
      }),
    );
    return s;
  });
  return http.createServer(async (req, res) => {
    if (req.headers.authorization !== `Bearer ${RELAY_TOKEN}`)
      return void res.writeHead(401).end();
    const chunks: Buffer[] = [];
    for await (const c of req) chunks.push(c as Buffer);
    const body = Buffer.concat(chunks);
    if (body.includes('"tools/list"')) stats.lists++;
    const headers = new Headers();
    for (const [k, v] of Object.entries(req.headers))
      if (typeof v === "string") headers.set(k, v);
    const r = await mcp.fetch(
      new Request(`http://127.0.0.1${req.url}`, {
        method: req.method,
        headers,
        body: req.method === "POST" ? new Uint8Array(body) : undefined,
      }),
    );
    res.writeHead(r.status, Object.fromEntries(r.headers));
    res.end(Buffer.from(await r.arrayBuffer()));
  });
}

/** A raw HTTP/1.1 request, for headers fetch will not send (Host). Returns the status line. */
function raw(port: number, request: string): Promise<string> {
  return new Promise((ok, ko) => {
    const sock = net.connect(port, "127.0.0.1", () => sock.end(request));
    let data = "";
    sock.on("data", (d) => (data += d));
    sock.on("end", () => ok(data.split("\r\n")[0] ?? ""));
    sock.on("error", ko);
  });
}

async function connect(url: string, token = TOKEN) {
  const client = new Client({ name: "remote-test", version: "1" });
  await client.connect(
    new StreamableHTTPClientTransport(new URL(url), {
      authProvider: { token: async () => token },
    }),
  );
  return client;
}

const text = (r: unknown) => JSON.stringify(r);

test("remote MCP: mail through the OE gateway and Teams through the relay, one discovery for every request", async () => {
  const reads: unknown[] = [];
  const stats = { lists: 0 };
  const gateway = gatewayFixture(reads);
  const relay = relayFixture(stats);
  const gatewayPort = await listen(gateway);
  const relayPort = await listen(relay);
  const teams = await connectTeams({
    url: `http://127.0.0.1:${relayPort}/mcp`,
    token: RELAY_TOKEN,
  });
  const outlook = new ServiceOutlook(
    httpGateway(`http://127.0.0.1:${gatewayPort}`, GATEWAY_KEY),
  );
  const port = await freePort();
  const lines: string[] = [];
  const server = createHttpServer({
    token: TOKEN,
    hosts: [`127.0.0.1:${port}`],
    factory: () => unifiedServer(outlook, teams, true),
    log: (l) => lines.push(l),
  });
  await listen(server, port);
  const url = `http://127.0.0.1:${port}/mcp`;
  const clients: Client[] = [];
  try {
    for (let i = 0; i < 2; i++) {
      const c = await connect(url);
      clients.push(c);
      const names = (await c.listTools()).tools.map((t) => t.name);
      assert.ok(
        names.includes("list_unread") &&
          names.includes("list_calendar_events") &&
          names.includes("get_attachment"),
      );
      assert.ok(names.includes("teams_list_chats"));
      assert.ok(
        !names.includes("send_email"),
        "service mode exposes no mailbox writes",
      );
      assert.match(
        text(await c.callTool({ name: "teams_list_chats", arguments: {} })),
        /chat-fixture/,
      );
      const unread = await c.callTool({
        name: "list_unread",
        arguments: { limit: 5 },
      });
      assert.ok(!unread.isError);
      assert.match(text(unread), /fixture mail/);
    }
    assert.equal(
      stats.lists,
      1,
      "Teams tools discovered once, not per HTTP request",
    );
    assert.ok(reads.length >= 2);
    assert.ok(
      reads.every(
        (r) => (r as { account: string }).account === "a@msc.example",
      ),
    );
    assert.ok(lines.some((l) => /^POST \/mcp 200 \d+ms$/.test(l)));
    assert.ok(
      !lines.some(
        (l) =>
          l.includes(TOKEN) ||
          l.includes(GATEWAY_KEY) ||
          l.includes(RELAY_TOKEN),
      ),
      "no secret in the log",
    );
  } finally {
    for (const c of clients) await c.close();
    await server.closeMcp();
    await teams.close();
    await stop(server);
    await stop(gateway);
    await stop(relay);
  }
});

test("remote MCP guards: auth, Host, Origin, path, method, body size; healthz open", async () => {
  const port = await freePort();
  const server = createHttpServer({
    token: TOKEN,
    hosts: [`127.0.0.1:${port}`],
    factory: () => new McpServer({ name: "x", version: "1" }),
    maxBody: 1024,
    log: () => {},
  });
  await listen(server, port);
  const base = `http://127.0.0.1:${port}`;
  const auth = {
    Authorization: `Bearer ${TOKEN}`,
    "Content-Type": "application/json",
    Accept: "application/json, text/event-stream",
  };
  try {
    assert.equal(
      (await fetch(`${base}/mcp`, { method: "POST", body: "{}" })).status,
      401,
    );
    const wrong = await fetch(`${base}/mcp`, {
      method: "POST",
      headers: { Authorization: `Bearer ${"x".repeat(40)}` },
      body: "{}",
    });
    assert.equal(wrong.status, 401);
    assert.match(wrong.headers.get("www-authenticate") ?? "", /^Bearer/);
    assert.equal(
      (
        await fetch(`${base}/mcp`, {
          method: "POST",
          headers: { ...auth, Origin: "https://evil.example" },
          body: "{}",
        })
      ).status,
      403,
    );
    assert.equal((await fetch(`${base}/other`, { headers: auth })).status, 404);
    assert.equal(
      (await fetch(`${base}/mcp`, { method: "PUT", headers: auth, body: "{}" }))
        .status,
      405,
    );
    assert.equal(
      (
        await fetch(`${base}/mcp`, {
          method: "POST",
          headers: auth,
          body: "x".repeat(2048),
        })
      ).status,
      413,
    );
    const health = await fetch(`${base}/healthz`);
    assert.equal(health.status, 200);
    assert.deepEqual(await health.json(), { ok: true });
    const status = await raw(
      port,
      `POST /mcp HTTP/1.1\r\nHost: evil.example\r\nAuthorization: Bearer ${TOKEN}\r\nContent-Length: 2\r\nConnection: close\r\n\r\n{}`,
    );
    assert.match(status, /^HTTP\/1\.1 421/);
    const upper = await raw(
      port,
      `GET /healthz HTTP/1.1\r\nHost: anything.example\r\nConnection: close\r\n\r\n`,
    );
    assert.match(upper, /^HTTP\/1\.1 200/);
  } finally {
    await server.closeMcp();
    await stop(server);
  }
});

test("remote MCP: Teams down at startup, tools appear once the relay answers, mail unaffected meanwhile", async () => {
  const stats = { lists: 0 };
  const relayPort = await freePort();
  const teams = await connectTeams({
    url: `http://127.0.0.1:${relayPort}/mcp`,
    token: RELAY_TOKEN,
    retryMs: 0,
    connectTimeoutMs: 2000,
  });
  assert.equal(teams.tools.length, 0);
  assert.match(teams.error ?? "", /unavailable/);
  const gateway = gatewayFixture([]);
  const gatewayPort = await listen(gateway);
  const outlook = new ServiceOutlook(
    httpGateway(`http://127.0.0.1:${gatewayPort}`, GATEWAY_KEY),
  );
  const port = await freePort();
  const server = createHttpServer({
    token: TOKEN,
    hosts: [`127.0.0.1:${port}`],
    factory: () => unifiedServer(outlook, teams, true),
    log: () => {},
  });
  await listen(server, port);
  const relay = relayFixture(stats);
  const clients: Client[] = [];
  try {
    const first = await connect(`http://127.0.0.1:${port}/mcp`);
    clients.push(first);
    const before = (await first.listTools()).tools.map((t) => t.name);
    assert.ok(
      before.includes("list_unread") &&
        !before.some((n) => n.startsWith("teams_")),
    );
    await listen(relay, relayPort);
    // the first settles a discovery the requests above may have started while the relay was down; the second runs against it
    await teams.refresh();
    await teams.refresh();
    assert.equal(teams.error, undefined);
    const second = await connect(`http://127.0.0.1:${port}/mcp`);
    clients.push(second);
    assert.match(
      text(await second.callTool({ name: "teams_list_chats", arguments: {} })),
      /chat-fixture/,
    );
  } finally {
    for (const c of clients) await c.close();
    await server.closeMcp();
    await teams.close();
    await stop(server);
    await stop(gateway);
    await stop(relay);
  }
});
