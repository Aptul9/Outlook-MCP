import assert from "node:assert/strict";
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import type { AddressInfo } from "node:net";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/client";
import { createMcpHandler, InMemoryTransport, McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";
import { describe, registerTeamsTools, TeamsConfigError, teamsConfig, teamsConfigFromEnv } from "../src/teams.ts";

const TOKEN = "test-token-123";
const WRITE = { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true };
const READ = { readOnlyHint: true, openWorldHint: false };

type Seen = { method: string; auth: string; body: string };

/** A local HTTP server on 127.0.0.1; `handle` answers in node:http terms. */
async function listen(handle: (req: IncomingMessage, res: ServerResponse, body: Buffer) => void | Promise<void>) {
  const srv = createServer(async (req, res) => {
    const chunks: Buffer[] = [];
    for await (const c of req) chunks.push(c as Buffer);
    await handle(req, res, Buffer.concat(chunks));
  });
  await new Promise<void>((ok) => srv.listen(0, "127.0.0.1", ok));
  const { port } = srv.address() as AddressInfo;
  return {
    url: `http://127.0.0.1:${port}/mcp`,
    close: () => {
      srv.closeAllConnections();
      return new Promise<void>((ok) => srv.close(() => ok()));
    },
  };
}

/** Fake TeamsRelay: the SDK's stateless createMcpHandler (as TeamsRelay serves /mcp), bearer TOKEN required. */
async function fakeRelay() {
  const seen: Seen[] = [];
  const handler = createMcpHandler(() => {
    const s = new McpServer({ name: "fake-teamsrelay", version: "1.0.0" });
    s.registerTool(
      "send_message",
      {
        title: "Send a message",
        description: "Sends a Teams message",
        inputSchema: z.object({ chat: z.string().describe("Chat name"), text: z.string().min(1) }),
        outputSchema: z.object({ id: z.string(), chat: z.string() }),
        annotations: WRITE,
        _meta: { "example/kind": "write" },
      },
      async ({ chat, text }) => ({
        content: [
          { type: "text", text: `sent ${text}` },
          { type: "image", data: "aGk=", mimeType: "image/png" },
          { type: "resource_link", uri: "https://teams.example/chat/1", name: "chat" },
        ],
        structuredContent: { id: "m1", chat },
      }),
    );
    s.registerTool("fails", { description: "Always a tool error", annotations: READ }, async () => ({ content: [{ type: "text", text: "Chat not found" }], isError: true }));
    s.registerTool("list_accounts", { description: "Teams accounts", annotations: READ }, async () => ({ content: [{ type: "text", text: "slot 1" }] }));
    s.registerTool("dupe", { description: "Collides with a local tool" }, async () => ({ content: [{ type: "text", text: "upstream dupe" }] }));
    return s;
  });
  const srv = await listen(async (req, res, body) => {
    seen.push({ method: req.method ?? "", auth: String(req.headers.authorization ?? ""), body: body.toString() });
    if (req.headers.authorization !== `Bearer ${TOKEN}`) {
      res.writeHead(401, { "WWW-Authenticate": 'Bearer realm="teamsrelay"' }).end("Missing or wrong token");
      return;
    }
    const headers = new Headers();
    for (const [k, v] of Object.entries(req.headers)) if (v !== undefined) headers.set(k, String(v));
    const hasBody = req.method !== "GET" && req.method !== "HEAD" && req.method !== "DELETE";
    const r = await handler.fetch(new Request(`http://127.0.0.1${req.url}`, { method: req.method, headers, body: hasBody ? new Uint8Array(body) : undefined }));
    res.writeHead(r.status, Object.fromEntries(r.headers));
    res.end(Buffer.from(await r.arrayBuffer()));
  });
  return { ...srv, seen, calls: () => seen.filter((s) => s.body.includes('"tools/call"')).length };
}

/** This server, with a local tool and a pre-taken teams_ name, linked to an in-memory MCP client. */
async function local(register: (s: McpServer) => Promise<unknown>) {
  const server = new McpServer({ name: "outlook", version: "3.0.0" });
  server.registerTool("list_accounts", { description: "Mailboxes", annotations: READ }, async () => ({ content: [{ type: "text", text: "mailboxes" }] }));
  server.registerTool("teams_dupe", { description: "Local tool" }, async () => ({ content: [{ type: "text", text: "local dupe" }] }));
  const relay = await register(server);
  const [a, b] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test", version: "1.0.0" });
  await Promise.all([server.connect(a), client.connect(b)]);
  return { server, client, relay };
}

test("off without a URL: no request, no tool", async () => {
  let fetches = 0;
  const spy = (async () => (fetches++, new Response(null, { status: 500 }))) as typeof fetch;
  const server = new McpServer({ name: "outlook", version: "3.0.0" });
  for (const env of [{}, { OUTLOOK_MCP_TEAMS_TOKEN: TOKEN }, { OUTLOOK_MCP_TEAMS_URL: "  " }]) {
    const relay = await registerTeamsTools(server, { env, fetch: spy });
    assert.deepEqual(relay.tools, []);
    assert.equal(relay.error, undefined);
    await relay.close();
  }
  assert.equal(fetches, 0);
  assert.equal(teamsConfigFromEnv({}), null);
});

test("config: https or localhost http only, no credentials or query, token required", () => {
  for (const url of ["https://relay.example/mcp", "http://localhost:3000/mcp", "http://127.0.0.1:1/mcp", "http://[::1]:1/mcp"]) {
    assert.equal(teamsConfig(url, TOKEN)?.url.href, new URL(url).href);
  }
  for (const url of ["http://relay.example/mcp", "http://127.0.0.2/mcp", "https://u:p@relay.example/mcp", "https://relay.example/mcp?x=1", "https://relay.example/mcp#f", "ftp://relay.example/mcp", "not a url", "file:///etc/passwd"]) {
    assert.throws(() => teamsConfig(url, TOKEN), TeamsConfigError, url);
  }
  assert.throws(() => teamsConfig("https://relay.example/mcp", undefined), /needs OUTLOOK_MCP_TEAMS_TOKEN/);
  assert.throws(() => teamsConfig("https://relay.example/mcp", "a b"), /HTTP header/);
  assert.equal(teamsConfigFromEnv({ OUTLOOK_MCP_TEAMS_URL: "https://relay.example/mcp", OUTLOOK_MCP_TEAMS_TOKEN: TOKEN })?.token, TOKEN);
});

/** Runs fn with console.error captured; returns what it wrote. */
async function stderrOf(fn: () => Promise<void>): Promise<string> {
  const lines: string[] = [];
  const orig = console.error;
  console.error = (...a: unknown[]) => void lines.push(a.map(String).join(" "));
  try {
    await fn();
  } finally {
    console.error = orig;
  }
  return lines.join("\n");
}

test("bad config: no throw, no request, error set, local tools untouched", async () => {
  let fetches = 0;
  const spy = (async () => (fetches++, new Response(null, { status: 500 }))) as typeof fetch;
  const server = new McpServer({ name: "outlook", version: "3.0.0" });
  server.registerTool("list_accounts", { description: "Mailboxes" }, async () => ({ content: [{ type: "text", text: "mailboxes" }] }));
  const cases = [
    { OUTLOOK_MCP_TEAMS_URL: "http://relay.example/mcp", OUTLOOK_MCP_TEAMS_TOKEN: TOKEN },
    { OUTLOOK_MCP_TEAMS_URL: "https://relay.example/mcp" },
    { OUTLOOK_MCP_TEAMS_URL: "not a url", OUTLOOK_MCP_TEAMS_TOKEN: TOKEN },
    { OUTLOOK_MCP_TEAMS_URL: `https://user:${TOKEN}@relay.example/mcp`, OUTLOOK_MCP_TEAMS_TOKEN: TOKEN },
  ];
  const err = await stderrOf(async () => {
    for (const env of cases) {
      const r = await registerTeamsTools(server, { env, fetch: spy });
      assert.deepEqual(r.tools, []);
      assert.match(r.error ?? "", /^Teams tools off, invalid configuration: OUTLOOK_MCP_TEAMS_/);
      await r.close();
    }
  });
  assert.equal(fetches, 0);
  assert.equal(err.split("\n").length, cases.length);
  assert.ok(!err.includes(TOKEN));
  const [a, b] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test", version: "1.0.0" });
  await Promise.all([server.connect(a), client.connect(b)]);
  assert.deepEqual((await client.listTools()).tools.map((t) => t.name), ["list_accounts"]);
  await server.close();
});

test("relay text and the token never reach stderr", async () => {
  const SECRET = "PAYLOAD-SECRET-42";
  // Answers initialize normally, then each later request differently: HTTP 500 echoing the token and payload, a JSON-RPC
  // error echoing them, or an SSE stream with a malformed event carrying them.
  const relay = await fakeRelay();
  let mode: "pass" | "http500" | "rpc" | "sse" = "pass";
  const evil = (async (input: string | URL, init?: RequestInit) => {
    const body = String(init?.body ?? "");
    const id = body.startsWith("{") ? JSON.parse(body).id : undefined;
    const echo = `${TOKEN} ${SECRET}`;
    if (mode === "http500" && init?.method === "POST") return new Response(`boom ${echo}`, { status: 500 });
    if (mode === "rpc" && id !== undefined) return Response.json({ jsonrpc: "2.0", id, error: { code: -32603, message: `failed ${echo}` } });
    if (mode === "sse" && id !== undefined) return new Response(`event: message\ndata: {not json ${echo}\n\n`, { headers: { "content-type": "text/event-stream" } });
    return fetch(input, init);
  }) as typeof fetch;
  try {
    const results: string[] = [];
    const err = await stderrOf(async () => {
      mode = "http500";
      const down = await registerTeamsTools(new McpServer({ name: "o", version: "1" }), { url: relay.url, token: TOKEN, env: {}, fetch: evil });
      assert.match(down.error ?? "", /^Teams relay unavailable: HTTP 500$/);
      mode = "pass";
      const { client, relay: handle } = await local((s) => registerTeamsTools(s, { url: relay.url, token: TOKEN, env: {}, fetch: evil, callTimeoutMs: 500 }));
      for (const m of ["http500", "rpc", "sse"] as const) {
        mode = m;
        const r = await client.callTool({ name: "teams_list_accounts", arguments: {} });
        assert.equal(r.isError, true, m);
        results.push((r.content as any)[0].text);
      }
      mode = "pass";
      await (handle as Awaited<ReturnType<typeof registerTeamsTools>>).close();
    });
    assert.ok(err.length > 0, "something was logged");
    assert.ok(!err.includes(TOKEN), `token in stderr:\n${err}`);
    assert.ok(!err.includes(SECRET), `relay text in stderr:\n${err}`);
    assert.ok(!err.includes("boom") && !err.includes("not json"), `relay text in stderr:\n${err}`);
    assert.equal(results[0], "Teams relay error: HTTP 500");
    assert.match(results[1]!, /^Teams relay error -32603: .*failed \[redacted\] PAYLOAD-SECRET-42/);
    assert.ok(results.every((t) => !t.includes(TOKEN)), results.join("\n"));
  } finally {
    await relay.close();
  }
});

test("describe: kinds only", () => {
  assert.equal(describe(new TypeError("fetch failed", { cause: { code: "ECONNREFUSED" } })), "network error (ECONNREFUSED)");
  assert.equal(describe(new TypeError("fetch failed", { cause: { code: "bad code with secret" } })), "network error");
  assert.equal(describe(new Error(`some ${TOKEN}`)), "unexpected error");
  assert.equal(describe(`string ${TOKEN}`), "unexpected error");
});

test("relay tools: teams_ prefix, schemas and annotations kept, results and errors forwarded unchanged", async () => {
  const relay = await fakeRelay();
  try {
    const { client, relay: handle } = await local((s) => registerTeamsTools(s, { env: { OUTLOOK_MCP_TEAMS_URL: relay.url, OUTLOOK_MCP_TEAMS_TOKEN: TOKEN } }));
    const h = handle as Awaited<ReturnType<typeof registerTeamsTools>>;
    assert.equal(h.error, undefined);
    assert.deepEqual([...h.tools].sort(), ["teams_fails", "teams_list_accounts", "teams_send_message"]);

    const { tools } = await client.listTools();
    const byName = new Map(tools.map((t) => [t.name, t]));
    assert.deepEqual([...byName.keys()].sort(), ["list_accounts", "teams_dupe", "teams_fails", "teams_list_accounts", "teams_send_message"]);
    assert.equal(byName.get("list_accounts")!.description, "Mailboxes");
    assert.equal(byName.get("teams_dupe")!.description, "Local tool");
    const send = byName.get("teams_send_message")!;
    assert.equal(send.title, "Send a message");
    assert.equal(send.description, "Sends a Teams message");
    assert.deepEqual(send.annotations, WRITE);
    assert.deepEqual(send._meta, { "example/kind": "write" });
    assert.deepEqual(send.inputSchema.required, ["chat", "text"]);
    assert.equal((send.inputSchema.properties as any).chat.description, "Chat name");
    assert.deepEqual((send.outputSchema as any).required, ["id", "chat"]);
    assert.deepEqual(byName.get("teams_fails")!.annotations, READ);

    const ok = await client.callTool({ name: "teams_send_message", arguments: { chat: "Ops", text: "hello" } });
    assert.deepEqual(ok.content, [
      { type: "text", text: "sent hello" },
      { type: "image", data: "aGk=", mimeType: "image/png" },
      { type: "resource_link", uri: "https://teams.example/chat/1", name: "chat" },
    ]);
    assert.deepEqual(ok.structuredContent, { id: "m1", chat: "Ops" });
    assert.ok(!ok.isError);

    const bad = await client.callTool({ name: "teams_fails", arguments: {} });
    assert.equal(bad.isError, true);
    assert.deepEqual(bad.content, [{ type: "text", text: "Chat not found" }]);

    const before = relay.calls();
    const invalid = await client.callTool({ name: "teams_send_message", arguments: { chat: "Ops" } });
    assert.equal(invalid.isError, true);
    assert.match((invalid.content as any)[0].text, /Input validation error/);
    assert.equal(relay.calls(), before, "invalid arguments never reach the relay");

    const mine = await client.callTool({ name: "teams_dupe", arguments: {} });
    assert.deepEqual(mine.content, [{ type: "text", text: "local dupe" }]);

    assert.ok(relay.seen.length > 0);
    assert.ok(relay.seen.every((s) => s.auth === `Bearer ${TOKEN}`));
    assert.deepEqual(relay.seen.filter((s) => s.body.includes('"tools/call"')).map((s) => JSON.parse(s.body).params.name), ["send_message", "fails"]);

    await h.close();
    const after = await client.callTool({ name: "teams_list_accounts", arguments: {} });
    assert.equal(after.isError, true);
    assert.match((after.content as any)[0].text, /^Teams relay error: /);
    await h.close();
  } finally {
    await relay.close();
  }
});

test("closing this server's transport closes the relay connection", async () => {
  const relay = await fakeRelay();
  try {
    const { server, client } = await local((s) => registerTeamsTools(s, { url: relay.url, token: TOKEN, env: {} }));
    assert.deepEqual((await client.callTool({ name: "teams_list_accounts", arguments: {} })).content, [{ type: "text", text: "slot 1" }]);
    await server.close();
    const [a, b] = InMemoryTransport.createLinkedPair();
    const again = new Client({ name: "test", version: "1.0.0" });
    await Promise.all([server.connect(a), again.connect(b)]);
    const r = await again.callTool({ name: "teams_list_accounts", arguments: {} });
    assert.equal(r.isError, true);
    assert.match((r.content as any)[0].text, /^Teams relay error: /);
    await server.close();
  } finally {
    await relay.close();
  }
});

test("wrong token: no tool, error reported, nothing thrown", async () => {
  const relay = await fakeRelay();
  try {
    const server = new McpServer({ name: "outlook", version: "3.0.0" });
    const r = await registerTeamsTools(server, { url: relay.url, token: "wrong-token", env: {} });
    assert.deepEqual(r.tools, []);
    assert.match(r.error ?? "", /^Teams relay unavailable: /);
    assert.ok(relay.seen.length > 0 && relay.seen.every((s) => s.auth === "Bearer wrong-token"));
  } finally {
    await relay.close();
  }
});

test("redirect refused: the token never reaches the redirect target", async () => {
  const target: Seen[] = [];
  const other = await listen((req, res) => {
    target.push({ method: req.method ?? "", auth: String(req.headers.authorization ?? ""), body: "" });
    res.writeHead(200, { "Content-Type": "application/json" }).end("{}");
  });
  const redirector = await listen((_req, res) => void res.writeHead(307, { Location: other.url }).end());
  try {
    const r = await registerTeamsTools(new McpServer({ name: "outlook", version: "3.0.0" }), { url: redirector.url, token: TOKEN, env: {} });
    assert.deepEqual(r.tools, []);
    assert.match(r.error ?? "", /^Teams relay unavailable: /);
    assert.equal(target.length, 0);
  } finally {
    await redirector.close();
    await other.close();
  }
});

test("a relay that never answers: discovery gives up within the connect timeout", async () => {
  const silent = await listen(() => {});
  try {
    const t0 = Date.now();
    const r = await registerTeamsTools(new McpServer({ name: "outlook", version: "3.0.0" }), { url: silent.url, token: TOKEN, env: {}, connectTimeoutMs: 300 });
    assert.deepEqual(r.tools, []);
    assert.match(r.error ?? "", /^Teams relay unavailable: /);
    assert.ok(Date.now() - t0 < 3000, `took ${Date.now() - t0} ms`);
  } finally {
    await silent.close();
  }
});

test("a forwarded call is bounded by the call timeout", async () => {
  const relay = await fakeRelay();
  let hold = false;
  const slow = (async (input: string | URL, init?: RequestInit) => {
    if (hold && String(init?.body ?? "").includes('"tools/call"')) return new Promise<Response>(() => {});
    return fetch(input, init);
  }) as typeof fetch;
  try {
    const { client, relay: handle } = await local((s) => registerTeamsTools(s, { url: relay.url, token: TOKEN, env: {}, callTimeoutMs: 300, fetch: slow }));
    hold = true;
    const t0 = Date.now();
    const r = await client.callTool({ name: "teams_list_accounts", arguments: {} });
    assert.equal(r.isError, true);
    assert.match((r.content as any)[0].text, /^Teams relay error: /);
    assert.ok(Date.now() - t0 < 3000, `took ${Date.now() - t0} ms`);
    await (handle as Awaited<ReturnType<typeof registerTeamsTools>>).close();
  } finally {
    await relay.close();
  }
});
