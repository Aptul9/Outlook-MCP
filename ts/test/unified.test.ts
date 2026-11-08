import assert from "node:assert/strict";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import { createMcpHandler, McpServer } from "@modelcontextprotocol/server";
import http from "node:http";
import path from "node:path";

test("single stdio server discovers Teams alongside mail/calendar and calls both", async () => {
  const upstream = createMcpHandler(() => {
    const s = new McpServer({ name: "fixture-teams", version: "1" });
    s.registerTool(
      "list_chats",
      { description: "fixture chats", annotations: { readOnlyHint: true } },
      async () => ({ content: [{ type: "text", text: "chat-fixture" }] }),
    );
    return s;
  });
  const server = http.createServer(async (req, res) => {
    if (req.url === "/api/communication/accounts") {
      res.setHeader("Content-Type", "application/json");
      return res.end(JSON.stringify({ accounts: [{ upn: "a@msc.example" }] }));
    }
    if (req.url === "/api/communication/read") {
      res.setHeader("Content-Type", "application/json");
      return res.end(
        JSON.stringify({ value: [{ Id: "mail-fixture", Subject: "fixture" }] }),
      );
    }
    const chunks: Buffer[] = [];
    for await (const c of req) chunks.push(Buffer.from(c));
    const headers = new Headers();
    for (const [k, v] of Object.entries(req.headers))
      if (v) headers.set(k, String(v));
    const r = await upstream.fetch(
      new Request("http://127.0.0.1/mcp", {
        method: req.method,
        headers,
        body:
          req.method === "POST"
            ? new Uint8Array(Buffer.concat(chunks))
            : undefined,
      }),
    );
    res.writeHead(r.status, Object.fromEntries(r.headers));
    res.end(Buffer.from(await r.arrayBuffer()));
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  const port = (server.address() as { port: number }).port;
  const env = Object.fromEntries(
    Object.entries(process.env).filter(
      (e): e is [string, string] => typeof e[1] === "string",
    ),
  );
  const tr = new StdioClientTransport({
    command: process.execPath,
    args: [path.resolve("src/main.ts")],
    env: {
      ...env,
      OUTLOOK_MCP_SERVICE_URL: `http://127.0.0.1:${port}`,
      OUTLOOK_MCP_BROKER_KEY: "key",
      OUTLOOK_MCP_TEAMS_URL: `http://127.0.0.1:${port}/mcp`,
      OUTLOOK_MCP_TEAMS_TOKEN: "fixture",
    },
  });
  const c = new Client({ name: "integration", version: "1" });
  try {
    await c.connect(tr);
    const names = (await c.listTools()).tools.map((t) => t.name);
    assert.ok(names.includes("teams_list_chats"));
    assert.ok(names.includes("list_unread"));
    assert.ok(names.includes("list_calendars"));
    assert.match(
      JSON.stringify(
        await c.callTool({ name: "teams_list_chats", arguments: {} }),
      ),
      /chat-fixture/,
    );
    assert.match(
      JSON.stringify(await c.callTool({ name: "list_unread", arguments: {} })),
      /mail-fixture/,
    );
  } finally {
    await c.close();
    await tr.close();
    server.closeAllConnections();
    await new Promise<void>((r) => server.close(() => r()));
  }
});
