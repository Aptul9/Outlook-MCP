import assert from "node:assert/strict";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/client";
import { StdioClientTransport } from "@modelcontextprotocol/client/stdio";
import http from "node:http";
import path from "node:path";

test("real stdio MCP service mode lists read tools only and reads unread fixture", async () => {
  const server = http.createServer(async (req, res) => {
    res.setHeader("Content-Type", "application/json");
    if (req.url === "/api/communication/accounts")
      return res.end(
        JSON.stringify({ accounts: [{ upn: "a@msc.example", status: "ok" }] }),
      );
    res.end(
      JSON.stringify({
        value: [{ Id: "fixture", Subject: "test", IsRead: false }],
      }),
    );
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  const port = (server.address() as { port: number }).port;
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [path.resolve("src/main.ts")],
    env: {
      ...Object.fromEntries(
        Object.entries(process.env).filter(
          (e): e is [string, string] => typeof e[1] === "string",
        ),
      ),
      OUTLOOK_MCP_SERVICE_URL: `http://127.0.0.1:${port}`,
      OUTLOOK_MCP_BROKER_KEY: "key",
    },
  });
  const client = new Client({ name: "test", version: "1" });
  try {
    await client.connect(transport);
    const tools = await client.listTools();
    const names = tools.tools.map((t) => t.name);
    assert.ok(names.includes("list_unread"));
    assert.ok(!names.includes("send_email"));
    const r = await client.callTool({
      name: "list_unread",
      arguments: { limit: 10 },
    });
    assert.ok(!r.isError);
    assert.match(JSON.stringify(r), /fixture/);
  } finally {
    await client.close();
    await transport.close();
    await new Promise<void>((r, j) => server.close((e) => (e ? j(e) : r())));
  }
});
