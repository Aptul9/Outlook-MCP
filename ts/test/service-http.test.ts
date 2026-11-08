import assert from "node:assert/strict";
import { test } from "node:test";
import http from "node:http";
import { ServiceOutlook, httpGateway } from "../src/service.ts";
import { listEmails } from "../src/mail.ts";

test("HTTP service path end to end: metadata then read, no Microsoft token exported", async () => {
  const bodies: unknown[] = [];
  const server = http.createServer(async (req, res) => {
    assert.equal(req.headers.authorization, "Bearer local-key");
    res.setHeader("Content-Type", "application/json");
    if (req.url === "/api/communication/accounts")
      return res.end(
        JSON.stringify({ accounts: [{ upn: "a@msc.example", status: "ok" }] }),
      );
    let body = "";
    for await (const c of req) body += c;
    const b = JSON.parse(body);
    bodies.push(b);
    assert.equal(b.account, "a@msc.example");
    assert.match(b.path, /^me\/messages/);
    res.end(JSON.stringify({ value: [{ Id: "1", Subject: "fixture" }] }));
  });
  await new Promise<void>((r) => server.listen(0, "127.0.0.1", r));
  try {
    const port = (server.address() as { port: number }).port;
    const o = new ServiceOutlook(
      httpGateway(`http://127.0.0.1:${port}`, "local-key"),
    );
    const rows = await listEmails(o, "msc", {
      folder: "all",
      unread_only: true,
    });
    assert.equal(rows.length, 1);
    assert.equal(bodies.length, 1);
  } finally {
    await new Promise<void>((r, j) => server.close((e) => (e ? j(e) : r())));
  }
});
