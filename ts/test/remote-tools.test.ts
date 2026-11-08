// Remote tool surface: the server's disk is never the caller's. Attachments go out as base64 in the request and come back as
// content blocks; no local path is read and nothing is written.
import assert from "node:assert/strict";
import fs from "node:fs";
import { test } from "node:test";
import { Client } from "@modelcontextprotocol/client";
import { InMemoryTransport } from "@modelcontextprotocol/server";
import { createServer } from "../src/main.ts";
import { DOWNLOAD_DIR } from "../src/tools-write.ts";
import { fake } from "./helpers.ts";

async function connected(o: Parameters<typeof createServer>[0], remote: boolean) {
  const server = createServer(o, remote);
  const client = new Client({ name: "t", version: "1" });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(a), client.connect(b)]);
  return { client, close: async () => (await client.close(), await server.close()) };
}

test("remote send_email: base64 files attached, local paths not accepted", async () => {
  const { o, calls } = fake({ "POST me/sendmail": () => undefined });
  const { client, close } = await connected(o, true);
  try {
    const tools = (await client.listTools()).tools;
    const send = tools.find((t) => t.name === "send_email")!;
    const props = Object.keys((send.inputSchema as { properties: object }).properties);
    assert.ok(props.includes("files") && !props.includes("attachments"));
    const content = Buffer.from("hello attachment").toString("base64");
    const r = await client.callTool({ name: "send_email", arguments: { to: "a@x.it", subject: "s", body: "b", files: [{ name: "../../etc/note.txt", content_base64: content }] } });
    assert.ok(!r.isError, JSON.stringify(r));
    const att = calls[0]!.body.Message.Attachments[0];
    assert.deepEqual([att.Name, att.ContentBytes], ["note.txt", content]);
    const bad = await client.callTool({ name: "send_email", arguments: { to: "a@x.it", body: "b", files: [{ name: "x", content_base64: "not base64!" }] } });
    assert.ok(bad.isError);
    assert.equal(calls.length, 1, "malformed file stops before any request");
    // a local path smuggled in is dropped by the schema, never read
    await client.callTool({ name: "send_email", arguments: { to: "a@x.it", body: "b", attachments: ["C:/Windows/win.ini"] } });
    assert.equal(calls.at(-1)!.body.Message.Attachments, undefined);
  } finally {
    await close();
  }
});

test("remote get_attachment: bytes in the result, nothing on disk, no save_dir", async () => {
  const bytes = Buffer.from("%PDF-1.7 fixture").toString("base64");
  const { o } = fake({
    "GET me/messages/m1/attachments?": () => ({ value: [{ Id: "a1", Name: "offer.pdf", ContentType: "application/pdf", Size: 16 }] }),
    "GET me/messages/m1/attachments/a1": () => ({ "@odata.type": "#Microsoft.OutlookServices.FileAttachment", Name: "offer.pdf", ContentType: "application/pdf", ContentBytes: bytes }),
  });
  const before = fs.existsSync(DOWNLOAD_DIR) ? fs.readdirSync(DOWNLOAD_DIR).length : -1;
  const { client, close } = await connected(o, true);
  try {
    const tool = (await client.listTools()).tools.find((t) => t.name === "get_attachment")!;
    assert.ok(!Object.keys((tool.inputSchema as { properties: object }).properties).includes("save_dir"));
    const r = await client.callTool({ name: "get_attachment", arguments: { message_id: "m1" } });
    assert.ok(!r.isError, JSON.stringify(r));
    const blocks = r.content as { type: string; resource?: { blob: string; mimeType: string } }[];
    assert.equal(blocks[1]!.type, "resource");
    assert.deepEqual([blocks[1]!.resource!.blob, blocks[1]!.resource!.mimeType], [bytes, "application/pdf"]);
    assert.equal(fs.existsSync(DOWNLOAD_DIR) ? fs.readdirSync(DOWNLOAD_DIR).length : -1, before);
  } finally {
    await close();
  }
});

test("local mode keeps file-path attachments and save_dir", async () => {
  const { o } = fake({});
  const { client, close } = await connected(o, false);
  try {
    const tools = (await client.listTools()).tools;
    const props = (n: string) => Object.keys((tools.find((t) => t.name === n)!.inputSchema as { properties: object }).properties);
    assert.ok(props("send_email").includes("attachments") && !props("send_email").includes("files"));
    assert.ok(props("get_attachment").includes("save_dir"));
  } finally {
    await close();
  }
});
