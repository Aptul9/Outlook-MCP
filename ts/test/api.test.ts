import assert from "node:assert/strict";
import { test } from "node:test";
import { ApiError } from "../src/api.ts";
import { listCalendarEvents } from "../src/calendar.ts";
import { htmlToMarkdown, formatBodyForHtml } from "../src/html.ts";
import { listEmails, readThread } from "../src/mail.ts";
import { manageEmail, parseRecipients, sendEmail } from "../src/send.ts";
import { fake, jwt, NOW } from "./helpers.ts";

test("htmlToMarkdown: links unwrapped, tables, quotes, entities, no scripts", () => {
  const md = htmlToMarkdown(
    '<style>x{}</style><p>Ciao&nbsp;<b>Mario</b> &amp; co</p><a href="https://eur01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fa.example%2Fp&data=1">doc</a><table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>x|y</td></tr></table><blockquote>old<br>text</blockquote>',
  );
  assert.match(md, /^Ciao \*\*Mario\*\* & co/);
  assert.match(md, /\[doc\]\(https:\/\/a\.example\/p\)/);
  assert.match(md, /\| A \| B \|\n\| --- \| --- \|\n\| 1 \| x\\\|y \|/);
  assert.match(md, /> old\n> text/);
  assert.doesNotMatch(md, /x\{\}/);
  assert.equal(formatBodyForHtml("a < b & c\nd"), '<div style="font-family:Aptos,Calibri,sans-serif;font-size:11pt;">a &lt; b &amp; c<br>\nd</div>');
  assert.equal(formatBodyForHtml("<p>x</p>"), "<p>x</p>");
});

test("parseRecipients: separators, display names, objects, duplicates", () => {
  assert.deepEqual(parseRecipients("a@x.it; B@x.it, a@x.it").map((r) => r.EmailAddress.Address), ["a@x.it", "B@x.it"]);
  assert.deepEqual(parseRecipients(['"Mario Rossi" <m@x.it>', { address: "l@x.it", name: "L" }]), [
    { EmailAddress: { Address: "m@x.it", Name: "Mario Rossi" } },
    { EmailAddress: { Address: "l@x.it", Name: "L" } },
  ]);
});

test("401 reloads the token from the broker once and retries", async () => {
  const second = jwt({ exp: NOW + 7200 });
  const { o, calls } = fake({ "GET me/MailFolders/Inbox/messages": (c) => (c.headers.authorization === `Bearer ${second}` ? { value: [{ Id: "1" }] } : new Response("", { status: 401 })) }, [jwt(), second]);
  assert.deepEqual(await listEmails(o, undefined, {}), [{ Id: "1" }]);
  assert.equal(calls.length, 2);
  const { o: o2 } = fake({ "GET me/": () => new Response("secret body".repeat(100), { status: 500 }) });
  await assert.rejects(listEmails(o2, undefined, { folder: "all" }), (e: Error) => e instanceof ApiError && e.message.length < 400);
});

test("listEmails: filter and order without query, search plus local filters with query", async () => {
  const { o, calls } = fake({
    "GET me/MailFolders/Inbox/messages": () => ({ value: [] }),
    "GET me/messages?": () => ({ value: [{ Id: "a", IsRead: true, From: { EmailAddress: { Address: "x@y.it" } } }, { Id: "b", IsRead: false, From: { EmailAddress: { Address: "boss@y.it" } } }] }),
  });
  await listEmails(o, undefined, { unread_only: true, start_date: "2026-11-01" });
  const u = decodeURIComponent(calls[0]!.url);
  assert.match(u, /\$filter=IsRead eq false and ReceivedDateTime ge 2026-11-01T00:00:00Z/);
  assert.match(u, /\$orderby=ReceivedDateTime desc/);
  const hits = await listEmails(o, undefined, { folder: "all", query: 'q"x', unread_only: true, from_address: "boss" });
  assert.deepEqual(hits.map((m) => m.Id), ["b"]);
  assert.match(decodeURIComponent(calls[1]!.url), /\$search="qx"/);
});

test("readThread: conversation id from a message id, quoted, oldest first", async () => {
  const { o, calls } = fake({
    "GET me/messages/m1?": () => ({ ConversationId: "AAQk'1" }),
    "GET me/messages?": () => ({ value: [{ Id: "2", ReceivedDateTime: "2026-11-02T00:00:00Z" }, { Id: "1", ReceivedDateTime: "2026-11-01T00:00:00Z" }] }),
  });
  assert.deepEqual((await readThread(o, undefined, "m1")).map((m) => m.Id), ["1", "2"]);
  assert.match(decodeURIComponent(calls[1]!.url), /ConversationId eq 'AAQk''1'/);
});

test("reply-all: thread Cc kept plus extra, excluded dropped, sent; failure deletes the draft", async () => {
  const draft = { Id: "d1", Subject: "RE: x", ToRecipients: [{ EmailAddress: { Address: "boss@y.it" } }], CcRecipients: [{ EmailAddress: { Address: "me@y.it" } }, { EmailAddress: { Address: "c@y.it" } }] };
  const { o, calls } = fake({ "POST me/messages/m1/createreplyall": () => draft, "PATCH me/messages/d1": () => ({}), "POST me/messages/d1/send": () => undefined });
  const r = await sendEmail(o, undefined, { reply_to_id: "m1", reply_all: true, body: "ok", cc: "new@y.it", exclude_recipients: "me@y.it" });
  assert.deepEqual([r.status, r.to, r.cc], ["sent", ["boss@y.it"], ["c@y.it", "new@y.it"]]);
  assert.match(calls[0]!.body.Comment, /^<div style=/);
  assert.deepEqual(calls.map((c) => `${c.method} ${c.url}`), ["POST me/messages/m1/createreplyall", "PATCH me/messages/d1", "POST me/messages/d1/send"]);
  const f = fake({ "POST me/messages/m1/createreply": () => draft, "PATCH me/messages/d1": () => new Response("no", { status: 400 }), "DELETE me/messages/d1": () => undefined });
  await assert.rejects(sendEmail(f.o, undefined, { reply_to_id: "m1", body: "ok" }), ApiError);
  assert.equal(f.calls.at(-1)!.method, "DELETE");
});

test("new email: sendmail with Text body; draft goes to me/messages; missing attachment stops before any request", async () => {
  const { o, calls } = fake({ "POST me/sendmail": () => undefined, "POST me/messages": () => ({ Id: "d9" }) });
  await sendEmail(o, undefined, { to: "a@x.it", subject: "s", body: "a < b" });
  assert.equal(calls[0]!.body.Message.Body.ContentType, "Text");
  assert.equal(calls[0]!.body.SaveToSentItems, "true");
  const d = await sendEmail(o, undefined, { to: "a@x.it", subject: "s", body: "<p>x</p>", save_as_draft: true });
  assert.deepEqual([d.status, d.draft_id, calls[1]!.body.Body.ContentType], ["draft", "d9", "HTML"]);
  await assert.rejects(sendEmail(o, undefined, { to: "a@x.it", body: "x", attachments: ["C:/no/such/file.pdf"] }), /not found/);
  assert.equal(calls.length, 2);
});

test("manageEmail: patches, moves to well-known and named folders, refuses unknown actions", async () => {
  const { o, calls } = fake({
    "PATCH me/messages/": () => ({}),
    "POST me/messages/": (c) => ({ Id: `new-${c.body.DestinationId}` }),
    "GET me/MailFolders?": () => ({ value: [{ Id: "AAMkF1", DisplayName: "Clienti", ChildFolderCount: 0 }] }),
  });
  assert.deepEqual(await manageEmail(o, undefined, ["a", "b"], "mark_read"), [{ id: "a", status: "marked_read" }, { id: "b", status: "marked_read" }]);
  assert.equal((await manageEmail(o, undefined, "a", "delete"))[0]!.status, "moved_to_DeletedItems");
  assert.equal((await manageEmail(o, undefined, "a", "move", "clienti"))[0]!.new_id, "new-AAMkF1");
  await assert.rejects(manageEmail(o, undefined, "a", "purge"), /Unsupported action/);
  await assert.rejects(manageEmail(o, undefined, "a", "move", "nowhere"), /not found/);
  assert.deepEqual(calls[0]!.body, { IsRead: true });
});

test("calendar view: zone required, bounded range and size, Prefer time zone, has_more", async () => {
  const { o, calls } = fake({ "GET me/calendars/c%2F1/calendarview": () => ({ value: [{ Id: "e1" }], "@odata.nextLink": "https://elsewhere.example" }) });
  const r = await listCalendarEvents(o, undefined, "2026-11-09T00:00:00+01:00", "2026-11-16T00:00:00Z", { calendar_id: "c/1", limit: 5000, time_zone: "Europe/Rome" });
  assert.deepEqual(r, { events: [{ Id: "e1" }], has_more: true });
  const u = new URL(`https://h/${calls[0]!.url}`);
  assert.equal(u.searchParams.get("startDateTime"), "2026-11-08T23:00:00.000Z");
  assert.equal(u.searchParams.get("$top"), "1000");
  assert.equal(calls[0]!.headers.prefer, 'outlook.timezone="Europe/Rome"');
  await assert.rejects(listCalendarEvents(o, undefined, "2026-11-09T00:00:00", "2026-11-10T00:00:00Z"), /time zone/);
  await assert.rejects(listCalendarEvents(o, undefined, "2026-11-10T00:00:00Z", "2026-11-09T00:00:00Z"), /after start/);
  await assert.rejects(listCalendarEvents(o, undefined, "2025-01-01T00:00:00Z", "2026-11-09T00:00:00Z"), /366 days/);
});

test('missing reply attachment creates no draft',async()=>{
 const {o,calls}=fake({});
 await assert.rejects(sendEmail(o,undefined,{reply_to_id:'m1',body:'text',attachments:['C:/no/such/attachment.pdf']}),/not found/);
 assert.equal(calls.length,0);
});
