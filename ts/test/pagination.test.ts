import assert from "node:assert/strict";
import { test } from "node:test";
import type { Obj } from "../src/api.ts";
import { listEmails, listEmailsPage, matchesLocalFilters, nextLinkPath, readThread, searchEmailsPage } from "../src/mail.ts";
import { fake, type Call } from "./helpers.ts";

const REST = "https://outlook.office.com/api/v2.0/";

const msg = (i: number, over: Obj = {}): Obj => ({
  Id: `m${i}`,
  ReceivedDateTime: new Date(Date.UTC(2026, 10, 8) - i * 60_000).toISOString(),
  From: { EmailAddress: { Name: `User ${i}`, Address: `user${i}@x.it` } },
  IsRead: true,
  ...over,
});
const ids = (r: { items: Obj[] }) => r.items.map((m) => m.Id);

/** Serves `all` newest first in pages of the requested $top, with a $skip nextLink on `linkPath` while mail remains. */
function pager(all: Obj[], linkPath = "me/messages") {
  return (c: Call) => {
    const u = new URL(`https://h/${c.url}`);
    const top = Number(u.searchParams.get("$top"));
    const end = Number(u.searchParams.get("$skip") ?? 0) + top;
    return { value: all.slice(end - top, end), ...(end < all.length ? { "@odata.nextLink": `${REST}${linkPath}?%24top=${top}&%24skip=${end}` } : {}) };
  };
}

test("unfiltered: exact limit across pages, clamp at 1000, has_more only while mail remains", async () => {
  const all = Array.from({ length: 1200 }, (_, i) => msg(i));
  const { o, calls } = fake({ "GET me/messages": pager(all) });

  const r = await listEmailsPage(o, undefined, { folder: "all", limit: 250 });
  assert.deepEqual(ids(r), all.slice(0, 250).map((m) => m.Id));
  assert.deepEqual([r.limit, r.pages, r.scanned, r.has_more, r.incomplete, r.stop_reason, calls.length], [250, 3, 300, true, false, "limit", 3]);
  assert.match(calls[0]!.url, /^me\/messages\?\$top=100&/);
  assert.match(calls[1]!.url, /^me\/messages\?%24top=100&%24skip=100$/);

  const big = await listEmailsPage(o, undefined, { folder: "all", limit: 5000 });
  assert.deepEqual([big.limit, big.items.length, new Set(ids(big)).size, big.pages, big.has_more], [1000, 1000, 1000, 10, true]);

  const seven = await listEmailsPage(o, undefined, { folder: "all", limit: 7 });
  assert.deepEqual([ids(seven).length, seven.pages, seven.has_more], [7, 1, true]);
  assert.match(calls.at(-1)!.url, /\$top=7&/);

  assert.deepEqual((await listEmails(o, undefined, { folder: "all", limit: 250 })).map((m) => m.Id), ids(r));

  const exact = fake({ "GET me/messages": pager(all.slice(0, 1000)) });
  const full = await listEmailsPage(exact.o, undefined, { folder: "all", limit: 1000 });
  assert.deepEqual([full.items.length, full.pages, full.has_more, full.incomplete, full.stop_reason], [1000, 10, false, false, "limit"]);

  const few = fake({ "GET me/messages": pager(all.slice(0, 30)) });
  const short = await listEmailsPage(few.o, undefined, { folder: "all", limit: 50 });
  assert.deepEqual([short.items.length, short.pages, short.has_more, short.incomplete, short.stop_reason], [30, 1, false, false, "exhausted"]);
  assert.match(few.calls[0]!.url, /\$top=50&/);
});

test("filtered: local filters page through older mail until limit, end, or the page cap", async () => {
  const boss = { From: { EmailAddress: { Name: "Boss", Address: "boss@y.it" } } };
  const all = Array.from({ length: 1000 }, (_, i) => msg(i, [150, 420, 999].includes(i) ? boss : {}));
  const { o, calls } = fake({ "GET me/MailFolders": pager(all, "me/MailFolders('Inbox')/messages") });

  const three = await listEmailsPage(o, undefined, { from_address: "BOSS", limit: 3 });
  assert.deepEqual(ids(three), ["m150", "m420", "m999"]);
  assert.deepEqual([three.pages, three.scanned, three.has_more, three.incomplete, three.stop_reason], [10, 1000, false, false, "limit"]);
  assert.match(calls[0]!.url, /^me\/MailFolders\/Inbox\/messages\?\$top=100&/);
  assert.match(calls[1]!.url, /^me\/MailFolders\('Inbox'\)\/messages\?%24top=100&%24skip=100$/);

  const two = await listEmailsPage(o, undefined, { from_address: "boss", limit: 2 });
  assert.deepEqual([ids(two), two.pages, two.has_more, two.incomplete], [["m150", "m420"], 5, true, false]);

  const all3 = await listEmailsPage(o, undefined, { from_address: "boss", limit: 50 });
  assert.deepEqual([ids(all3).length, all3.pages, all3.has_more, all3.incomplete, all3.stop_reason], [3, 10, false, false, "exhausted"]);

  const capped = await listEmailsPage(o, undefined, { from_address: "boss", limit: 50, max_pages: 3 });
  assert.deepEqual([ids(capped), capped.pages, capped.has_more, capped.incomplete, capped.stop_reason], [["m150"], 3, true, true, "max_pages"]);

  const huge = Array.from({ length: 6000 }, (_, i) => msg(i));
  const h = fake({ "GET me/messages": pager(huge) });
  const dflt = await listEmailsPage(h.o, undefined, { folder: "all", flagged_only: true });
  assert.deepEqual([dflt.items.length, dflt.pages, dflt.scanned, dflt.incomplete, dflt.stop_reason], [0, 20, 2000, true, "max_pages"]);
  const ceiling = await listEmailsPage(h.o, undefined, { folder: "all", flagged_only: true, max_pages: 999 });
  assert.deepEqual([ceiling.pages, ceiling.incomplete], [50, true]);
});

test("search: query plus local filter pages on, has_more from matches left in the last page", async () => {
  const all = Array.from({ length: 300 }, (_, i) => msg(i, [250, 260, 270].includes(i) ? { IsRead: false } : {}));
  const { o, calls } = fake({ "GET me/messages": pager(all) });
  const r = await searchEmailsPage(o, undefined, { query: 'q"x', unread_only: true, limit: 2 });
  assert.deepEqual([ids(r), r.pages, r.has_more, r.incomplete, r.stop_reason], [["m250", "m260"], 3, true, false, "limit"]);
  assert.match(decodeURIComponent(calls[0]!.url), /\$search="qx"/);
  await assert.rejects(searchEmailsPage(o, undefined, {}), /query or from_address is required/);
});

test("nextLink: off-origin or off-collection links are refused and never requested", async () => {
  const { o, calls } = fake({ "GET me/messages": () => ({ value: [msg(0)], "@odata.nextLink": "https://evil.example/api/v2.0/me/messages?%24skip=1" }) });
  const r = await listEmailsPage(o, undefined, { folder: "all", from_address: "boss" });
  assert.deepEqual([r.items.length, r.pages, r.has_more, r.incomplete, r.stop_reason, calls.length], [0, 1, true, true, "untrusted_next_link", 1]);

  const inbox = "me/MailFolders/Inbox/messages";
  const ok: [string, string, string][] = [
    [`${REST}me/messages?%24top=100&%24skip=100`, "me/messages", "me/messages?%24top=100&%24skip=100"],
    ["https://OUTLOOK.office.com:443/api/v2.0/me/Messages/?%24skip=1#x", "me/messages", "me/Messages/?%24skip=1"],
    [`${REST}me/MailFolders('Inbox')/messages?%24skip=10`, inbox, "me/MailFolders('Inbox')/messages?%24skip=10"],
    [`${REST}me/MailFolders('AAMk%2Fa%2Bb%3D')/messages?%24skip=10`, "me/MailFolders/AAMk%2Fa%2Bb%3D/messages", "me/MailFolders('AAMk%2Fa%2Bb%3D')/messages?%24skip=10"],
  ];
  for (const [link, base, want] of ok) assert.equal(nextLinkPath(link, base), want, link);

  const bad: unknown[] = [
    "http://outlook.office.com/api/v2.0/me/messages",
    "https://evil.example/api/v2.0/me/messages",
    "https://outlook.office.com.evil.example/api/v2.0/me/messages",
    "https://outlook.office.com@evil.example/api/v2.0/me/messages",
    "https://u:p@outlook.office.com/api/v2.0/me/messages",
    "https://outlook.office.com:8443/api/v2.0/me/messages",
    "https://outlook.office.com\\@evil.example/api/v2.0/me/messages",
    "https://outlook.office.com/api/v1.0/me/messages",
    "https://outlook.office.com/api/v2.0/../beta/me/messages",
    "https://outlook.office.com/api/v2.0/%2e%2e/beta/me/messages",
    "https://outlook.office.com/api/v2.0/me/sendmail",
    "https://outlook.office.com/api/v2.0/me%2Fmessages",
    "https://outlook.office.com/api/v2.0/me/messages%2F..%2Fsendmail",
    "https://outlook.office.com/api/v2.0/me/messages/AAMk1/attachments",
    "https://outlook.office.com/api/v2.0/users/boss@y.it/messages",
    "https://outlook.office.com/api/v2.0/me/MailFolders/SentItems/messages",
    "https://outlook.office.com/api/v2.0/me/messages/%E0%A4%A",
    "/api/v2.0/me/messages?%24skip=1",
    "me/messages?%24skip=1",
    "javascript:alert(1)",
    "",
    42,
    null,
    { href: `${REST}me/messages` },
  ];
  for (const link of bad) assert.equal(nextLinkPath(link, "me/messages"), null, String(link));
  assert.equal(nextLinkPath(`${REST}me/messages?%24skip=1`, inbox), null);
  assert.equal(nextLinkPath(`${REST}me/MailFolders('SentItems')/messages`, inbox), null);
});

test("nextLink: folder ids keep their case, path words and well-known names do not", () => {
  const base = "me/MailFolders/AAMkAbC%3D/messages";
  assert.equal(nextLinkPath(`${REST}ME/mailfolders('AAMkAbC%3D')/MESSAGES?%24skip=1`, base), "ME/mailfolders('AAMkAbC%3D')/MESSAGES?%24skip=1");
  assert.equal(nextLinkPath(`${REST}me/MailFolders/AAMkAbC%3D/messages?%24skip=1`, base), "me/MailFolders/AAMkAbC%3D/messages?%24skip=1");
  for (const other of ["AAMkABC%3D", "aamkabc%3D", "AAMKAbC%3D"]) {
    assert.equal(nextLinkPath(`${REST}me/MailFolders('${other}')/messages?%24skip=1`, base), null, other);
    assert.equal(nextLinkPath(`${REST}me/MailFolders/${other}/messages?%24skip=1`, base), null, other);
  }
  assert.equal(nextLinkPath(`${REST}me/mailfolders('inbox')/messages?%24skip=1`, "me/MailFolders/Inbox/messages"), "me/mailfolders('inbox')/messages?%24skip=1");
  assert.equal(nextLinkPath(`${REST}me/MailFolders/SENTITEMS/messages`, "me/MailFolders/SentItems/messages"), "me/MailFolders/SENTITEMS/messages");
});

test("nextLink: a repeating link stops paging, duplicate ids are dropped", async () => {
  const { o, calls } = fake({ "GET me/messages": () => ({ value: [msg(0), msg(1)], "@odata.nextLink": `${REST}me/messages?%24skip=2` }) });
  const r = await listEmailsPage(o, undefined, { folder: "all", limit: 10 });
  assert.deepEqual([ids(r), r.pages, r.has_more, r.incomplete, r.stop_reason, calls.length], [["m0", "m1"], 2, true, true, "repeated_next_link", 2]);
});

test("unknown folder is an error instead of a silent whole-mailbox list", async () => {
  const { o, calls } = fake({
    "GET me/MailFolders?": () => ({ value: [{ Id: "AAMkF1", DisplayName: "Clienti", ChildFolderCount: 0 }] }),
    "GET me/MailFolders/AAMkF1/messages": () => ({ value: [msg(0)] }),
  });
  await assert.rejects(listEmails(o, undefined, { folder: "Fornitori" }), /Folder not found: Fornitori/);
  assert.ok(calls.every((c) => !c.url.includes("messages")));
  assert.deepEqual((await listEmails(o, undefined, { folder: "clienti" })).map((m) => m.Id), ["m0"]);
});

test("date bounds with an offset compare as instants", () => {
  const start = { start_date: "2026-11-01T00:00:00+01:00" };
  assert.equal(matchesLocalFilters({ ReceivedDateTime: "2026-10-31T23:30:00Z" }, start), true);
  assert.equal(matchesLocalFilters({ ReceivedDateTime: "2026-10-31T22:30:00Z" }, start), false);
  assert.equal(matchesLocalFilters({ ReceivedDateTime: "2026-10-31T23:30:00Z" }, { end_date: "2026-11-01T00:00:00+01:00" }), false);
  assert.equal(matchesLocalFilters({ ReceivedDateTime: "2026-11-01T23:59:59Z" }, { end_date: "2026-11-01" }), true);
});

test("readThread follows nextLink pages and sorts oldest first", async () => {
  const { o, calls } = fake({
    "GET me/messages?": (c) =>
      c.url.includes("skip")
        ? { value: [{ Id: "1", ReceivedDateTime: "2026-11-01T00:00:00Z" }] }
        : { value: [{ Id: "3", ReceivedDateTime: "2026-11-03T00:00:00Z" }, { Id: "2", ReceivedDateTime: "2026-11-02T00:00:00Z" }], "@odata.nextLink": `${REST}me/messages?%24skip=2` },
  });
  assert.deepEqual((await readThread(o, undefined, "AAQk1")).map((m) => m.Id), ["1", "2", "3"]);
  assert.equal(calls.length, 2);
});
