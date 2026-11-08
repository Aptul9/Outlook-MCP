import assert from "node:assert/strict";
import { test } from "node:test";
import { accountsFrom, httpLoader, sshLoader, Tokens, TokenError } from "../src/tokens.ts";
import { jwt, NOW } from "./helpers.ts";

const row = (over: Record<string, unknown> = {}, claims: Record<string, unknown> = {}) => ({ upn: String(claims.upn ?? "mario@msc.example"), token: jwt(claims), exp: NOW + 3600, phone_status: "ok", ...over });

test("accountsFrom keeps only Outlook mail tokens for the same mailbox, with time left", () => {
  const rows = [
    row(),
    row({}, { upn: "other@x.example", aud: "https://graph.microsoft.com" }),
    row({}, { upn: "cal@x.example", scp: "Calendars.Read" }),
    row({}, { upn: "old@x.example", exp: NOW + 60 }),
    row({ upn: "liar@x.example" }, { upn: "real@x.example" }),
    row({ exp: NOW + 100 }, { upn: "rowexp@x.example" }),
    row(),
    row({exp:0},{upn:"zero@x.example"}),
    row({exp:"bad"},{upn:"bad@x.example"}),
    { upn: 1, token: 2 },
    null,
  ];
  assert.deepEqual(accountsFrom({ account: rows }, NOW).map((a) => a.upn), ["mario@msc.example"]);
  assert.deepEqual(accountsFrom({ account: "x" }, NOW), []);
  assert.deepEqual(accountsFrom(null, NOW), []);
});

test("pick: exact, unique fragment, ambiguous, unknown after one reload", async () => {
  let loads = 0;
  const t = new Tokens(async () => (loads++, { account: [row(), row({}, { upn: "a@serco.example" }), row({}, { upn: "b@serco.example" })] }), () => NOW);
  assert.equal((await t.pick()).upn, "mario@msc.example");
  assert.equal((await t.pick("MSC")).upn, "mario@msc.example");
  assert.equal((await t.pick("a@serco.example")).upn, "a@serco.example");
  await assert.rejects(t.pick("serco"), /several mailboxes/);
  const before = loads;
  await assert.rejects(t.pick("nobody"), TokenError);
  assert.equal(loads, before + 1);
});

test("tokens are cached until one expires, and concurrent loads share one request", async () => {
  let now = NOW;
  let loads = 0;
  const t = new Tokens(async () => (loads++, { account: [row()] }), () => now);
  await Promise.all([t.accounts(), t.accounts(), t.accounts()]);
  assert.equal(loads, 1);
  await t.accounts();
  assert.equal(loads, 1);
  now = NOW + 3500;
  await t.accounts();
  assert.equal(loads, 2);
});

test("broker HTTP: https or localhost only, no credentials or query in the URL", () => {
  assert.throws(() => httpLoader("http://broker.example/api/token-posta", "k"), TokenError);
  assert.throws(() => httpLoader("https://u:p@broker.example/x", "k"), TokenError);
  assert.throws(() => httpLoader("https://broker.example/x?a=1", "k"), TokenError);
  httpLoader("https://broker.example/api/token-posta", "k");
  httpLoader("http://127.0.0.1:3180/api/token-posta", "k");
});

test("broker SSH: host and container names cannot carry options or shell syntax", () => {
  assert.throws(() => sshLoader("-oProxyCommand=x", "c", 8180), TokenError);
  assert.throws(() => sshLoader("host;rm", "c", 8180), TokenError);
  assert.throws(() => sshLoader("host", "c d", 8180), TokenError);
  sshLoader("oecopilot-vm", "oe-copilot-agent", 8180);
});
