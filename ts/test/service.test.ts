import assert from "node:assert/strict";
import { test } from "node:test";
import { ServiceOutlook, httpGateway, sshGateway } from "../src/service.ts";
test("service resolves exact or unique account and sends no credentials", async () => {
  const calls: unknown[] = [];
  const o = new ServiceOutlook(async (p, b) => {
    calls.push([p, b]);
    return p.endsWith("accounts")
      ? { accounts: [{ upn: "a@msc.example" }, { upn: "b@msc.example" }] }
      : { value: [{ Id: "one" }] };
  });
  assert.equal((await o.account("a@msc.example")).upn, "a@msc.example");
  await assert.rejects(o.account("msc"), /ambiguous/);
  const r = await o.request("a@msc.example", "me/messages?$top=1");
  assert.equal(r.value.length, 1);
  assert.deepEqual(calls.at(-1), [
    "/api/communication/read",
    { account: "a@msc.example", path: "me/messages?$top=1" },
  ]);
  await assert.rejects(
    o.request("a@msc.example", "me/messages", "POST", {}),
    /read-only/,
  );
});
test("gateway configurations reject remote HTTP, credentials and SSH injection", () => {
  assert.throws(() => httpGateway("http://remote.example", "x"));
  assert.throws(() => httpGateway("https://u:p@remote.example", "x"));
  assert.throws(() => sshGateway("-oProxyCommand=x"));
  assert.throws(() => sshGateway("host", "c;whoami"));
});

test("http gateway keeps its fixed endpoint and bearer key, refuses redirects", async () => {
  let seen: RequestInit | undefined;
  const gateway = httpGateway(
    "http://127.0.0.1:3180",
    "key",
    async (_u, init) => {
      seen = init;
      return new Response('{"value":[]}', { status: 200 });
    },
  );
  await gateway("/api/communication/read", {
    account: "a@msc.example",
    path: "me/messages",
  });
  assert.equal(seen?.redirect, "error");
  assert.equal(new Headers(seen?.headers).get("Authorization"), "Bearer key");
  await assert.rejects(gateway("https://other.example", {}), /Unknown/);
});

test("gateway keeps reauthentication error and wraps invalid JSON", async () => {
  const denied = httpGateway(
    "http://localhost",
    "k",
    async () => new Response('{"errore":"Accedi di nuovo"}', { status: 409 }),
  );
  await assert.rejects(
    denied("/api/communication/accounts"),
    /Accedi di nuovo/,
  );
  const bad = httpGateway(
    "http://localhost",
    "k",
    async () => new Response("<html>"),
  );
  await assert.rejects(bad("/api/communication/accounts"), /invalid JSON/);
});
test("email-like fragment resolves once instead of pretending exact UPN", async () => {
  const calls: unknown[] = [];
  const o = new ServiceOutlook(async (p, b) => {
    calls.push([p, b]);
    return p.endsWith("accounts")
      ? { accounts: [{ upn: "a@msc.example" }] }
      : { value: [] };
  });
  await o.request("a@msc", "me/messages");
  await o.request("a@msc.example", "me/messages?$skip=1");
  assert.equal(calls.filter((x: any) => x[0].endsWith("accounts")).length, 1);
  assert.equal((calls[1] as any)[1].account, "a@msc.example");
});
