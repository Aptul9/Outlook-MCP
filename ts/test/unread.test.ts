import assert from "node:assert/strict";
import { test } from "node:test";
import { Outlook } from "../src/api.ts";
import { Tokens } from "../src/tokens.ts";
import { listUnread } from "../src/unread.ts";
import { jwt, NOW } from "./helpers.ts";
test("unread parallelism bounded, one account failure does not hide others", async () => {
  const accounts = ["a", "b", "c", "d"].map((n) => ({
    upn: n + "@x.example",
    token: jwt({ upn: n + "@x.example" }),
    exp: NOW + 3600,
  }));
  let active = 0,
    max = 0;
  const o = new Outlook(
    new Tokens(
      async () => ({ account: accounts }),
      () => NOW,
    ),
    async (_u, init) => {
      active++;
      max = Math.max(max, active);
      await new Promise((r) => setTimeout(r, 5));
      active--;
      if (
        new Headers(init?.headers)
          .get("Authorization")
          ?.includes(accounts[1]!.token)
      )
        return new Response("", { status: 403 });
      return new Response(JSON.stringify({ value: [{ Id: "1" }] }));
    },
  );
  const r = await listUnread(o, { concurrency: 2 });
  assert.equal(r.length, 4);
  assert.equal(max, 2);
  assert.equal(r.filter((x) => x.error).length, 1);
  assert.equal(r.filter((x) => x.result).length, 3);
});

test('unknown unread selector isolated and metadata fetched only once',async()=>{
 let loads=0;const o=new Outlook(new Tokens(async()=>{loads++;return{account:[{upn:'a@x.example',token:jwt({upn:'a@x.example'}),exp:NOW+3600}]};},()=>NOW),async()=>new Response('{"value":[]}'));
 const r=await listUnread(o,{accounts:['a@x.example','missing']});assert.equal(loads,1);assert.equal(r.length,2);assert.equal(r.filter(x=>x.error).length,1);
});
