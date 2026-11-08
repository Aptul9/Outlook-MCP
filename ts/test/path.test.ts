import assert from "node:assert/strict";
import { test } from "node:test";
import { Outlook } from "../src/api.ts";
import { Tokens } from "../src/tokens.ts";
test("core transport refuses absolute or non-me paths before loading credentials", async () => {
  let loaded = 0;
  const o = new Outlook(
    new Tokens(async () => {
      loaded++;
      return { account: [] };
    }),
  );
  for (const p of [
    "https://evil.example",
    "../users",
    "users/x",
    "me\\messages",
  ])
    await assert.rejects(
      o.request(undefined, p),
      /Invalid Outlook request path/,
    );
  assert.equal(loaded, 0);
});
