import assert from "node:assert/strict";
import { test } from "node:test";
import { spawnSync } from "node:child_process";
test("SSH inline module form supports top-level await and stdin", () => {
  const encoded = Buffer.from(
    "let s='';for await(const c of process.stdin)s+=c;process.stdout.write('ok:'+s)",
  ).toString("base64");
  const r = spawnSync(
    process.execPath,
    [
      "--input-type=module",
      "-e",
      `await import('data:text/javascript;base64,${encoded}')`,
    ],
    { input: "hello", encoding: "utf8" },
  );
  assert.equal(r.status, 0, r.stderr);
  assert.equal(r.stdout, "ok:hello");
});
