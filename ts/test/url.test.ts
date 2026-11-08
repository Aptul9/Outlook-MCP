import assert from "node:assert/strict";
import { test } from "node:test";
import { backendUrl } from "../src/url.ts";

test("backend URLs: https anywhere, http only on loopback or a container service name", () => {
  for (const ok of [
    "https://teams.example/mcp",
    "http://localhost:3180/x",
    "http://127.0.0.1:1",
    "http://[::1]:1",
    "http://agent:8180",
    "http://webapp:8090/mcp",
  ]) {
    assert.ok(backendUrl(ok), ok);
  }
  for (const bad of [
    "http://agent.example:8180",
    "http://10.0.0.5:8180",
    "http://1234/",
    "http://u:p@agent:8180",
    "http://agent:8180/?k=1",
    "http://agent:8180/#f",
    "ftp://agent",
    "not a url",
  ]) {
    assert.equal(backendUrl(bad), null, bad);
  }
});
