import { Outlook } from "../src/api.ts";
import { Tokens } from "../src/tokens.ts";

export const NOW = 1_800_000_000;

export function jwt(over: Record<string, unknown> = {}): string {
  const c = { upn: "mario@msc.example", aud: "https://outlook.office.com", exp: NOW + 3600, scp: "Mail.ReadWrite Calendars.ReadWrite", ...over };
  return `e30.${Buffer.from(JSON.stringify(c)).toString("base64url")}.sig`;
}

export type Call = { url: string; method: string; headers: Record<string, string>; body: any };

/** Fake Outlook REST: routes keyed by "METHOD path-prefix", path without the base URL. */
export function fake(routes: Record<string, (c: Call) => unknown>, tokens = [jwt()]) {
  const calls: Call[] = [];
  let i = 0;
  const t = new Tokens(async () => ({ account: [{ upn: "mario@msc.example", token: tokens[Math.min(i++, tokens.length - 1)], exp: NOW + 3600 }] }), () => NOW);
  const fetchFn = (async (input: string | URL, init?: RequestInit) => {
    const url = String(input).replace("https://outlook.office.com/api/v2.0/", "");
    const c: Call = { url, method: init?.method ?? "GET", headers: Object.fromEntries(new Headers(init?.headers).entries()), body: init?.body ? JSON.parse(String(init.body)) : undefined };
    calls.push(c);
    const key = Object.keys(routes).find((k) => {
      const [m, p] = k.split(" ");
      return m === c.method && url.startsWith(p!);
    });
    if (!key) return new Response("not found", { status: 404 });
    const out = routes[key]!(c);
    if (out instanceof Response) return out;
    return new Response(out === undefined ? null : JSON.stringify(out), { status: out === undefined ? 202 : 200 });
  }) as typeof fetch;
  return { o: new Outlook(t, fetchFn), calls };
}
