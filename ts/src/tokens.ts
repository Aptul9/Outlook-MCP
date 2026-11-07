// Access tokens for Outlook REST from the OE Copilot token broker (GET /api/token-posta). OE Copilot owns sign-in and renewal of every
// mailbox (Outlook Mobile, Outlook web session); this process only reads current access tokens, keeps them in memory, never on disk.

import { execFile } from "node:child_process";

export const OUTLOOK_REST = "https://outlook.office.com/api/v2.0";
const MARGIN_S = 120;

export interface Account {
  upn: string;
  token: string;
  exp: number;
  status: string;
  message: string;
}

export class TokenError extends Error {}

export function claims(jwt: string): Record<string, unknown> {
  const part = jwt.split(".")[1];
  if (!part) return {};
  try {
    const v = JSON.parse(Buffer.from(part, "base64url").toString("utf-8")) as unknown;
    return v !== null && typeof v === "object" && !Array.isArray(v) ? (v as Record<string, unknown>) : {};
  } catch {
    return {};
  }
}

/**
 * The broker rows that hold a usable Outlook token: audience, mail permission, the same mailbox as the row and time left are read from the
 * token itself, never from the row. The API address is fixed here, so a broker answer cannot send requests anywhere else.
 */
export function accountsFrom(payload: unknown, nowS: number): Account[] {
  const rows = payload !== null && typeof payload === "object" ? (payload as Record<string, unknown>).account : undefined;
  if (!Array.isArray(rows)) return [];
  const out: Account[] = [];
  for (const row of rows as unknown[]) {
    if (row === null || typeof row !== "object") continue;
    const r = row as Record<string, unknown>;
    if (typeof r.token !== "string" || typeof r.upn !== "string") continue;
    const c = claims(r.token);
    const upn = String(c.upn || c.preferred_username || c.unique_name || "").toLowerCase();
    if (!upn || upn !== r.upn.toLowerCase()) continue;
    if (c.aud !== "https://outlook.office.com" && c.aud !== "https://outlook.office365.com") continue;
    const scopes = String(c.scp ?? "").split(" ");
    if (!scopes.includes("Mail.Read") && !scopes.includes("Mail.ReadWrite")) continue;
    const exp = Math.min(Number(c.exp) || 0, Number(r.exp) || Number(c.exp) || 0);
    if (exp <= nowS + MARGIN_S) continue;
    if (out.some((a) => a.upn === upn)) continue;
    out.push({ upn, token: r.token, exp, status: String(r.phone_status ?? "ok"), message: String(r.phone_msg ?? "") });
  }
  return out;
}

type Loader = () => Promise<unknown>;

/** HTTP straight to the broker: https, or http on this machine only; no redirects, so the key goes nowhere else. */
export function httpLoader(url: string, key: string): Loader {
  const u = new URL(url);
  const local = ["127.0.0.1", "::1", "[::1]", "localhost"].includes(u.hostname);
  if (u.username || u.password || u.search || u.hash || !(u.protocol === "https:" || (u.protocol === "http:" && local))) {
    throw new TokenError("OUTLOOK_MCP_BROKER_URL must be https, or http on localhost, without credentials or query");
  }
  return async () => {
    const r = await fetch(u, { headers: { Authorization: `Bearer ${key}`, Accept: "application/json" }, redirect: "manual", signal: AbortSignal.timeout(20_000) });
    if (!r.ok) throw new TokenError(`Token broker answered HTTP ${r.status}`);
    return r.json();
  };
}

/**
 * Over SSH to the OE Copilot host: the request runs inside the agent container with its own API_KEY, so the key never leaves the host.
 * Arguments go to ssh as a list (no local shell); the script travels on stdin.
 */
export function sshLoader(host: string, container: string, port: number): Loader {
  if (!/^[\w.@-]+$/.test(host) || !/^[\w.-]+$/.test(container) || !Number.isInteger(port)) throw new TokenError("Invalid SSH host, container or port");
  const script = `const r=await fetch("http://127.0.0.1:${port}/api/token-posta",{headers:{Authorization:"Bearer "+process.env.API_KEY}});if(!r.ok){console.error("HTTP "+r.status);process.exit(2)}process.stdout.write(await r.text())`;
  return () =>
    new Promise((resolve, reject) => {
      const child = execFile(
        "ssh",
        ["-o", "BatchMode=yes", "-o", "ConnectTimeout=10", host, "docker", "exec", "-i", container, "node", "--input-type=module", "-"],
        { timeout: 45_000, maxBuffer: 4_000_000, windowsHide: true },
        (err, stdout) => {
          if (err) return reject(new TokenError(`Token broker over SSH failed (${host}): ${err.code ?? err.message}`));
          try {
            resolve(JSON.parse(stdout));
          } catch {
            reject(new TokenError("Token broker answer is not JSON"));
          }
        },
      );
      child.stdin?.end(script);
    });
}

export function loaderFromEnv(env: NodeJS.ProcessEnv = process.env): Loader {
  if (env.OUTLOOK_MCP_BROKER_URL) {
    if (!env.OUTLOOK_MCP_BROKER_KEY) throw new TokenError("OUTLOOK_MCP_BROKER_URL needs OUTLOOK_MCP_BROKER_KEY");
    return httpLoader(env.OUTLOOK_MCP_BROKER_URL, env.OUTLOOK_MCP_BROKER_KEY);
  }
  return sshLoader(env.OUTLOOK_MCP_SSH_HOST || "oecopilot-vm", env.OUTLOOK_MCP_AGENT_CONTAINER || "oe-copilot-agent", Number(env.OUTLOOK_MCP_AGENT_PORT || 8180));
}

export class Tokens {
  private cache: Account[] = [];
  private pending: Promise<Account[]> | null = null;
  private readonly load: Loader;
  private readonly nowS: () => number;
  constructor(load: Loader, nowS: () => number = () => Math.floor(Date.now() / 1000)) {
    this.load = load;
    this.nowS = nowS;
  }

  async accounts(refresh = false): Promise<Account[]> {
    const fresh = this.cache.filter((a) => a.exp > this.nowS() + MARGIN_S);
    if (!refresh && fresh.length && fresh.length === this.cache.length) return fresh;
    this.pending ??= this.load()
      .then((p) => (this.cache = accountsFrom(p, this.nowS())))
      .finally(() => (this.pending = null));
    return this.pending;
  }

  /** Exact address first, then a fragment matching one mailbox only ("msc", "serco"); omitted: the first mailbox. */
  async pick(account?: string, refresh = false): Promise<Account> {
    const all = await this.accounts(refresh);
    if (!all.length) throw new TokenError("No Outlook mailbox available from the OE Copilot token broker");
    if (!account?.trim()) return all[0]!;
    const want = account.trim().toLowerCase();
    const exact = all.find((a) => a.upn === want);
    if (exact) return exact;
    const hits = all.filter((a) => a.upn.includes(want));
    if (hits.length === 1) return hits[0]!;
    if (hits.length > 1) throw new TokenError(`'${account}' matches several mailboxes: ${hits.map((a) => a.upn).join(", ")}`);
    if (!refresh) return this.pick(account, true);
    throw new TokenError(`Mailbox '${account}' not found among: ${all.map((a) => a.upn).join(", ")}`);
  }
}
