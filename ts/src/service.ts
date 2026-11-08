// Service boundary: OE-Copilot owns credentials; MCP sends only bounded read requests.
import { execFile } from "node:child_process";
import { Outlook, ApiError, type Obj } from "./api.ts";
import { Tokens, selectAccount, type Account } from "./tokens.ts";
export type GatewayCall = (path: string, body?: unknown) => Promise<Obj>;

export function httpGateway(
  base: string,
  key: string,
  fetchFn = fetch,
): GatewayCall {
  const u = new URL(base);
  if (
    u.username ||
    u.password ||
    u.search ||
    u.hash ||
    u.pathname !== "/" ||
    !(
      u.protocol === "https:" ||
      (u.protocol === "http:" &&
        ["localhost", "127.0.0.1", "[::1]"].includes(u.hostname))
    )
  )
    throw new Error("Invalid communication service URL");
  return async (p, body) => {
    if (!["/api/communication/accounts", "/api/communication/read"].includes(p))
      throw new ApiError("Unknown communication endpoint");
    const r = await fetchFn(new URL(p, u), {
      method: body === undefined ? "GET" : "POST",
      headers: {
        Authorization: `Bearer ${key}`,
        "Content-Type": "application/json",
      },
      body: body === undefined ? undefined : JSON.stringify(body),
      redirect: "error",
      signal: AbortSignal.timeout(45_000),
    });
    let payload: Obj;
    try {
      payload = (await r.json()) as Obj;
    } catch {
      throw new ApiError(
        `Communication service returned invalid JSON (HTTP ${r.status})`,
      );
    }
    if (!r.ok)
      throw new ApiError(
        `Communication service HTTP ${r.status}: ${String(payload.errore ?? "request failed").slice(0, 200)}`,
      );
    return payload;
  };
}

export function sshGateway(
  host: string,
  container = "oe-copilot-agent",
  port = 8180,
): GatewayCall {
  if (
    !/^[A-Za-z0-9_][\w.@-]*$/.test(host) ||
    !/^[A-Za-z0-9_][\w.-]*$/.test(container) ||
    !Number.isInteger(port) ||
    port < 1 ||
    port > 65535
  )
    throw new Error("Invalid communication SSH configuration");
  return (p, body) =>
    new Promise((resolve, reject) => {
      if (
        !["/api/communication/accounts", "/api/communication/read"].includes(p)
      )
        return reject(new ApiError("Unknown communication endpoint"));
      const input = JSON.stringify({ path: p, body });
      const script = `let s='';for await(const c of process.stdin)s+=c;const i=JSON.parse(s);const r=await fetch('http://127.0.0.1:${port}'+i.path,{method:i.body===undefined?'GET':'POST',headers:{Authorization:'Bearer '+process.env.API_KEY,'Content-Type':'application/json'},body:i.body===undefined?undefined:JSON.stringify(i.body),redirect:'error',signal:AbortSignal.timeout(30000)});process.stdout.write(JSON.stringify({status:r.status,body:await r.text()}));`;
      // Script is encoded, so the remote shell never interprets request data or credentials.
      const encoded = Buffer.from(script).toString("base64");
      const command = `node --input-type=module -e "await import('data:text/javascript;base64,${encoded}')"`;
      const child = execFile(
        "ssh",
        [
          "-o",
          "BatchMode=yes",
          "-o",
          "ConnectTimeout=10",
          host,
          `docker exec -i ${container} ${command}`,
        ],
        { timeout: 45_000, maxBuffer: 9_000_000, windowsHide: true },
        (err, stdout) => {
          if (err)
            return reject(
              new ApiError("Communication service SSH request failed"),
            );
          try {
            const envelope = JSON.parse(stdout);
            let body: Obj;
            try {
              body = JSON.parse(envelope.body);
            } catch {
              throw new ApiError("Communication service returned invalid JSON");
            }
            if (envelope.status < 200 || envelope.status >= 300)
              return reject(
                new ApiError(
                  `Communication service HTTP ${envelope.status}: ${String(body.errore ?? "request failed").slice(0, 200)}`,
                ),
              );
            resolve(body);
          } catch {
            reject(new ApiError("Communication service returned invalid JSON"));
          }
        },
      );
      child.stdin?.end(input);
    });
}

export class ServiceOutlook extends Outlook {
  private readonly gateway: GatewayCall;
  private knownAccounts = new Set<string>();
  constructor(gateway: GatewayCall) {
    // Metadata only. Never request or export Microsoft bearer tokens in service mode.
    super(new Tokens(async () => ({ account: [] })));
    this.gateway = gateway;
    this.tokens.accounts = async () => this.metadata();
    this.tokens.pick = async (account) => this.account(account);
  }
  private async metadata(): Promise<Account[]> {
    const r = await this.gateway("/api/communication/accounts");
    if (!Array.isArray(r.accounts))
      throw new ApiError("Invalid communication account list");
    const accounts: Account[] = r.accounts
      .filter(
        (a: unknown): a is Obj =>
          a !== null &&
          typeof a === "object" &&
          typeof (a as Obj).upn === "string" &&
          /^[^\s@]+@[^\s@]+$/.test((a as Obj).upn),
      )
      .map((a: Obj) => ({
        upn: String(a.upn).toLowerCase(),
        status: String(a.status ?? "ok"),
        message: String(a.message ?? ""),
        exp: 0,
        token: "",
      }));
    this.knownAccounts = new Set(accounts.map((a) => a.upn));
    return accounts;
  }
  override async account(account?: string): Promise<Account> {
    const all = await this.metadata();
    return selectAccount(all, account);
  }
  override async request(
    account: string | undefined,
    p: string,
    method = "GET",
    data?: unknown,
    _retry = true,
    extra?: Record<string, string>,
  ): Promise<Obj> {
    if (method !== "GET" || data !== undefined)
      throw new ApiError("Communication service mode supports read-only tools");
    const who =
      account && this.knownAccounts.has(account.toLowerCase())
        ? { upn: account.toLowerCase() }
        : await this.account(account);
    return this.gateway("/api/communication/read", {
      account: who.upn,
      path: p,
      ...(extra?.Prefer ? { prefer: extra.Prefer } : {}),
    });
  }
}
