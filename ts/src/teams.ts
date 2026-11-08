// Optional Teams tools: the tools of a TeamsRelay /mcp endpoint, discovered once over Streamable HTTP and served on this server with
// the teams_ prefix, next to mail and calendar. Off unless OUTLOOK_MCP_TEAMS_URL is set. The bearer token goes to that one address
// only: every request is checked against it, redirects are refused, and no tool argument can change where a call goes.

import { Client, ProtocolError, SdkError, SdkErrorCode, SdkHttpError, StreamableHTTPClientTransport, UnauthorizedError, type FetchLike, type Tool } from "@modelcontextprotocol/client";
import { fromJsonSchema, type CallToolResult, type JsonSchemaType, type McpServer } from "@modelcontextprotocol/server";
import { backendUrl } from "./url.ts";

export const TEAMS_PREFIX = "teams_";
export const CONNECT_TIMEOUT_MS = 15_000;
export const CALL_TIMEOUT_MS = 120_000;
const TOOL_NAME = /^[A-Za-z0-9_.-]{1,128}$/;

export class TeamsConfigError extends Error {}
/** Failures whose text is written here, never taken from the relay: safe for logs */
class RelayError extends Error {}

export interface TeamsConfig {
  url: URL;
  token: string;
}

export interface TeamsOptions {
  /** Read when url or token is not given: OUTLOOK_MCP_TEAMS_URL, OUTLOOK_MCP_TEAMS_TOKEN. Default process.env */
  env?: NodeJS.ProcessEnv;
  url?: string;
  token?: string;
  /** Bound on each step of discovery (HTTP answer, initialize, tools/list). Default 15 s */
  connectTimeoutMs?: number;
  /** Bound on each forwarded tool call. Default 120 s */
  callTimeoutMs?: number;
  /** Underlying fetch; default the global one. The address check, no-redirect rule and timeouts wrap it */
  fetch?: FetchLike;
}

export interface TeamsRelay {
  /** Tool names registered on the server, prefix included */
  tools: string[];
  /** Set when the Teams configuration is invalid or discovery failed; mail and calendar tools are unaffected. Never carries relay text */
  error?: string;
  /** Closes the relay connection; idempotent. Also runs when the server's transport closes */
  close(): Promise<void>;
}

/** null when the URL is unset (Teams off). https, or http on localhost, without credentials, query or fragment; token required. */
export function teamsConfig(url: string | undefined, token: string | undefined): TeamsConfig | null {
  if (!url?.trim()) return null;
  const u = backendUrl(url);
  if (!u) {
    throw new TeamsConfigError(
      "OUTLOOK_MCP_TEAMS_URL must be https, or http on localhost or a container service name, without credentials, query or fragment",
    );
  }
  const t = token?.trim() ?? "";
  if (!t) throw new TeamsConfigError("OUTLOOK_MCP_TEAMS_URL needs OUTLOOK_MCP_TEAMS_TOKEN");
  if (!/^[\x21-\x7e]+$/.test(t)) throw new TeamsConfigError("OUTLOOK_MCP_TEAMS_TOKEN holds characters an HTTP header cannot carry");
  return { url: u, token: t };
}

export function teamsConfigFromEnv(env: NodeJS.ProcessEnv = process.env): TeamsConfig | null {
  return teamsConfig(env.OUTLOOK_MCP_TEAMS_URL, env.OUTLOOK_MCP_TEAMS_TOKEN);
}

const failure = (text: string): CallToolResult => ({ content: [{ type: "text", text }], isError: true });

/**
 * A failure as kind, code or HTTP status only: never an error message or body from the relay, which can echo request data.
 * Used for stderr and for the `error` of the result.
 */
export function describe(e: unknown): string {
  if (e instanceof RelayError || e instanceof TeamsConfigError) return e.message;
  if (e instanceof UnauthorizedError) return "the relay refused the token";
  if (e instanceof SdkHttpError) return `HTTP ${e.status}`;
  if (e instanceof SdkError) {
    if (e.code === SdkErrorCode.RequestTimeout) return "timed out";
    if (e.code === SdkErrorCode.ConnectionClosed || e.code === SdkErrorCode.NotConnected) return "connection closed";
    return `SDK error ${e.code}`;
  }
  if (e instanceof ProtocolError) return `JSON-RPC error ${e.code}`;
  if (e instanceof Error && e.name === "AbortError") return "aborted";
  if (e instanceof TypeError) {
    const code = (e.cause as { code?: unknown } | undefined)?.code;
    return typeof code === "string" && /^[A-Z0-9_]{1,40}$/.test(code) ? `network error (${code})` : "network error";
  }
  return "unexpected error";
}

/** The configured address only, never a redirect (the token would follow it), response headers within timeoutMs(). */
function guardedFetch(url: URL, timeoutMs: () => number, base: FetchLike): FetchLike {
  return async (input, init) => {
    if (String(input) !== url.href) throw new RelayError("request to another address refused");
    const ms = timeoutMs();
    const timer = new AbortController();
    const t = setTimeout(() => timer.abort(new RelayError(`no answer within ${ms} ms`)), ms);
    try {
      return await base(input, { ...init, redirect: "error", signal: init?.signal ? AbortSignal.any([init.signal, timer.signal]) : timer.signal });
    } finally {
      clearTimeout(t);
    }
  };
}

/** A connected relay client: the SDK client plus its own header timeout (discovery bound first, call bound after). */
interface Link {
  client: Client;
  setHeaderMs(ms: number): void;
  close(): Promise<void>;
}

/** One discovered tool, converted once and registered on every server instance. */
interface Discovered {
  upstream: string;
  name: string;
  config: {
    title?: string;
    description?: string;
    inputSchema: ReturnType<typeof fromJsonSchema<Record<string, unknown>>>;
    outputSchema?: ReturnType<typeof fromJsonSchema>;
    annotations?: Tool["annotations"];
    icons?: Tool["icons"];
    _meta?: Tool["_meta"];
  };
}

export interface TeamsBridge {
  /** Tool names currently offered, prefix included */
  readonly tools: string[];
  /** Configuration error or last discovery failure; never carries relay text */
  readonly error?: string;
  /**
   * Registers the discovered tools on `server`. Calls go through the bridge's one relay connection, so many short-lived
   * server instances (one per HTTP request) share a single discovery. When the last discovery failed, is older than
   * maxAgeMs, or a call lost the connection, a rediscovery starts in the background; this server keeps the current list.
   */
  register(server: McpServer): string[];
  /** Discovers now, sharing a discovery already running. Never throws */
  refresh(): Promise<void>;
  /** Closes the relay connection; idempotent */
  close(): Promise<void>;
}

export interface BridgeOptions extends TeamsOptions {
  /** Least time between two rediscovery attempts after a failure. Default 60 s */
  retryMs?: number;
  /** Age after which a successful tool list is fetched again. Default 10 min */
  maxAgeMs?: number;
}

const OFF: TeamsBridge = { tools: [], register: () => [], refresh: async () => {}, close: async () => {} };

/**
 * Connects to the TeamsRelay MCP endpoint and lists its tools, for every server built afterwards (see TeamsBridge.register).
 * Never throws: no URL gives no network and no tool; an invalid configuration, an unreachable relay or a refused token give
 * no tool and `error`, and the relay is tried again later. stderr gets only messages from describe(), never relay text or the token.
 */
export async function connectTeams(options: BridgeOptions = {}): Promise<TeamsBridge> {
  const env = options.env ?? process.env;
  let config: TeamsConfig | null;
  try {
    config = teamsConfig(options.url ?? env.OUTLOOK_MCP_TEAMS_URL, options.token ?? env.OUTLOOK_MCP_TEAMS_TOKEN);
  } catch (e) {
    const error = `Teams tools off, invalid configuration: ${describe(e)}`;
    console.error(`outlook-mcp: ${error}`);
    return { ...OFF, error };
  }
  if (!config) return OFF;
  const { url, token } = config;
  const redact = (s: string) => s.split(token).join("[redacted]").slice(0, 500);
  const connectMs = options.connectTimeoutMs ?? CONNECT_TIMEOUT_MS;
  const callMs = options.callTimeoutMs ?? CALL_TIMEOUT_MS;
  const retryMs = options.retryMs ?? 60_000;
  const maxAgeMs = options.maxAgeMs ?? 600_000;

  let link: Link | null = null;
  let discovered: Discovered[] = [];
  let error: string | undefined;
  let lastAttempt = 0;
  let lastSuccess = 0;
  let stale = false;
  let running: Promise<void> | undefined;
  let closed = false;

  const open = async (): Promise<Link> => {
    let headerMs = connectMs;
    let closing: Promise<void> | undefined;
    const client = new Client({ name: "outlook-mcp-teams", version: "3.0.0" });
    client.onerror = (e) => {
      if (!closing) console.error(`outlook-mcp: Teams relay: ${describe(e)}`);
    };
    const transport = new StreamableHTTPClientTransport(url, {
      authProvider: { token: async () => token },
      fetch: guardedFetch(url, () => headerMs, options.fetch ?? fetch),
    });
    const l: Link = { client, setHeaderMs: (ms) => void (headerMs = ms), close: () => (closing ??= client.close().catch(() => {})) };
    try {
      await client.connect(transport, { timeout: connectMs });
    } catch (e) {
      await l.close();
      throw e;
    }
    return l;
  };

  const listAll = async (client: Client): Promise<Tool[]> => {
    const tools: Tool[] = [];
    let cursor: string | undefined;
    const seen = new Set<string>();
    for (let page = 0; page < 20; page++) {
      const result = await client.listTools(cursor ? { cursor } : undefined, { timeout: connectMs });
      tools.push(...result.tools);
      if (!result.nextCursor) return tools;
      if (seen.has(result.nextCursor)) break;
      seen.add(result.nextCursor);
      cursor = result.nextCursor;
    }
    throw new RelayError("tool pagination did not finish");
  };

  const convert = (tools: Tool[]): Discovered[] => {
    const out: Discovered[] = [];
    let badNames = 0;
    for (const t of tools) {
      const name = TEAMS_PREFIX + t.name;
      if (!TOOL_NAME.test(name)) {
        badNames++;
        continue;
      }
      try {
        out.push({
          upstream: t.name,
          name,
          config: {
            title: t.title,
            description: t.description,
            inputSchema: fromJsonSchema<Record<string, unknown>>(t.inputSchema as JsonSchemaType),
            outputSchema: t.outputSchema ? fromJsonSchema(t.outputSchema as JsonSchemaType) : undefined,
            annotations: t.annotations,
            icons: t.icons,
            _meta: t._meta,
          },
        });
      } catch {
        console.error(`outlook-mcp: Teams tool ${redact(name)} skipped: schema not usable`);
      }
    }
    if (badNames) console.error(`outlook-mcp: ${badNames} Teams tool(s) skipped: name not valid for MCP`);
    return out;
  };

  const discover = async () => {
    lastAttempt = Date.now();
    try {
      let tools: Tool[] | undefined;
      if (link && !stale) {
        link.setHeaderMs(connectMs);
        try {
          tools = await listAll(link.client);
        } catch {
          // the old connection may be gone (relay restart): one fresh connection below
        } finally {
          link?.setHeaderMs(callMs);
        }
      }
      if (!tools) {
        const fresh = await open();
        try {
          tools = await listAll(fresh.client);
        } catch (e) {
          await fresh.close();
          throw e;
        }
        if (closed) {
          await fresh.close();
          return;
        }
        fresh.setHeaderMs(callMs);
        const old = link;
        link = fresh;
        // calls still running on the old connection keep it until they settle; new calls use the fresh one
        if (old) setTimeout(() => void old.close(), callMs).unref();
      }
      discovered = convert(tools);
      error = undefined;
      stale = false;
      lastSuccess = Date.now();
    } catch (e) {
      error = `Teams relay unavailable: ${describe(e)}`;
      console.error(`outlook-mcp: ${error}`);
    }
  };

  const refresh = () => {
    if (closed) return Promise.resolve();
    return (running ??= discover().finally(() => (running = undefined)));
  };

  const due = () => {
    const now = Date.now();
    if (running || closed) return false;
    if (error || stale) return now - lastAttempt >= retryMs;
    return now - lastSuccess >= maxAgeMs;
  };

  const call = async (upstream: string, args: Record<string, unknown>, signal: AbortSignal): Promise<CallToolResult> => {
    const l = link;
    if (!l) return failure(`Teams relay error: ${error ?? "not connected"}`);
    try {
      return (await l.client.callTool({ name: upstream, arguments: args }, { timeout: callMs, maxTotalTimeout: callMs, signal })) as CallToolResult;
    } catch (e) {
      if (e instanceof ProtocolError) return failure(`Teams relay error ${e.code}: ${redact(e.message)}`);
      // transport failure: the next server instance triggers a reconnect (rate limited by retryMs)
      if (!(e instanceof Error && e.name === "AbortError")) stale = true;
      return failure(`Teams relay error: ${describe(e)}`);
    }
  };

  await refresh();

  return {
    get tools() {
      return discovered.map((d) => d.name);
    },
    get error() {
      return error;
    },
    register(server) {
      if (due()) void refresh();
      const names: string[] = [];
      for (const d of discovered) {
        try {
          server.registerTool(d.name, d.config, (args, ctx) => call(d.upstream, args as Record<string, unknown>, ctx.mcpReq.signal));
          names.push(d.name);
        } catch (e) {
          const taken = e instanceof Error && e.message === `Tool ${d.name} is already registered`;
          console.error(`outlook-mcp: Teams tool ${redact(d.name)} skipped: ${taken ? "name already registered" : "schema not usable"}`);
        }
      }
      return names;
    },
    refresh,
    async close() {
      closed = true;
      await running;
      await link?.close();
    },
  };
}

/**
 * One server, one relay connection (stdio and tests): connectTeams, register on `server`, close the connection when the
 * server's transport closes. Same guarantees as connectTeams; a name already registered on `server` is skipped, never replaced.
 */
export async function registerTeamsTools(server: McpServer, options: TeamsOptions = {}): Promise<TeamsRelay> {
  const bridge = await connectTeams({ ...options, retryMs: Number.POSITIVE_INFINITY, maxAgeMs: Number.POSITIVE_INFINITY });
  const tools = bridge.register(server);
  if (!tools.length) {
    await bridge.close();
    return { tools: [], error: bridge.error, close: bridge.close };
  }
  const previous = server.server.onclose;
  server.server.onclose = () => {
    previous?.();
    void bridge.close();
  };
  return { tools, error: bridge.error, close: bridge.close };
}
