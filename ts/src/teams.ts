// Optional Teams tools: the tools of a TeamsRelay /mcp endpoint, discovered once over Streamable HTTP and served on this server with
// the teams_ prefix, next to mail and calendar. Off unless OUTLOOK_MCP_TEAMS_URL is set. The bearer token goes to that one address
// only: every request is checked against it, redirects are refused, and no tool argument can change where a call goes.

import { Client, ProtocolError, SdkError, SdkErrorCode, SdkHttpError, StreamableHTTPClientTransport, UnauthorizedError, type FetchLike, type Tool } from "@modelcontextprotocol/client";
import { fromJsonSchema, type CallToolResult, type JsonSchemaType, type McpServer } from "@modelcontextprotocol/server";

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
  let u: URL;
  try {
    u = new URL(url.trim());
  } catch {
    throw new TeamsConfigError("OUTLOOK_MCP_TEAMS_URL is not a valid URL");
  }
  const local = ["127.0.0.1", "::1", "[::1]", "localhost"].includes(u.hostname);
  if (u.username || u.password || u.search || u.hash || !(u.protocol === "https:" || (u.protocol === "http:" && local))) {
    throw new TeamsConfigError("OUTLOOK_MCP_TEAMS_URL must be https, or http on localhost, without credentials, query or fragment");
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

/**
 * Connects to the TeamsRelay MCP endpoint, lists its tools once and registers each on `server` as teams_<name>, with its title,
 * description, input and output schema, annotations, icons and _meta. A call goes upstream unchanged and its result (content
 * blocks, structuredContent, isError) comes back unchanged; a transport failure becomes an isError result, a JSON-RPC error from
 * the relay an isError result with its message, token redacted.
 * Never throws for Teams: no URL configured gives no network and no tool; an invalid configuration, an unreachable relay or a
 * refused token give no tool and `error`. Mail and calendar tools stay as registered. A name already registered on `server` is
 * skipped, never replaced. stderr gets only messages from describe(), never relay text or the token.
 */
export async function registerTeamsTools(server: McpServer, options: TeamsOptions = {}): Promise<TeamsRelay> {
  const off = { tools: [], close: async () => {} };
  const env = options.env ?? process.env;
  let config: TeamsConfig | null;
  try {
    config = teamsConfig(options.url ?? env.OUTLOOK_MCP_TEAMS_URL, options.token ?? env.OUTLOOK_MCP_TEAMS_TOKEN);
  } catch (e) {
    const error = `Teams tools off, invalid configuration: ${describe(e)}`;
    console.error(`outlook-mcp: ${error}`);
    return { ...off, error };
  }
  if (!config) return off;
  const token = config.token;
  const redact = (s: string) => s.split(token).join("[redacted]").slice(0, 500);
  const connectMs = options.connectTimeoutMs ?? CONNECT_TIMEOUT_MS;
  const callMs = options.callTimeoutMs ?? CALL_TIMEOUT_MS;
  let headerMs = connectMs;

  let closing: Promise<void> | undefined;
  const client = new Client({ name: "outlook-mcp-teams", version: "3.0.0" });
  client.onerror = (e) => {
    if (!closing) console.error(`outlook-mcp: Teams relay: ${describe(e)}`);
  };
  const transport = new StreamableHTTPClientTransport(config.url, {
    authProvider: { token: async () => token },
    fetch: guardedFetch(config.url, () => headerMs, options.fetch ?? fetch),
  });
  const close = () => (closing ??= client.close().catch(() => {}));

  let tools: Tool[];
  try {
    await client.connect(transport, { timeout: connectMs });
    tools = [];
    let cursor: string | undefined;
    const seenCursors = new Set<string>();
    for (let page = 0; page < 20; page++) {
      const result = await client.listTools(cursor ? { cursor } : undefined, { timeout: connectMs });
      tools.push(...result.tools);
      if (!result.nextCursor) break;
      if (seenCursors.has(result.nextCursor) || page === 19) throw new Error("Teams tool pagination did not finish");
      seenCursors.add(result.nextCursor);
      cursor = result.nextCursor;
    }
  } catch (e) {
    await close();
    const error = `Teams relay unavailable: ${describe(e)}`;
    console.error(`outlook-mcp: ${error}`);
    return { tools: [], error, close };
  }
  headerMs = callMs;

  const names: string[] = [];
  let badNames = 0;
  for (const t of tools) {
    const name = TEAMS_PREFIX + t.name;
    if (!TOOL_NAME.test(name)) {
      badNames++;
      continue;
    }
    try {
      server.registerTool(
        name,
        {
          title: t.title,
          description: t.description,
          inputSchema: fromJsonSchema<Record<string, unknown>>(t.inputSchema as JsonSchemaType),
          outputSchema: t.outputSchema ? fromJsonSchema(t.outputSchema as JsonSchemaType) : undefined,
          annotations: t.annotations,
          icons: t.icons,
          _meta: t._meta,
        },
        async (args, ctx) => {
          try {
            return await client.callTool({ name: t.name, arguments: args }, { timeout: callMs, maxTotalTimeout: callMs, signal: ctx.mcpReq.signal });
          } catch (e) {
            return failure(e instanceof ProtocolError ? `Teams relay error ${e.code}: ${redact(e.message)}` : `Teams relay error: ${describe(e)}`);
          }
        },
      );
      names.push(name);
    } catch (e) {
      const taken = e instanceof Error && e.message === `Tool ${name} is already registered`;
      console.error(`outlook-mcp: Teams tool ${redact(name)} skipped: ${taken ? "name already registered" : "schema not usable"}`);
    }
  }
  if (badNames) console.error(`outlook-mcp: ${badNames} Teams tool(s) skipped: name not valid for MCP`);
  if (!names.length) {
    await close();
    return { tools: [], close };
  }
  const previous = server.server.onclose;
  server.server.onclose = () => {
    previous?.();
    void close();
  };
  return { tools: names, close };
}
