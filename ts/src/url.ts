// The one rule for every backend address that receives a credential (OE Copilot broker or gateway, TeamsRelay /mcp).
// https anywhere. Plain http only to this machine, or to a single-label host name: a service on the same container
// network (agent, webapp), which public DNS does not resolve. Never credentials, query or fragment in the address.

const LOOPBACK = new Set(["localhost", "127.0.0.1", "::1", "[::1]"]);
const SERVICE_NAME = /^[a-z0-9](?:[a-z0-9_-]{0,61}[a-z0-9])?$/i;

/** True when plain http to this host keeps the credential on the machine or its container network. */
export function plainHttpAllowed(hostname: string): boolean {
  return LOOPBACK.has(hostname) || SERVICE_NAME.test(hostname);
}

/** The parsed address, or null when it breaks the rule above (or is not a URL). */
export function backendUrl(raw: string): URL | null {
  let u: URL;
  try {
    u = new URL(raw.trim());
  } catch {
    return null;
  }
  if (u.username || u.password || u.search || u.hash) return null;
  if (u.protocol === "https:") return u;
  return u.protocol === "http:" && plainHttpAllowed(u.hostname) ? u : null;
}
