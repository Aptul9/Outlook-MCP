// HTML mail bodies to Markdown and back (port of the html helpers in outlook_api.py).

const ENTITIES: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: "\u00a0" };

export function htmlUnescape(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (m, e: string) => {
    if (e[0] === "#") {
      const n = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(n) && n >= 0 && n <= 0x10ffff ? String.fromCodePoint(n) : m;
    }
    return ENTITIES[e.toLowerCase()] ?? m;
  });
}

export function htmlEscape(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#x27;");
}

const stripTags = (s: string) => s.replace(/<[^>]+>/g, "");

/** Defender SafeLinks wrap every URL: the target is in the url parameter. */
export function unwrapSafelink(url: string): string {
  if (!url) return url;
  try {
    const u = new URL(url);
    if (u.hostname.toLowerCase().includes("safelinks.protection.outlook.com")) return u.searchParams.get("url") || url;
  } catch {
    // not a URL
  }
  return url;
}

function table(inner: string): string {
  const rows: string[] = [];
  for (const [, r = ""] of inner.matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const cells = [...r.matchAll(/<(?:td|th)[^>]*>([\s\S]*?)<\/(?:td|th)>/gi)].map(([, c = ""]) =>
      htmlUnescape(stripTags(c.replace(/<br\s*\/?>/gi, " ")).replace(/<[^>]+>/g, " "))
        .replace(/\u00a0/g, " ")
        .replace(/\s+/g, " ")
        .trim()
        .replace(/\|/g, "\\|"),
    );
    if (!cells.some(Boolean)) continue;
    rows.push(`| ${cells.join(" | ")} |`);
    if (rows.length === 1) rows.push(`| ${cells.map(() => "---").join(" | ")} |`);
  }
  return rows.length ? `\n\n${rows.join("\n")}\n\n` : "";
}

export function htmlToMarkdown(raw: string): string {
  if (!raw) return "";
  let s = raw
    .replace(/<(style|script|head|title|xml)[^>]*>[\s\S]*?<\/\1>/gi, "")
    .replace(/<!\[if[^\]]*\]>[\s\S]*?<!\[endif\]>/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "");
  for (let level = 1; level <= 6; level++) {
    s = s.replace(new RegExp(`<h${level}[^>]*>([\\s\\S]*?)</h${level}>`, "gi"), (_, t: string) => `\n\n${"#".repeat(level)} ${stripTags(t).trim()}\n\n`);
  }
  s = s
    .replace(/<pre[^>]*>([\s\S]*?)<\/pre>/gi, (_, t: string) => `\n\`\`\`\n${htmlUnescape(stripTags(t).trim())}\n\`\`\`\n`)
    .replace(/<code[^>]*>([\s\S]*?)<\/code>/gi, (_, t: string) => `\`${htmlUnescape(stripTags(t).trim())}\``)
    .replace(/<(b|strong)(?:\s[^>]*)?>([\s\S]*?)<\/\1>/gi, (_, __, t: string) => (t.trim() ? `**${t.trim()}**` : ""))
    .replace(/<(i|em)(?:\s[^>]*)?>([\s\S]*?)<\/\1>/gi, (_, __, t: string) => (t.trim() ? `*${t.trim()}*` : ""))
    .replace(/<a\s+[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi, (_, href: string, t: string) => {
      const url = unwrapSafelink(htmlUnescape(href.trim()));
      const text = unwrapSafelink(htmlUnescape(stripTags(t).trim()));
      return !text || text === url ? url : `[${text}](${url})`;
    })
    .replace(/<table[^>]*>([\s\S]*?)<\/table>/gi, (_, t: string) => table(t))
    .replace(/<blockquote[^>]*>([\s\S]*?)<\/blockquote>/gi, (_, t: string) => {
      const inner = htmlToMarkdown(t);
      return `\n${inner.split("\n").filter((l) => l.trim()).map((l) => `> ${l}`).join("\n")}\n`;
    })
    .replace(/<li[^>]*>/gi, "\n* ")
    .replace(/<\/li>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|ul|ol|table)>/gi, "\n\n")
    .replace(/<hr\s*[^>]*\/?>/gi, "\n---\n");
  s = htmlUnescape(stripTags(s))
    .replace(/\u00a0/g, " ")
    .replace(/\u200b/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n");
  return s.trim();
}

const HTML_TAG = /<(?:p|div|span|b|strong|i|em|u|mark|br|hr|ul|ol|li|table|thead|tbody|tr|th|td|h[1-6]|a|blockquote|pre|code|font|img)\b/i;

export const looksLikeHtml = (text: string | undefined) => !!text && HTML_TAG.test(text);

/** Body for an HTML reply or forward comment: HTML kept, plain text escaped with line breaks kept. */
export function formatBodyForHtml(body: string, isHtml = false): string {
  if (!body) return "";
  if (isHtml || looksLikeHtml(body)) return body;
  const escaped = htmlEscape(body).replace(/\r\n/g, "\n").replace(/\n/g, "<br>\n");
  return `<div style="font-family:Aptos,Calibri,sans-serif;font-size:11pt;">${escaped}</div>`;
}
