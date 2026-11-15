/** Only newly authored content is validated. Provider history is opaque: never fetch,
 * sanitize, reserialize or copy its external images/CID attachments. */
export const MAX_REPLY_HISTORY_BYTES = 4 * 1024 * 1024;
export function replyFragment(body: { contentType: string; content: string }): string {
  if (body.contentType === "Text") return `<div>${body.content.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r\n|\r|\n/g, "<br>")}</div>`;
  if (/<\/?(?:html|head|body|style)\b|\bon\w+\s*=|(?:javascript|vbscript)\s*:/i.test(body.content)) throw new Error("invalid_reply_fragment");
  return body.content;
}
export function prependReplyHtml(history: unknown, fragment: string): string {
  if (typeof history !== "string") throw new Error("reply_history_unavailable");
  // Skip comments (including Outlook conditional comments); preserve every existing byte.
  const tags = /<!--[\s\S]*?-->|<body\b(?:[^>"']|"[^"]*"|'[^']*')*>/gi;
  let match: RegExpExecArray | null;
  let offset = 0;
  while ((match = tags.exec(history))) {
    if (!match[0].startsWith("<!--")) { offset = match.index + match[0].length; break; }
  }
  if (!offset && /<html\b|<head\b/i.test(history)) throw new Error("reply_history_body_missing");
  const merged = history.slice(0, offset) + fragment + history.slice(offset);
  if (Buffer.byteLength(merged, "utf8") > MAX_REPLY_HISTORY_BYTES) throw new Error("reply_history_too_large");
  return merged;
}
