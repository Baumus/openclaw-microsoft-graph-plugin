import { describe, expect, it } from "vitest";
import { MAX_REPLY_HISTORY_BYTES, prependReplyHtml, replyFragment } from "./mail-reply.js";
describe("opaque native reply history", () => {
  it("keeps Outlook styles, conditional comments, images, attributes and old markup byte-for-byte", () => {
    const history = '<html><head><!--[if mso]><style>p{margin:0}</style><![endif]--></head><!-- <body> --><BODY title="a>b"><table><tr><td>Old</td></tr></table><img src=https://images.example.invalid/logo width=200 height=97><img src="cid:logo@original"></BODY></html>';
    expect(prependReplyHtml(history, '<p>Reply</p>')).toBe(history.replace('<BODY title="a>b">', '<BODY title="a>b"><p>Reply</p>'));
  });
  it("preserves long threads beyond the authored-content limit without truncation", () => {
    const history = '<body>' + 'x'.repeat(100_000) + '</body>';
    expect(prependReplyHtml(history, '<p>Reply</p>')).toBe('<body><p>Reply</p>' + 'x'.repeat(100_000) + '</body>');
    expect(() => prependReplyHtml('x'.repeat(MAX_REPLY_HISTORY_BYTES), 'x')).toThrow('reply_history_too_large');
  });
  it("fails on unavailable history instead of composing an incomplete replacement", () => {
    expect(() => prependReplyHtml(undefined, 'reply')).toThrow('reply_history_unavailable');
    expect(() => prependReplyHtml('<html><head></head></html>', 'reply')).toThrow('reply_history_body_missing');
  });
  it("escapes newly authored text and refuses document/style/event injection", () => {
    expect(replyFragment({ contentType: 'Text', content: '<x>&\nNext' })).toBe('<div>&lt;x&gt;&amp;<br>Next</div>');
    for (const content of ['<html>new</html>', '<style>p{display:none}</style>', '<p onclick="bad()">Reply</p>']) expect(() => replyFragment({ contentType: 'HTML', content })).toThrow('invalid_reply_fragment');
  });
});
