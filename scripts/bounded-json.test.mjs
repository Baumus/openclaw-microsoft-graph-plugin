import { describe, expect, it, vi } from "vitest";
import { boundedJson } from "./bounded-json.mjs";

describe("boundedJson", () => {
  it("parses a bounded streaming JSON response", async () => {
    const response = new Response(JSON.stringify({ ok: true }), {
      headers: { "content-type": "application/json" },
    });
    await expect(boundedJson(response)).resolves.toEqual({ ok: true });
  });

  it("cancels a headerless response as soon as the streaming limit is exceeded", async () => {
    const cancel = vi.fn();
    let emitted = false;
    const body = new ReadableStream({
      pull(controller) {
        if (!emitted) {
          emitted = true;
          controller.enqueue(new Uint8Array(65));
        } else {
          controller.close();
        }
      },
      cancel,
    });
    const response = new Response(body);
    await expect(boundedJson(response, { maximum: 64, tooLarge: "too_large" })).rejects.toThrow("too_large");
    expect(cancel).toHaveBeenCalledTimes(1);
  });

  it("cancels without reading when Content-Length exceeds the limit", async () => {
    const cancel = vi.fn();
    const body = new ReadableStream({ cancel });
    const response = new Response(body, { headers: { "content-length": "65" } });
    await expect(boundedJson(response, { maximum: 64, tooLarge: "too_large" })).rejects.toThrow("too_large");
    expect(cancel).toHaveBeenCalledTimes(1);
  });

  it("cancels when Content-Length understates the streamed response size", async () => {
    const cancel = vi.fn();
    let emitted = false;
    const body = new ReadableStream({
      pull(controller) {
        if (!emitted) {
          emitted = true;
          controller.enqueue(new Uint8Array(65));
        } else {
          controller.close();
        }
      },
      cancel,
    });
    const response = new Response(body, { headers: { "content-length": "1" } });
    await expect(boundedJson(response, { maximum: 64, tooLarge: "too_large" })).rejects.toThrow("too_large");
    expect(cancel).toHaveBeenCalledTimes(1);
  });
});
