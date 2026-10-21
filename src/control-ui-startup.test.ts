// @vitest-environment happy-dom
import { readFileSync } from "node:fs";
import { describe, expect, it, vi } from "vitest";
import plugin from "./control-ui.js";

// Vitest stubs CSS even with esbuild text attributes; the emitted entry is
// independently checked in a real browser against the unchanged source bytes.
vi.mock("./control-ui.css", () => ({ default: readFileSync("src/control-ui.css", "utf8") }));

const selector = 'style[data-microsoft-graph="styles"]';
function activate(fail?: "page" | "navigation", throwingDisposer = false) {
  const pageStop = vi.fn(() => { if (throwingDisposer) throw new Error("dispose"); });
  const navigationStop = vi.fn();
  const host = { locale: "en", ui: {
    registerPage: vi.fn(() => { if (fail === "page") throw new Error("page"); return pageStop; }),
    registerNavigation: vi.fn(() => { if (fail === "navigation") throw new Error("navigation"); return navigationStop; }),
  } };
  return { host, pageStop, navigationStop, start: () => plugin.activate(host as never) as () => void };
}

describe("self-contained native startup", () => {
  it("registers page/navigation synchronously with the exact original CSS already present", () => {
    const test = activate();
    test.host.ui.registerPage.mockImplementation(() => {
      expect(document.querySelector(selector)?.textContent).toBe(readFileSync("src/control-ui.css", "utf8"));
      return test.pageStop;
    });
    const stop = test.start();
    try {
      expect(test.host.ui.registerPage).toHaveBeenCalledOnce();
      expect(test.host.ui.registerNavigation).toHaveBeenCalledOnce();
      expect(document.querySelector(selector)?.textContent).toBe(readFileSync("src/control-ui.css", "utf8"));
      expect(document.querySelectorAll(selector)).toHaveLength(1);
    } finally { stop(); }
    expect(document.querySelectorAll(selector)).toHaveLength(0);
    stop();
    expect(test.pageStop).toHaveBeenCalledOnce();
    expect(test.navigationStop).toHaveBeenCalledOnce();
  });
  it("shares styles across overlapping activations and survives disposal in either order", () => {
    for (const reverse of [false, true]) {
      const first = activate().start(), second = activate().start();
      expect(document.querySelectorAll(selector)).toHaveLength(1);
      (reverse ? second : first)();
      expect(document.querySelectorAll(selector)).toHaveLength(1);
      (reverse ? first : second)();
      expect(document.querySelectorAll(selector)).toHaveLength(0);
    }
  });
  it.each(["page", "navigation"] as const)("rolls back failed %s registration without removing another activation's styles", fail => {
    for (const overlapping of [false, true]) {
      const stop = overlapping ? activate().start() : undefined;
      const test = activate(fail);
      expect(test.start).toThrow(fail);
      expect(document.querySelectorAll(selector)).toHaveLength(overlapping ? 1 : 0);
      expect(test.pageStop).toHaveBeenCalledTimes(fail === "navigation" ? 1 : 0);
      stop?.();
      expect(document.querySelectorAll(selector)).toHaveLength(0);
    }
  });
  it("releases styles and remaining registrations even when a disposer throws", () => {
    const test = activate(undefined, true), stop = test.start();
    expect(stop).toThrow("dispose");
    expect(test.navigationStop).toHaveBeenCalledOnce();
    expect(document.querySelectorAll(selector)).toHaveLength(0);
    expect(stop).not.toThrow();
  });
});
