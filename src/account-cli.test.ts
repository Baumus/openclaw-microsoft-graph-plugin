import { describe, expect, it, vi } from "vitest";
import { registerAccountCli } from "./account-cli.js";

describe("personal account CLI installation guidance", () => {
  it("separates zero-touch startup upgrades from explicit setup without doing account work", () => {
    let registration!: (ctx: { program: unknown }) => void;
    let action!: () => void;
    const api = { registerCli: (callback: typeof registration) => { registration = callback; } };
    registerAccountCli(api as never);
    const command = { command: () => command, description: () => command, action: (callback: () => void) => { action = callback; return command; } };
    registration({ program: command });
    const write = vi.spyOn(process.stdout, "write").mockReturnValue(true);
    try {
      action(); const output = String(write.mock.calls[0]?.[0]);
      expect(output).toContain("automatically migrates an existing valid connection at Gateway startup");
      expect(output).toContain("No Control UI, personal login, administrator role, confirmation button");
      expect(output).toContain("Settings → Labs → Custom plugin UI");
      expect(output).toContain("Settings → Profile → Identity");
      expect(output).toContain("gateway-owner/dashboard token is not your personal identity");
      expect(output).toContain("Microsoft-Einrichtung");
      expect(output).toContain("Limited access → Request admin");
      expect(output).toContain("guidance only");
      expect(output).toContain("permitted read-only request");
    } finally { write.mockRestore(); }
  });
});
