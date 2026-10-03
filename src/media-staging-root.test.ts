import { mkdtemp, mkdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import entry, { readProtectedMediaSource } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

/**
 * Regression: OpenClaw stores inbound media under the gateway STATE directory
 * (`<stateDir>/media/inbound`), not under the agent workspace. The staging root used to be
 * resolved from `workspaceDir` alone, so the plugin looked in a directory that does not exist and
 * every protected media read failed with `invalid_source_media_uri` — no mail/calendar/to-do
 * attachment and no OneDrive upload could ever be sourced, regardless of input.
 */
describe("protected media staging root", () => {
  const current = graphPolicyFixture();
  const policy = { version: 2 as const, rules: current.rules, services: current.services };

  function register(stateDir: string): void {
    entry.register({
      pluginConfig: { enabled: true, policy },
      registerTool: vi.fn(),
      registerCli: vi.fn(),
      registerGatewayMethod: vi.fn(),
      runtime: { state: { resolveStateDir: vi.fn(() => stateDir) } },
      on: vi.fn(),
      logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
    } as never);
  }

  it("reads media the host saved under <stateDir>/media/inbound", async () => {
    const base = await mkdtemp(join(tmpdir(), "mg-media-"));
    const stateDir = join(base, "state");
    // The workspace deliberately has no media directory at all, which is what the gateway does.
    const workspaceDir = join(base, "state", "workspace");
    await mkdir(join(stateDir, "media", "inbound"), { recursive: true });
    await mkdir(workspaceDir, { recursive: true });

    const bytes = Buffer.from("# design brief\n");
    await writeFile(join(stateDir, "media", "inbound", "brief---abc123.md"), bytes);

    register(stateDir);

    const read = await readProtectedMediaSource("media://inbound/brief---abc123.md", workspaceDir);
    expect(read.equals(bytes)).toBe(true);
  });

  it("still reads media a host places under <workspaceDir>/media/inbound", async () => {
    const base = await mkdtemp(join(tmpdir(), "mg-media-ws-"));
    const stateDir = join(base, "state");
    const workspaceDir = join(base, "workspace");
    // State dir exists but holds no media; the file lives under the workspace instead.
    await mkdir(join(stateDir, "media", "inbound"), { recursive: true });
    await mkdir(join(workspaceDir, "media", "inbound"), { recursive: true });

    const bytes = Buffer.from("workspace-sourced\n");
    await writeFile(join(workspaceDir, "media", "inbound", "ws---def456.txt"), bytes);

    register(stateDir);

    const read = await readProtectedMediaSource("media://inbound/ws---def456.txt", workspaceDir);
    expect(read.equals(bytes)).toBe(true);
  });

  it("still rejects traversal and unknown names", async () => {
    const base = await mkdtemp(join(tmpdir(), "mg-media-neg-"));
    const stateDir = join(base, "state");
    const workspaceDir = join(base, "workspace");
    await mkdir(join(stateDir, "media", "inbound"), { recursive: true });
    await mkdir(workspaceDir, { recursive: true });
    await writeFile(join(base, "outside.txt"), Buffer.from("secret\n"));

    register(stateDir);

    await expect(readProtectedMediaSource("media://inbound/../outside.txt", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
    await expect(readProtectedMediaSource("media://inbound/not-there.md", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
    await expect(readProtectedMediaSource("file:///etc/passwd", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
  });
});
