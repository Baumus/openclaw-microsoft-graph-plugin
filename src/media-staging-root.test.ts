import { mkdtemp, mkdir, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import entry, { openProtectedMediaUploadSource, readProtectedMediaSource } from "./index.js";
import { graphPolicyFixture } from "./fixtures/graph-access-policy.js";

const resolverKey = Symbol.for("@baumus/openclaw-microsoft-graph/plugin-state-dir-resolver");
const previousResolver = (globalThis as Record<symbol, unknown>)[resolverKey];
const temporaryDirectories: string[] = [];

afterEach(async () => {
  if (previousResolver === undefined) delete (globalThis as Record<symbol, unknown>)[resolverKey];
  else (globalThis as Record<symbol, unknown>)[resolverKey] = previousResolver;
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

async function roots() {
  const base = await mkdtemp(join(tmpdir(), "mg-media-root-"));
  temporaryDirectories.push(base);
  const stateDir = join(base, "state");
  const workspaceDir = join(base, "workspace");
  await mkdir(join(stateDir, "media", "inbound"), { recursive: true });
  await mkdir(join(workspaceDir, "media", "inbound"), { recursive: true });
  return { base, stateDir, workspaceDir };
}

function register(stateDir?: string): void {
  const current = graphPolicyFixture();
  const policy = { version: 2 as const, rules: current.rules, services: current.services };
  entry.register({
    pluginConfig: { enabled: true, policy },
    registerTool: vi.fn(),
    registerCli: vi.fn(),
    registerGatewayMethod: vi.fn(),
    ...(stateDir ? { runtime: { state: { resolveStateDir: vi.fn(() => stateDir) } } } : {}),
    on: vi.fn(),
    logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
  } as never);
}

describe("protected media staging root", () => {
  it("reads and streams media from the host state directory", async () => {
    const { stateDir, workspaceDir } = await roots();
    const bytes = Buffer.from("# design brief\n");
    await writeFile(join(stateDir, "media", "inbound", "brief---abc123.md"), bytes);
    register(stateDir);

    expect(await readProtectedMediaSource("media://inbound/brief---abc123.md", workspaceDir)).toEqual(bytes);
    const stream = await openProtectedMediaUploadSource("media://inbound/brief---abc123.md", workspaceDir);
    try {
      expect(stream.size).toBe(bytes.length);
      expect(await stream.readChunk(0, bytes.length)).toEqual(bytes);
      await stream.assertUnchanged();
    } finally {
      await stream.close();
    }
  });

  it("uses the workspace only on hosts without a state resolver", async () => {
    const { workspaceDir } = await roots();
    const bytes = Buffer.from("workspace-sourced\n");
    await writeFile(join(workspaceDir, "media", "inbound", "ws---def456.txt"), bytes);
    delete (globalThis as Record<symbol, unknown>)[resolverKey];
    register();

    expect(await readProtectedMediaSource("media://inbound/ws---def456.txt", workspaceDir)).toEqual(bytes);
    const stream = await openProtectedMediaUploadSource("media://inbound/ws---def456.txt", workspaceDir);
    try {
      expect(await stream.readChunk(0, bytes.length)).toEqual(bytes);
    } finally {
      await stream.close();
    }
  });

  it("never substitutes a same-named workspace file when state media is absent", async () => {
    const { stateDir, workspaceDir } = await roots();
    await writeFile(join(workspaceDir, "media", "inbound", "shared.txt"), "wrong file");
    register(stateDir);

    await expect(readProtectedMediaSource("media://inbound/shared.txt", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
    await expect(openProtectedMediaUploadSource("media://inbound/shared.txt", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
  });

  it("never substitutes a workspace file when the state path is rejected", async () => {
    const { base, stateDir, workspaceDir } = await roots();
    const outside = join(base, "outside.txt");
    await writeFile(outside, "protected bytes");
    await symlink(outside, join(stateDir, "media", "inbound", "shared.txt"));
    await writeFile(join(workspaceDir, "media", "inbound", "shared.txt"), "wrong file");
    register(stateDir);

    await expect(readProtectedMediaSource("media://inbound/shared.txt", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
    await expect(openProtectedMediaUploadSource("media://inbound/shared.txt", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
  });

  it("rejects invalid URIs", async () => {
    const { stateDir, workspaceDir } = await roots();
    register(stateDir);
    await expect(readProtectedMediaSource("media://inbound/../outside.txt", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
    await expect(readProtectedMediaSource("media://inbound/not-there.md", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
    await expect(readProtectedMediaSource("file:///etc/passwd", workspaceDir)).rejects.toThrow("invalid_source_media_uri");
  });
});
