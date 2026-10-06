/** Read-only, best-effort ClawHub version check for the administrator UI. */
const PACKAGE_NAME = "@baumus/openclaw-microsoft-graph";
const PACKAGE_URL = "https://clawhub.ai/api/v1/packages/%40baumus%2Fopenclaw-microsoft-graph";
const CACHE_MS = 15 * 60_000;
const MAX_RESPONSE_BYTES = 16_384;

type VersionStatus = { currentVersion: string; latestVersion: string; updateAvailable: boolean };
type HandlerContext = { params: unknown; respond(ok: boolean, payload?: unknown, error?: unknown): void };
type GatewayApi = { version?: string; registerGatewayMethod(method: string, handler: (context: HandlerContext) => void | Promise<void>, options: { scope: "operator.admin" }): void };

function parts(value: unknown): number[] | undefined {
  if (typeof value !== "string" || !/^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$/.test(value)) return undefined;
  const result = value.split(".").map(Number);
  return result.every(Number.isSafeInteger) ? result : undefined;
}

export function newerRelease(current: string, payload: unknown): VersionStatus | undefined {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) return undefined;
  const pkg = (payload as { package?: unknown }).package;
  if (!pkg || typeof pkg !== "object" || Array.isArray(pkg)) return undefined;
  const { name, latestVersion, scanStatus } = pkg as Record<string, unknown>;
  const installed = parts(current);
  const latest = parts(latestVersion);
  if (name !== PACKAGE_NAME || scanStatus !== "clean" || !installed || !latest) return undefined;
  const updateAvailable = latest.some((n, index) => n > installed[index] && latest.slice(0, index).every((prior, i) => prior === installed[i]));
  return { currentVersion: current, latestVersion: latestVersion as string, updateAvailable };
}

async function readJsonBounded(response: Response): Promise<unknown> {
  if (!response.ok || !response.body) throw new Error("update_check_unavailable");
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let bytes = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      bytes += value.byteLength;
      if (bytes > MAX_RESPONSE_BYTES) throw new Error("update_check_unavailable");
      chunks.push(value);
    }
  } finally { reader.releaseLock(); }
  const merged = new Uint8Array(bytes);
  let offset = 0;
  for (const chunk of chunks) { merged.set(chunk, offset); offset += chunk.byteLength; }
  return JSON.parse(new TextDecoder().decode(merged));
}

export function registerUpdateStatusMethod(api: GatewayApi, fetcher: typeof fetch = fetch): void {
  let cache: { value: VersionStatus; until: number } | undefined;
  let pending: Promise<VersionStatus | undefined> | undefined;
  async function check(): Promise<VersionStatus | undefined> {
    if (cache && Date.now() < cache.until) return cache.value;
    if (!pending) {
      pending = (async () => {
        if (!api.version) return undefined;
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), 4000);
        try {
          const response = await fetcher(PACKAGE_URL, { method: "GET", signal: controller.signal, redirect: "error", headers: { accept: "application/json" } });
          const result = newerRelease(api.version, await readJsonBounded(response));
          if (result) cache = { value: result, until: Date.now() + CACHE_MS };
          return result;
        } catch { return undefined; }
        finally { clearTimeout(timer); }
      })().finally(() => { pending = undefined; });
    }
    return pending;
  }
  api.registerGatewayMethod("microsoft-graph.updateStatus", async ({ params, respond }) => {
    if (!params || typeof params !== "object" || Array.isArray(params) || Object.keys(params).length !== 0) {
      respond(false, undefined, { code: "INVALID_REQUEST", message: "Invalid update status request" });
      return;
    }
    const status = await check();
    respond(true, status ?? { updateAvailable: false });
  }, { scope: "operator.admin" });
}
