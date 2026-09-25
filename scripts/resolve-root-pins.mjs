import { execFile } from "node:child_process";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { promisify } from "node:util";
import { exchangeRefreshToken, readCredential } from "../src/credential.ts";
import { boundedJson } from "./bounded-json.mjs";

const execFileAsync = promisify(execFile);

export async function resolveAllowedRoots({
  policy,
  credentialRef,
  readCredentialFn = readCredential,
  exchangeRefreshTokenFn = exchangeRefreshToken,
  fetchFn = fetch,
  write = (line) => process.stdout.write(line),
}) {
  if (typeof credentialRef !== "string" || !/^[A-Za-z0-9._/@+-]+$/.test(credentialRef)) throw new Error("invalid_secret_reference");
  const credential = await readCredentialFn(credentialRef);
  const token = await exchangeRefreshTokenFn(credential, ["Files.Read"]);

  for (const root of policy.services.onedrive.allowed_roots.filter((candidate) => !candidate.item_id)) {
    const encoded = root.path.split("/").map((segment) => encodeURIComponent(segment)).join("/");
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(root.drive_id)}/root:${encoded}?$select=id,name`;
    const response = await fetchFn(url, { headers: { accept: "application/json", authorization: `Bearer ${token}` } });
    if (!response.ok) throw new Error(`root_resolution_failed:${root.label}:${response.status}`);
    const item = await boundedJson(response, {
      tooLarge: `root_resolution_too_large:${root.label}`,
      invalid: `root_resolution_invalid:${root.label}`,
    });
    if (typeof item.id !== "string" || !item.id || typeof item.name !== "string") throw new Error(`root_resolution_invalid:${root.label}`);
    write(`${JSON.stringify({ label: root.label, item_id: item.id, resolved_name: item.name })}\n`);
  }
}

export async function main(argv = process.argv.slice(2)) {
  if (argv.length !== 2 || argv[0] !== "--credential-ref") throw new Error("usage: resolve-root-pins --credential-ref <pass-ref>");
  const { stdout } = await execFileAsync("openclaw", ["config", "get", "plugins.entries", "--json"], { encoding: "utf8", timeout: 10_000, maxBuffer: 2 * 1024 * 1024 });
  const policy = JSON.parse(stdout)?.["microsoft-graph"]?.config?.policy;
  if (!policy) throw new Error("policy_unavailable");
  await resolveAllowedRoots({ policy, credentialRef: argv[1] });
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) await main();
