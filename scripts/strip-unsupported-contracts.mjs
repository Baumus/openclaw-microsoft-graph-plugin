import { readFileSync, writeFileSync } from "node:fs";

const path = new URL("../openclaw.plugin.json", import.meta.url);
const manifest = JSON.parse(readFileSync(path, "utf8"));
if (manifest.contracts && Object.hasOwn(manifest.contracts, "tools")) {
  delete manifest.contracts.tools;
  if (Object.keys(manifest.contracts).length === 0) delete manifest.contracts;
  writeFileSync(path, `${JSON.stringify(manifest, null, 2)}\n`);
}
