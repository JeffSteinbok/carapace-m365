import { readFileSync, writeFileSync } from "node:fs";

const path = new URL("../openclaw.plugin.json", import.meta.url);
const manifest = JSON.parse(readFileSync(path, "utf8"));

// The reusable SDK workflow validates entry immediately after a fresh build.
// Packaging updates that generated file for distribution. OpenClaw loads the
// packaged adapter from package.json's openclaw.extensions.
delete manifest.entry;
delete manifest.openclaw;
delete manifest.source;

writeFileSync(path, JSON.stringify(manifest, null, 2) + "\n");
