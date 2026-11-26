import { execSync } from "node:child_process";
import { readFileSync, writeFileSync } from "node:fs";

if (process.argv.includes("--restore")) {
  // Regenerate the build manifest after npm has captured the archive contents.
  execSync("npm run --silent generate:plugin", { stdio: ["ignore", "ignore", "inherit"] });
  process.exit(0);
}

const path = new URL("../openclaw.plugin.json", import.meta.url);
const manifest = JSON.parse(readFileSync(path, "utf8"));

// The SDK's reusable CI still requires entry in the generated build manifest.
// OpenClaw loads the packaged adapter from package.json's openclaw.extensions.
delete manifest.entry;
delete manifest.openclaw;
delete manifest.source;

writeFileSync(path, JSON.stringify(manifest, null, 2) + "\n");
