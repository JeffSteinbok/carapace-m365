import { defineConfig } from "tsup";
import type { Plugin as EsbuildPlugin } from "esbuild";

// esbuild strips the "node:" prefix when bundling node built-ins that don't
// have a corresponding npm package (e.g. node:sqlite, added in Node 22).
// This plugin intercepts the bare "sqlite" specifier and marks it as an
// external that keeps the "node:sqlite" path, which Node resolves natively.
const keepNodeSqlitePlugin: EsbuildPlugin = {
  name: "keep-node-sqlite",
  setup(build) {
    build.onResolve({ filter: /^(node:)?sqlite$/ }, () => ({
      path: "node:sqlite",
      external: true,
    }));
  },
};

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  platform: "node",
  target: "node20",
  bundle: true,
  clean: true,
  sourcemap: true,
  noExternal: ["@carapace/m365-graph-auth", "@carapace/m365-mail-store"],
  esbuildPlugins: [keepNodeSqlitePlugin],
});
