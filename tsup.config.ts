import { defineConfig } from "tsup";
import { definePluginConfig } from "carapace-plugin-sdk/tsup";

const base = definePluginConfig();

export default defineConfig({
  ...base,
  noExternal: [...(base.noExternal as RegExp[]), "@carapace/m365-graph-auth"],
});
