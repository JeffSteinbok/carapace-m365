import { defineConfig } from "tsup";
import { definePluginConfig } from "carapace-plugin-sdk/tsup";

export default defineConfig({
  ...definePluginConfig(),
  noExternal: ["@carapace/m365-graph-auth"],
});
