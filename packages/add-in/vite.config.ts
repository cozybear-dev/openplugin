import react from "@vitejs/plugin-react";
import { defineConfig, type UserConfig } from "vite";

export default defineConfig(async ({ command }): Promise<UserConfig> => {
  const config: UserConfig = {
    plugins: [react()],
    server: {
      port: 3000,
      strictPort: true,
      headers: { "Access-Control-Allow-Origin": "*" }
    },
    preview: {
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" }
    },
    build: {
      rollupOptions: {
        input: {
          taskpane: "taskpane.html",
          commands: "commands.html",
          functions: "src/functions.ts"
        },
        output: {
          entryFileNames: (chunk) =>
            chunk.name === "functions" ? "functions.js" : "assets/[name]-[hash].js"
        }
      }
    }
  };

  if (command === "serve") {
    try {
      const devCerts = await import("office-addin-dev-certs");
      const https = await (devCerts.default ?? devCerts).getHttpsServerOptions();
      config.server = { ...config.server, https };
      config.preview = { ...config.preview, https };
    } catch (err) {
      console.warn("HTTPS developer certificates unavailable; serving HTTP.", err);
    }
  }

  return config;
});
