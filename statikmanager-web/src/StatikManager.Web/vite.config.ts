import { defineConfig } from "vite";

const apiTarget = "http://localhost:5156";

export default defineConfig({
  root: ".",
  build: {
    outDir: "../StatikManager.Api/wwwroot",
    emptyOutDir: true,
  },
  server: {
    port: 5173,
    strictPort: true,
    proxy: {
      "/api": {
        target: apiTarget,
        changeOrigin: true,
      },
    },
  },
});
