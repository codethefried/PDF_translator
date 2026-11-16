import path from "node:path";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

// base: "./" keeps asset paths relative, so the same build works on
// GitHub Pages (https://<user>.github.io/<repo>/), any sub-folder, or a plain web server.
export default defineConfig({
  base: "./",
  plugins: [react(), tailwindcss()],
  resolve: { alias: { "@": path.resolve(__dirname, "src") } },
  build: { outDir: "dist", chunkSizeWarningLimit: 4000 },
});
