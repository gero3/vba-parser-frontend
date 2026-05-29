import { defineConfig } from "vite";

export default defineConfig(({ mode }) => ({
  build: {
    assetsInlineLimit: 0,
    cssMinify: true,
    emptyOutDir: true,
    minify: "esbuild",
    sourcemap: mode !== "production",
    target: "es2022",
  },
}));
