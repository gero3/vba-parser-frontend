import { readFileSync } from "node:fs";
import { defineConfig } from "vite";

const appSourceOrder = [
  "types",
  "app",
  "extract",
  "results-ui",
  "zip-export",
  "model",
  "designer-ui",
  "analysis-ui",
  "runtime",
];

function vbaParserAppPlugin() {
  const virtualId = "virtual:vba-parser-app";
  const resolvedVirtualId = `\0${virtualId}`;

  return {
    name: "vba-parser-app",
    resolveId(id) {
      if (id === virtualId) return resolvedVirtualId;
      return undefined;
    },
    load(id) {
      if (id !== resolvedVirtualId) return undefined;
      return appSourceOrder
        .map((name) => {
          const source = readFileSync(`src/${name}.ts`, "utf8");
          return `// ---- ${name}.ts ----\n${source}`;
        })
        .join("\n\n");
    },
  };
}

export default defineConfig(({ mode }) => ({
  plugins: [vbaParserAppPlugin()],
  build: {
    assetsInlineLimit: 0,
    cssMinify: true,
    emptyOutDir: true,
    minify: "esbuild",
    sourcemap: mode !== "production",
    target: "es2022",
  },
}));
