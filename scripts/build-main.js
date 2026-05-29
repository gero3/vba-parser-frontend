import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { stripTypeScriptTypes } from "node:module";

const production = process.argv.includes("--prod");

const order = [
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

const banner = [
  "// Generated from split TypeScript sources by scripts/build-main.js.",
  "// Edit the .ts source files, then run: node scripts/build-main.js",
  "",
].join("\n");

const output = order
  .map((name) => {
    const source = readFileSync(`src/${name}.ts`, "utf8");
    return `// ---- ${name}.ts ----\n${stripTypeScriptTypes(source, { mode: "transform" })}`;
  })
  .join("\n\n");

writeFileSync("src/main.js", `${banner}${output}\n`);

if (production) {
  rmSync("dist", { force: true, recursive: true });
  mkdirSync("dist/src", { recursive: true });
  writeFileSync("dist/src/main.js", minifyJs(output));
  writeFileSync("dist/src/styles.css", minifyCss(readFileSync("src/styles.css", "utf8")));
  writeFileSync("dist/index.html", minifyHtml(readFileSync("index.html", "utf8")));
  if (existsSync("README.md")) copyFileSync("README.md", "dist/README.md");
}

function minifyJs(source) {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .replace(/^\s*\/\/.*$/gm, "")
    .replace(/\s+/g, " ")
    .replace(/\s*([{}()[\];,:?+\-*/%<>=|&])\s*/g, "$1")
    .trim();
}

function minifyCss(source) {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .replace(/\s+/g, " ")
    .replace(/\s*([{}:;,>])\s*/g, "$1")
    .trim();
}

function minifyHtml(source) {
  return source
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/\s+/g, " ")
    .replace(/>\s+</g, "><")
    .trim();
}
