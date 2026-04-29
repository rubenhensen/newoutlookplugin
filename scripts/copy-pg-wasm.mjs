// Copies @e4a/pg-wasm/bundler/index_bg.wasm into @e4a/pg-js/dist so the
// inlined wasm-bindgen `new URL("index_bg.wasm", import.meta.url)` call
// resolves under webpack. The file is missing from the @e4a/pg-js npm
// package — this is a known gap that the Thunderbird build sidesteps by
// using esbuild's bundler. Webpack instead requires the asset to live
// next to the importing module.

import { copyFileSync, existsSync, mkdirSync, statSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const src = resolve(root, "node_modules/@e4a/pg-wasm/bundler/index_bg.wasm");
const dst = resolve(root, "node_modules/@e4a/pg-js/dist/index_bg.wasm");

if (!existsSync(src)) {
  console.warn(`[postinstall] @e4a/pg-wasm WASM not found at ${src} — skipping.`);
  process.exit(0);
}

mkdirSync(dirname(dst), { recursive: true });
copyFileSync(src, dst);
const sz = statSync(dst).size;
console.log(`[postinstall] Copied pg-wasm into pg-js (${sz} bytes).`);
