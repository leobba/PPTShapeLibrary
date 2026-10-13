// 产物完整性自检：作为构建门禁使用（发现问题必须让 build:all 失败，而不是只打印一行）
//   node scripts/smoke.mjs
import { readFileSync, existsSync, readdirSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const here = dirname(fileURLToPath(import.meta.url));
const dist = join(here, "..", "dist");
const problems = [];

if (!existsSync(join(dist, "index.html"))) {
  console.error("smoke: dist/index.html 不存在，请先 npm run build");
  process.exit(1);
}

const html = readFileSync(join(dist, "index.html"), "utf8");
const refs = [...html.matchAll(/(?:src|href)="\.\/([^"]+)"/g)].map((m) => m[1]);
console.log("index.html 引用:", refs.join(", ") || "(无)");
if (refs.length === 0) problems.push("index.html 没有引用任何本地资源（构建产物异常）");

const missing = refs.filter((r) => !existsSync(join(dist, r)));
if (missing.length) problems.push("index.html 引用的文件缺失: " + missing.join(", "));

// 扫描 JS 与 CSS 里的相对资源引用（字体/图片）：此前只扫 .js，而 woff2 引用实际在 CSS 里，
// 等于字体检查从未生效。
const assetDir = join(dist, "assets");
const assetFiles = existsSync(assetDir)
  ? readdirSync(assetDir).filter((f) => /\.(js|css)$/.test(f))
  : [];
if (assetFiles.length === 0) problems.push("dist/assets 下没有 js/css 产物");
for (const f of assetFiles) {
  const text = readFileSync(join(assetDir, f), "utf8");
  const urls = [...text.matchAll(/["'(]\.\/([^"')]+\.(?:woff2|woff|ttf|otf|png|jpe?g|svg|gif))["')]/g)].map(
    (m) => m[1]
  );
  const miss = [...new Set(urls)].filter((u) => !existsSync(join(assetDir, u)));
  if (miss.length) problems.push(`${f} 引用的资源缺失: ${miss.join(", ")}`);
}
console.log("JS/CSS 内资源引用检查完成, 文件数:", assetFiles.length);

if (problems.length) {
  console.error("\nsmoke FAILED:");
  for (const p of problems) console.error("  - " + p);
  process.exit(1);
}
console.log("smoke OK");
