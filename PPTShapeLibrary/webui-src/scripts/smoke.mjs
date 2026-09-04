import { readFileSync, existsSync, readdirSync } from "fs";
const html = readFileSync("dist/index.html", "utf8");
const refs = [...html.matchAll(/(?:src|href)="\.\/([^"]+)"/g)].map((m) => m[1]);
console.log("index.html 引用:", refs.join(", ") || "(无)");
const missing = refs.filter((r) => !existsSync("dist/" + r));
console.log(missing.length ? "缺失: " + missing.join(", ") : "所有引用文件存在 ✓");
const jsFiles = readdirSync("dist/assets").filter((f) => f.endsWith(".js"));
for (const f of jsFiles) {
  const js = readFileSync("dist/assets/" + f, "utf8");
  const fonts = [...js.matchAll(/["']\.\/([^"']+\.woff2)["']/g)].map((m) => m[1]);
  const missF = fonts.filter((x) => !existsSync("dist/assets/" + x));
  if (missF.length) console.log(f + " 缺字体: " + missF.join(", "));
}
console.log("JS 内字体引用检查完成, JS 文件数:", jsFiles.length);
