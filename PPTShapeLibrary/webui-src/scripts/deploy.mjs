// 部署 webui-src/dist → ../WebUI（VSTO 加载目录）
import { cpSync, rmSync, mkdirSync, readdirSync, existsSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const here = dirname(fileURLToPath(import.meta.url));
const src = join(here, "..", "dist");
const dst = join(here, "..", "..", "WebUI");

if (!existsSync(src)) {
  console.error("dist 不存在，请先 npm run build");
  process.exit(1);
}
mkdirSync(dst, { recursive: true });
// 清掉旧 assets（避免旧 hash 残留）与旧的 index.html
for (const name of ["assets", "index.html"]) {
  const p = join(dst, name);
  if (existsSync(p)) rmSync(p, { recursive: true, force: true });
}
for (const name of readdirSync(src)) {
  cpSync(join(src, name), join(dst, name), { recursive: true });
}
console.log("已部署 dist → WebUI");
