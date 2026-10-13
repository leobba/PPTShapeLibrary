// 部署 webui-src/dist → ../WebUI（VSTO 加载目录）
// 采用"先清空目标再整体拷贝"（镜像）语义：只删 assets/index.html 会让历史遗留文件
// （如旧版静态前端的 css/、js/ 空目录）永久留在产物里，并被 csproj 的 WebUI\**\* 打进安装包。
import { cpSync, rmSync, mkdirSync, readdirSync, existsSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const here = dirname(fileURLToPath(import.meta.url));
const src = join(here, "..", "dist");
const dst = join(here, "..", "..", "WebUI");

if (!existsSync(src) || !existsSync(join(src, "index.html"))) {
  console.error("dist 不完整（缺少 index.html），请先 npm run build && npm run smoke");
  process.exit(1);
}
// 拷贝前先校验 index.html 引用的资源都存在，避免把坏产物部署进加载目录
const { readFileSync } = await import("fs");
const html = readFileSync(join(src, "index.html"), "utf8");
const refs = [...html.matchAll(/(?:src|href)="\.\/([^"]+)"/g)].map((m) => m[1]);
const missing = refs.filter((r) => !existsSync(join(src, r)));
if (missing.length) {
  console.error("index.html 引用的文件缺失，拒绝部署: " + missing.join(", "));
  process.exit(1);
}

rmSync(dst, { recursive: true, force: true });   // 镜像：清掉一切旧产物
mkdirSync(dst, { recursive: true });
for (const name of readdirSync(src)) {
  cpSync(join(src, name), join(dst, name), { recursive: true });
}
console.log("已部署 dist → WebUI（镜像，已清理旧产物）");

// 关键：VSTO 运行时加载的是 AppDomain.BaseDirectory\WebUI，也就是 bin\<Config>\WebUI\
// （csproj 用 WebUI\**\* 拷过去）。只更新源码树 WebUI\ 的话，改完前端跑 build:all
// 后 PowerPoint 里仍然是旧界面——实测踩过这个坑（CSP 改动没生效）。
// 这里把已经存在的 bin 输出目录一并镜像更新，让 `npm run build:all` 一步到位。
const root = join(here, "..", "..");
for (const cfg of ["Debug", "Release"]) {
  const outDir = join(root, "bin", cfg, "WebUI");
  if (!existsSync(outDir)) continue;          // 没构建过就跳过（首次构建时 csproj 会拷）
  rmSync(outDir, { recursive: true, force: true });
  mkdirSync(outDir, { recursive: true });
  for (const name of readdirSync(src)) {
    cpSync(join(src, name), join(outDir, name), { recursive: true });
  }
  console.log("已同步 → bin\\" + cfg + "\\WebUI（插件实际加载目录）");
}
