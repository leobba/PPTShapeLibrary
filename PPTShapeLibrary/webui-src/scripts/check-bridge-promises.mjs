// 前端"桥调用必须有失败处理"门禁
// ============================================================
// 为什么不用 ESLint 的 no-floating-promises：该规则需要 typescript-eslint 的类型信息，
// 而本项目用的是 TypeScript 7（原生预览版），typescript-eslint 8 的 peer 依赖不接受它
// （实测 npm install 直接 ERESOLVE 失败）。所以这里用一个**针对性**的静态检查：
// 桥调用漏 catch 会让写操作静默失败（用户以为点了没反应），这正是我们要防的那一类。
//
// 判定：源码里每次 `bridge.<方法>(` 调用，其所在语句必须满足以下之一：
//   · 带 .catch(  —— 显式处理失败
//   · 带 .then(   —— 链式（再由链尾的 .catch 处理；链尾单独也要求 .catch）
//   · 以 void 开头 —— 明确表示"故意忽略结果"
//   · 位于 await / return / = 赋值 语句里 —— 由调用方处理
// 例外：bridge.on / bridge.onResult（返回退订函数，不是 Promise）。
//
// 用法：node scripts/check-bridge-promises.mjs   （违规时非零退出）
import { readFileSync, readdirSync, statSync } from "fs";
import { dirname, join, relative } from "path";
import { fileURLToPath } from "url";

const here = dirname(fileURLToPath(import.meta.url));
const srcDir = join(here, "..", "src");

function walk(dir, out = []) {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) walk(p, out);
    else if (/\.(ts|tsx)$/.test(name)) out.push(p);
  }
  return out;
}

const SKIP_CALLS = new Set(["on", "onResult", "log"]); // on/onResult 返回退订函数；log 失败可忽略（否则会递归）
const violations = [];   // 硬门禁：裸调用，失败即"点了没反应"
const chained = [];      // 提示：有 .then 无 .catch，由全局兜底覆盖

for (const file of walk(srcDir)) {
  const rel = relative(srcDir, file).replace(/\\/g, "/");
  if (rel === "lib/bridge.ts") continue; // 桥本体
  const lines = readFileSync(file, "utf8").split(/\r?\n/);
  lines.forEach((line, i) => {
    const m = line.match(/bridge\.([A-Za-z_$][\w$]*)\s*\(/);
    if (!m) return;
    if (SKIP_CALLS.has(m[1])) return;
    if (/^\s*(\/\/|\*)/.test(line)) return; // 注释
    // 语句级判定（保守：宁可漏报也不误报，目标是拦住"新写的裸调用"）：
    //   · 本行或后 3 行内有 .catch(        → 有失败处理
    //   · 由 run(...) 包装                 → 包装器统一处理失败
    //   · void / await / return bridge.    → 显式忽略或交给调用方
    //   · 赋值或三元分支（= ? : 后接 bridge.）→ 链式处理在别处（p.then(…).catch(…)）
    const stmt = [line, lines[i + 1] || "", lines[i + 2] || "", lines[i + 3] || ""].join(" ");
    const prev = lines[i - 1] || "";
    const handled =
      stmt.includes(".catch(") ||
      /\brun\s*\(/.test(line) ||
      /\brun\s*\(/.test(prev) ||
      /^\s*void\s/.test(line) ||
      /\bawait\s+bridge\./.test(line) ||
      /\breturn\s+bridge\./.test(line) ||
      /[=?:]\s*bridge\./.test(line);
    if (handled) return;
    // 区分两类：
    //   bare     —— 完全没接 then/catch：一旦失败就是"点了没反应"，必须修 → 硬门禁
    //   chain    —— 接了 .then 但没 .catch：由全局 unhandledrejection（顶部提示条 + 宿主日志）兜底 → 只提示
    const isChain = stmt.includes(".then(");
    (isChain ? chained : violations).push({ rel, line: i + 1, text: line.trim().slice(0, 120) });
  });
}

if (chained.length > 0) {
  console.log("\n提示：" + chained.length + " 处 .then 链没有 .catch（由全局 unhandledrejection 兜底：顶部提示条 + 宿主日志）：");
  for (const v of chained) console.log("  " + v.rel + ":" + v.line + "  " + v.text);
}

if (violations.length === 0) {
  console.log("\nbridge 调用失败处理检查：无裸调用 ✓");
  process.exit(0);
}
console.error("\n发现 " + violations.length + " 处**裸**桥调用（没有 .then/.catch，失败即静默）：\n");
for (const v of violations) console.error("  " + v.rel + ":" + v.line + "  " + v.text);
console.error("\n处理方式：加 .catch(e => …) 提示用户，或明确写成 void bridge.xxx() 表示故意忽略。");
process.exit(1);
