# PPTShapeLibrary 素材包格式规范（Material Pack Spec）

> 版本：1.1 · 状态：实施中（向后兼容 v1）
> 本文档定义"素材包"（`.ppack.zip`）的文件结构、元数据字段与导入导出语义，
> 目标是让素材在 GitHub / 群聊 / 网盘等渠道**自由流通、可靠互导**。

---

## 1. 概述

素材包是把一批收藏（形状 / 样式）打包成单个文件的容器格式，扩展名 `.ppack.zip`。
一个素材包 = 一个标准 zip，内含：

```
xxx.ppack.zip
├── manifest.json            # 包描述：名称/作者/许可/规范版本 + 条目清单
└── items/
    └── {id}/                # 与本地收藏库条目目录同构（每个条目一个子目录）
        ├── meta.json        # 条目元数据（必需）
        ├── shape.pptx       # 形状条目（Kind=Shape 必需；单页无占位符形状）
        ├── style.json       # 样式条目（Kind=Style 必需；样式 JSON）
        └── preview.png      # 预览缩略图（推荐；缺失时前端显示占位）
```

> 说明：`shape.pptx` 是 PowerPoint 单页演示文稿（PowerPoint 生态内可读）。
> 本格式定位为 **PowerPoint 生态内流通容器**，不承诺跨软件（Keynote/Canva/SVG）可读。

## 2. manifest.json

| 字段 | 类型 | 必需 | 说明 |
|---|---|---|---|
| `format` | string | 是 | 固定 `"pptshape-pack-v1"` |
| `specVersion` | string | 否(v1.1+) | 规范版本，当前 `"1.1"`；旧包可缺失 |
| `name` | string | 是 | 包名（通常为文件名去扩展名） |
| `author` | string | 否(v1.1+) | 作者署名（导出时来自插件设置"作者署名"） |
| `license` | string | 否(v1.1+) | 内容许可，默认 `"CC BY 4.0"` |
| `sourceUrl` | string | 否 | 素材来源仓库/主页 URL（可选） |
| `description` | string | 否 | 一句话描述 |
| `exportedAt` | string | 是 | 导出时间 `yyyy-MM-dd HH:mm:ss` |
| `items` | array | 是 | 条目清单：`{ id, name, kind }` |

### 许可约定（内容层，与代码许可证分开）
- 代码：见仓库 `LICENSE`（Apache-2.0）。
- 素材内容：导出包默认标注 `CC BY 4.0`（允许复制/再分发/商用/改编，需署名）。
  社区分享素材时可在导出前在插件设置填写作者署名；包内 `license` 字段请勿移除。

## 3. 条目（items/{id}/）

### 3.1 meta.json 核心字段
| 字段 | 说明 |
|---|---|
| `id` | 条目 id，规范：`yyyyMMdd-HHmmss-` + 十六进制随机后缀（新版本生成 12 位；导入仍兼容旧版 6 位），仅字母数字与连字符 |
| `name` | 显示名（≤120 字符） |
| `kind` | `"Shape"`（形状）或 `"Style"`（样式） |
| `category` | 二级分类（组内子分类） |
| `group` | 一级分组（导入时若目标组不存在会自动建；条目也可导入到「我的」） |
| `tags` | 标签（逗号/分号分隔存储） |
| `createdAt` | 创建时间 |

### 3.2 文件约定
- `shape.pptx`：单页、无占位符、形状已归一化到页面左上 (0,0) 附近，便于插件插入时居中/缩放。
- `style.json`：与 ShapeStyleExtractor 输出一致（含填充/线条/阴影/文字等样式属性）。
- `preview.png`：建议宽高比与内容一致、透明背景、≤640px 为宜。

## 4. 导入语义（保证可靠流通）

1. **id 冲突自动重写**：本地已存在同 id 条目时，导入端为新条目生成新 id，绝不覆盖本地数据。
2. **孤儿条目跳过**：仅导入 manifest 列出的 id（无 manifest 的 zip 按目录结构全量导入）。
3. **安全校验（导入端强制）**：
   - 条目路径必须恰为 `items/{id}/{file}` 三段；
   - `id` 必须匹配白名单 `^[0-9A-Za-z_-]+$`；
   - 文件名不得含路径分隔符 / `..` / 冒号；
   - 落盘前对完整路径做前缀校验（禁止越出库目录）。
4. **重名可去重**：导入后可用「库管理 → 去重清理」扫描同类型/数量/尺寸的完全重复项。

## 5. 社区流通推荐路径（A+B 形态）

### A. 点对点分享
1. 插件内「库管理 → 导出素材包」（或批量条「导出素材包…」）；
2. 得到 `xxx.ppack.zip`，经任意渠道发送；
3. 对方「库管理 → 导入素材包」→ 条目进入素材库（含作者/许可来源信息）。

### B. 素材合集仓库（GitHub 等）
推荐仓库结构：

```
my-ppt-shapes/
├── README.md            # 素材清单 + 截图 + 导入说明
├── LICENSE              # 素材内容许可（如 CC BY 4.0）
└── packs/
    ├── 答辩模板包.ppack.zip
    ├── 扁平箭头集.ppack.zip
    └── ...
```

用户使用：下载仓库 → 解压后把 `packs/*.ppack.zip` 拷到
`%APPDATA%\PPTShapeLibrary\packs\`（或任一目录）→ 插件「导入素材包」选择文件导入。

### 一键导入目录（建议体验，v1.2 规划）
插件「库管理 → 导入素材包目录」：扫描指定目录下所有 `*.ppack.zip` 批量导入，
使"clone 仓库 → 一键导入"成为现实。

## 6. 兼容性与演进

- 解析器必须容忍：缺失 `specVersion`/`author`/`license` 的旧包（v1 全部可选）。
- 新增字段一律"可选 + 向后兼容"；破坏性变更必须升 `specVersion` 大版本。
- 官方导出/导入实现：`PackService.cs`（ExportPack / ImportPack）。
- 社区实现校验：可用任意语言解 zip 读 manifest 并核对 3.2/4 节约定（推荐仓库配 CI）。
