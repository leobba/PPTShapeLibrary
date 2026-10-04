# 办公资源网（bangongziyuan.com）模板批量下载器

一个尊重网站机制的 Python 下载器：它只自动化**你本人已登录且有下载权限**时的“立即下载”点击，
不破解会员校验、不绕过权限。请遵守网站服务条款与每日下载次数限制，控制频率、按需下载。

> ⚠️ 只用于下载你账号有权下载的模板。批量、高频下载可能触发网站的风控（“今日下载频繁、
> 疑似恶意下载”提示），脚本遇到这类提示会礼貌停止，请隔日再续传。

---

## 一、环境准备

- Python 3.8+
- 安装依赖：

```bash
pip install requests
```

## 二、第一步：把浏览器的登录 Cookie 给脚本

打开 https://www.bangongziyuan.com 并**登录你的账号**（确保有对应下载权限/会员），然后任选一种方式导出 Cookie：

**方式 A（推荐，最简单）：复制 Cookie 头，存成文本文件**
1. 在网页上按 `F12` 打开开发者工具 → 切到 **Network（网络）** 标签。
2. 刷新页面，随便点一个请求（比如文档请求 `www.bangongziyuan.com`）。
3. 在请求详情里找 **Request Headers → Cookie:**，整行复制。
4. 在本目录新建文件 `cookies_header.txt`，把复制的内容整行粘贴进去保存。
5. 在 `config.json` 里写：
   ```json
   "cookie_string": "@cookies_header.txt"
   ```
   （`@` 表示读取文件；这样即使 Cookie 里含引号/花括号也不会弄乱 JSON。）

**方式 B：用扩展导出 Netscape 格式 cookies.txt**
1. Chrome/Edge 安装 “Cookie-Editor” 扩展。
2. 在 bangongziyuan.com 页面打开扩展 → Export → **Netscape** 格式，保存为 `cookies.txt`。
3. 放到本目录，在 `config.json` 里填 `"cookie_file": "cookies.txt"`。

> Cookie 会过期；当提示“guid 校验失败”或下载被拒时，重新导出一次即可。

## 三、配置 `config.json`

把 `config.example.json` 复制为 `config.json` 并编辑：

```jsonc
{
  "cookie_string": "@cookies_header.txt",   // @文件 读取 Cookie 头；或 "cookie_file": "cookies.txt"

  "out_dir": "downloads",    // 文件保存目录
  "delay": 2.5,              // 每次下载间隔秒数（越大越安全）
  "timeout": 60,
  "retries": 3,              // 文件下载失败重试次数
  "limit": 0,                // 本轮最多下载条数，0=不限
  "max_pages": 0,            // 每个列表源最多翻页数，0=全部

  "sources": [               // 要抓的列表页（可多个）
    {
      "name": "工作总结",     // 保存用的子文件夹名
      "url": "https://www.bangongziyuan.com/ppt/cat1-11-stu-0.html",
      "include": ["述职", "汇报"],   // 标题必须包含其一；省略则不限
      "exclude": ["卡通"]            // 标题不得包含；省略则不限
    }
  ],

  "keywords": [],            // 便捷：按关键词搜索（等价于在站内搜索框输入）

  "include": [],             // 全局标题过滤（对上面所有源生效）
  "exclude": []
}
```

### 常用列表页 URL（在网站分类导航里也能复制到）

| 栏目 | URL 前缀 |
|---|---|
| PPT模板 → 工作总结 | `https://www.bangongziyuan.com/ppt/cat1-11-stu-0.html` |
| PPT模板 → 教育培训 | `https://www.bangongziyuan.com/ppt/cat1-10-stu-0.html` |
| PPT模板 → 述职汇报 | `https://www.bangongziyuan.com/ppt/cat1-178-stu-0.html` |
| PPT模板 → 论文答辩 | `https://www.bangongziyuan.com/ppt/cat1-23-stu-0.html` |
| PPT模板 → 个人简历 | `https://www.bangongziyuan.com/ppt/cat1-5-stu-0.html` |
| PPT模板 → 企业介绍 | `https://www.bangongziyuan.com/ppt/cat1-7-stu-0.html` |
| PPT模板（全部） | `https://www.bangongziyuan.com/ppt/` |
| PPT图表 | `https://www.bangongziyuan.com/chart/` |
| Word模板 | `https://www.bangongziyuan.com/word/` |
| Excel表格 | `https://www.bangongziyuan.com/excel/` |

URL 中的排序参数可改：`stu-0`=默认、`stu-1`=最新、`stu-3`=下载量。脚本会自动翻页抓到底。

## 四、使用

### 1) 校验 Cookie / 账号（强烈建议先做）

```bash
python bzy_downloader.py check -c config.json
```

正常会输出：`用户 id=xxx 类型=... 到期=...`；若是“guid 校验失败/会员不存在”则 Cookie 失效。

### 2) 只生成候选清单（不下载，先看看会选中什么）

```bash
python bzy_downloader.py list -c config.json
```

结果写入 `downloads/queue.json`，控制台会打印前 20 条标题。

### 3) 批量下载（可断点续传）

```bash
python bzy_downloader.py download -c config.json
```

- 已下载成功的不再重复下载（记录在 `downloads/state.json`，文件名 `proId_标题.原始后缀`）。
- 失败条目记录在 `state.json` 的 `fail`，**重跑同一命令即续传**。
- 服务器返回“下载频繁/次数用完”时脚本会停止，次日再跑即可。

### 4) 一条龙：抓取 + 下载

```bash
python bzy_downloader.py all -c config.json
```

### 常用附加参数（覆盖 config.json）

```bash
python bzy_downloader.py download -c config.json --limit 20 --delay 5 --include "述职,汇报" --exclude "卡通"
```

---

## 五、目录结构（运行后）

```
bangongziyuan_downloader/
├── bzy_downloader.py       # 主程序
├── _self_test.py           # 纯函数自测(无网络):  python _self_test.py
├── _integration_test.py    # 下载流程自测(模拟接口): python _integration_test.py
├── _resume_test.py         # 断点续传自测: python _resume_test.py
├── config.json             # 你的配置(由 config.example.json 复制)
├── cookies_header.txt      # (可选)方式 A 的 Cookie 头文件
├── cookies.txt             # (可选)方式 B 的 Netscape cookie 文件
└── downloads/
    ├── queue.json          # 候选清单（list 生成）
    ├── state.json          # 完成/失败状态（续传依据）
    └── 工作总结/            # 每个 source 一个子目录，文件名 = proId_标题.原格式
        ├── 12345_蓝色述职报告PPT模板.zip
        └── ...
```

## 六、常见问题

| 现象 | 处理 |
|---|---|
| check 提示 guid 校验失败 | Cookie 过期或复制不全，重新登录后导出 |
| 下载全部被拒：提示需会员/VIP/过期 | 当前账号无对应权限；请用有权限的账号 |
| 提示“今日下载频繁/次数已用完” | 触发了频率限制，**停止并次日重跑**（可续传） |
| 某几条失败 | 不影响整体；重跑命令自动续传失败项 |
| 控制台中文乱码 | 在终端运行 `chcp 65001` 后重试，或使用 Windows Terminal |

## 免责声明

本工具仅用于个人对**自己有权下载**的内容做批量备份，请遵守目标网站的用户协议、
会员规则与版权规定，自行承担使用责任。
