#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
办公资源网 (www.bangongziyuan.com) PPT/Word/Excel 模板批量下载器
=============================================================

原理（已对网站前端 JS 逆向确认）：
  1. 列表页/搜索页按页渲染模板卡片，分页 URL 为把末尾 ".html" 替换为 "-pn-页码.html"，
     每页 40 条；页面上有隐藏字段 #pptAll(总页数)、#pptNum(当前页码)。
  2. 卡片里的模板 id (proId) 与标题，点击“立即下载”时前端 JS 会
     POST /inter/addDownload 并携带 userId/proId/userBy* 等字段。
  3. 接口返回 status==1 时 data.path 即文件相对地址，真实文件 URL =
     https://image1.bangongziyuan.com + path(去掉其自带的 http://image1.bangongziyuan.com 前缀)。
  4. 登录态：浏览器登录后 cookie 里有 guid（会话）与 userInfo（用户 JSON），
     可用 POST /inter/getUserByguid {guid} 刷新用户信息（含下载权限/到期时间）。

本工具只自动化“你本人在网站上已登录、且有下载权限”时的点击行为，
不绕过会员校验、不破解限制；请遵守网站服务条款，控制频率，尊重每日下载次数上限。

用法：
  python bzy_downloader.py check                       # 校验 cookie/账号
  python bzy_downloader.py list    -c config.json      # 只抓取目录生成 queue.json
  python bzy_downloader.py download -c config.json     # 按 queue.json 批量下载（可续传）
  python bzy_downloader.py all     -c config.json      # 抓目录 + 下载 一步完成
更多参数见 --help / README.md
"""

import argparse
import html as htmlmod
import json
import os
import random
import re
import sys
import tempfile
import time
from urllib.parse import quote, unquote, urljoin

try:
    import requests
except ImportError:  # pragma: no cover
    sys.exit("缺少依赖 requests，请先运行:  pip install requests")

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SITE = "https://www.bangongziyuan.com"
CDN_BASE = "https://image1.bangongziyuan.com"
DEFAULT_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

TITLE_RE = re.compile(
    r'<a\b(?=[^>]*class="title")(?=[^>]*data-id="(\d+)")[^>]*href="([^"]*detail-\d+\.html[^"]*)"[^>]*>(.*?)</a>',
    re.S | re.I,
)
HIDDEN_RE = re.compile(r'<input[^>]*id="([^"]+)"[^>]*value="([^"]*)"[^>]*>', re.I)
PAGE_NUM_RE = re.compile(r"-pn-\d+", re.I)


# --------------------------------------------------------------------------- #
#  基础 HTTP / Cookie                                                          #
# --------------------------------------------------------------------------- #
def make_session(cookie_str="", cookie_file=""):
    s = requests.Session()
    s.headers.update({
        "User-Agent": DEFAULT_UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    })
    if cookie_file and os.path.isfile(cookie_file):
        load_netscape_cookies(s, cookie_file)
        print(f"[cookie] 已从文件加载: {cookie_file}")
    if cookie_str:
        if cookie_str.startswith("@"):   # "@文件名" = 从文件读取原始 Cookie 头
            cf = cookie_str[1:]
            if os.path.isfile(cf):
                with open(cf, encoding="utf-8", errors="replace") as f:
                    cookie_str = f.read().strip()
                print(f"[cookie] 已从文件加载 Cookie 头: {cf}")
            else:
                print(f"[!] cookie 文件不存在: {cf}")
                cookie_str = ""
        if cookie_str:
            load_cookie_header(s, cookie_str)
            print("[cookie] 已从字符串加载 cookie")
    return s


def load_netscape_cookies(session, path):
    """解析 Netscape cookies.txt（可用 Cookie-Editor 等扩展导出）。"""
    from http.cookiejar import MozillaCookieJar
    jar = MozillaCookieJar(path)
    jar.load(ignore_discard=True, ignore_expires=True)
    for c in jar:
        session.cookies.set(c.name, c.value, domain=c.domain, path=c.path)
    return session


def load_cookie_header(session, header):
    """解析浏览器“复制为 Cookie 头”格式:  name1=value1; name2=value2; ..."""
    for part in header.split(";"):
        part = part.strip()
        if not part or "=" not in part:
            continue
        name, _, value = part.partition("=")
        session.cookies.set(name.strip(), value.strip(), domain=".bangongziyuan.com", path="/")
    return session


def cookie_value(session, name):
    for c in session.cookies:
        if c.name.lower() == name.lower():
            return c.value
    return None


def http_get(session, url, referer=SITE, timeout=30):
    headers = {"Referer": referer}
    r = session.get(url, headers=headers, timeout=timeout)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or "utf-8"
    return r


def api_post(session, path, data, referer=SITE, timeout=30):
    headers = {
        "Referer": referer,
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json, text/javascript, */*; q=0.01",
    }
    r = session.post(urljoin(SITE, path), data=data, headers=headers, timeout=timeout)
    r.raise_for_status()
    try:
        return r.json()
    except ValueError:
        print("[!] 接口非 JSON 返回：", r.text[:300])
        return {"status": 0, "msg": r.text[:200], "_raw": True}


# --------------------------------------------------------------------------- #
#  用户信息                                                                     #
# --------------------------------------------------------------------------- #
def fetch_user(session):
    """返回 (user_dict, summary_str)。user_dict 为 /inter/getUserByguid 或 userInfo cookie 的 JSON。"""
    guid = cookie_value(session, "guid")
    user = None
    src = ""

    # 优先用 cookie 里的 userInfo（前端 downModel 实际就发这里面的字段）
    ui = cookie_value(session, "userInfo")
    if ui:
        for cand in (ui, unquote(ui) if "%" in ui else ui):
            try:
                parsed = json.loads(cand)
                if isinstance(parsed, dict) and parsed.get("id"):
                    user = parsed
                    src = "cookie:userInfo"
                    break
            except Exception:
                continue

    # 再用 guid 向服务端刷新一次（校验会话有效性 + 拿最新会员信息）
    if guid:
        try:
            j = api_post(session, "/inter/getUserByguid", {"guid": guid})
            if j.get("status") == 1 and isinstance(j.get("data"), dict):
                fresh = j["data"]
                if user is None:
                    user = fresh
                    src = "getUserByguid"
                else:
                    user.update(fresh)   # 以服务端为准补全
                src += "+getUserByguid"
            else:
                # guid 在且校验失败 => 会话凭证不可信，宁可失败提示
                return None, f"guid 校验失败: {j.get('msg')}（cookie 可能已过期，请重新导出）"
        except Exception as e:
            return None, f"getUserByguid 请求失败: {e}"
    elif user is None:
        return None, "未找到登录凭证：cookie 中缺少 guid / userInfo，请先登录网站并导出 Cookie。"

    if user is None:
        return None, "未找到登录凭证：cookie 中缺少 guid / userInfo，请先登录网站并导出 Cookie。"

    uid = user.get("id") or user.get("userId")
    member = user.get("member_type")
    expire = user.get("userByExpireDate") or user.get("expireDate") or ""
    utype = user.get("userBytype") or user.get("user_type") or ""
    limit = user.get("userByDownloadLimit") or user.get("downloadLimit") or "?"
    summary = (f"用户 id={uid}  类型={member or utype or '普通'}  "
               f"到期={expire or '无'}  下载额度字段={limit}")
    return user, summary


def build_download_params(user, pro_id, request_id="", ops_request_misc=""):
    """构造与网站前端 downModel() 一致的 addDownload 参数。"""
    return {
        "userId": str(user.get("id") or user.get("userId") or ""),
        "proId": str(pro_id),
        "inviterLevel": user.get("inviterLevel", 0) or 0,
        "userByExpireDate": user.get("userByExpireDate", "") or "",
        "userBytype": user.get("userBytype", "normal") or "normal",
        "userByDownloadLimit": user.get("userByDownloadLimit", 0) or 0,
        "userByCollectLimit": user.get("userByCollectLimit", 100) or 100,
        "request_id": request_id or "",
        "ops_request_misc": ops_request_misc or "",
    }


# --------------------------------------------------------------------------- #
#  目录抓取                                                                     #
# --------------------------------------------------------------------------- #
def page_url(base, n):
    """分类/搜索页翻页：末尾 .html 前插入 -pn-N。如 cat1-11-stu-0.html -> cat1-11-stu-0-pn-2.html"""
    if n <= 1:
        return base
    b = PAGE_NUM_RE.sub("", base)
    if b.endswith(".html"):
        return b[:-5] + f"-pn-{n}.html"
    return f"{b.rstrip('/')}-pn-{n}.html"


def hidden_values(html):
    out = {}
    for m in HIDDEN_RE.finditer(html):
        out[m.group(1)] = m.group(2)
    return out


def parse_items(html, page_url_str):
    """从一页 HTML 中解析模板卡片 -> [{proId, title, detail, ...}]"""
    items = []
    for m in TITLE_RE.finditer(html):
        pro_id, href, raw_title = m.group(1), m.group(2), m.group(3)
        title = htmlmod.unescape(re.sub(r"<[^>]+>", "", raw_title))
        title = re.sub(r"\s+", " ", title).strip()
        if not pro_id or not title:
            continue
        items.append({
            "proId": pro_id,
            "title": title,
            "detail": urljoin(SITE, href),
            "page": page_url_str,
        })
    # 去重（同一页可能重复出现）
    seen, uniq = set(), []
    for it in items:
        if it["proId"] not in seen:
            seen.add(it["proId"])
            uniq.append(it)
    return uniq


def crawl_source(session, url, max_pages=0, delay=(0.5, 1.2), timeout=30):
    """
    翻页抓取一个列表/搜索源。
    返回 (items, total_pages_seen, stop_reason)
    """
    url = url.strip()
    items = []
    n = 1
    page1_ids = None
    stop = ""
    while True:
        u = page_url(url, n)
        try:
            r = http_get(session, u, timeout=timeout)
            html = r.text
        except Exception as e:
            stop = f"请求失败 {u}: {e}"
            break

        page_items = parse_items(html, u)
        h = hidden_values(html)
        total = int(h.get("pptAll") or 0)
        cur = int(h.get("pptNum") or n)
        # 该页隐藏的请求标记(搜索页才有)，addDownload 时原样回传
        for it in page_items:
            it["request_id"] = h.get("request_id", "") or ""
            it["ops_request_misc"] = h.get("ops_request_misc", "") or ""

        if page_items:
            if n == 1:
                page1_ids = {it["proId"] for it in page_items}
            # 防死循环：非第一页却与第一页完全重复 -> 结束
            elif page1_ids and {it["proId"] for it in page_items} == page1_ids and n > 1:
                stop = f"第 {n} 页与第 1 页内容重复，提前停止"
                break
            items.extend(page_items)
            print(f"    [页 {n}] {u}  条目 {len(page_items)}  总页数信号 {total or '?'}")

        if total and n >= total:
            stop = f"已到最后一页({total})"
            break
        if not page_items:
            stop = f"第 {n} 页无条目，提前停止"
            break
        if max_pages and n >= max_pages:
            stop = f"达到 max_pages={max_pages}"
            break
        n += 1
        time.sleep(random.uniform(*delay))

    # 全局去重（保留首次出现顺序）
    seen, uniq = set(), []
    for it in items:
        if it["proId"] not in seen:
            seen.add(it["proId"])
            uniq.append(it)
    return uniq, n, stop


def matches_filters(item, include, exclude):
    title = item["title"].lower()
    if include and not any(k.lower() in title for k in include):
        return False
    if exclude and any(k.lower() in title for k in exclude):
        return False
    return True


# --------------------------------------------------------------------------- #
#  下载                                                                        #
# --------------------------------------------------------------------------- #
ILLEGAL = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def safe_name(name, max_len=120):
    name = ILLEGAL.sub("_", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name[:max_len] or "untitled"


def ext_of_url(u):
    path = u.split("?", 1)[0].rstrip("/")
    m = re.search(r"\.([A-Za-z0-9]{1,8})$", path)
    return ("." + m.group(1).lower()) if m else ""


def ext_from_headers(headers, fallback_url=""):
    cd = headers.get("Content-Disposition", "")
    m = re.search(r"filename\*?=(?:UTF-8'')?[\"']?([^\"';]+)", cd, re.I)
    if m:
        e = ext_of_url(m.group(1))
        if e:
            return e
    ct = (headers.get("Content-Type") or "").lower()
    if "openxmlformats-officedocument.presentationml" in ct:
        return ".pptx"
    if "openxmlformats-officedocument.wordprocessingml" in ct:
        return ".docx"
    if "openxmlformats-officedocument.spreadsheetml" in ct:
        return ".xlsx"
    if "vnd.ms-powerpoint" in ct:
        return ".ppt"
    if "vnd.ms-word" in ct or "msword" in ct:
        return ".doc"
    if "vnd.ms-excel" in ct:
        return ".xls"
    if "zip" in ct:
        return ".zip"
    if "rar" in ct:
        return ".rar"
    if "pdf" in ct:
        return ".pdf"
    return ""


def resolve_file_url(data):
    p = data.get("path") or data.get("url") or data.get("file") or ""
    if not p:
        return None
    p = p.strip()
    if p.startswith("http://image1.bangongziyuan.com"):
        p = p[len("http://image1.bangongziyuan.com"):]
    if p.startswith("http"):
        return p
    return urljoin(CDN_BASE, p)


def add_download(session, user, item, timeout=30, referer=None):
    params = build_download_params(user, item["proId"],
                                   item.get("request_id", ""),
                                   item.get("ops_request_misc", ""))
    ref = referer or item.get("detail") or SITE
    return api_post(session, "/inter/addDownload", params, referer=ref, timeout=timeout)


def download_file(session, url, dest_dir, filename_no_ext="", referer=SITE,
                  timeout=60, retries=3):
    os.makedirs(dest_dir, exist_ok=True)
    attempt = 0
    last_err = None
    while attempt <= retries:
        try:
            headers = {"Referer": referer}
            with session.get(url, headers=headers, timeout=timeout, stream=True) as r:
                if r.status_code != 200:
                    last_err = f"HTTP {r.status_code}"
                    raise RuntimeError(last_err)
                ext = ext_of_url(url) or ext_from_headers(r.headers, url)
                ctype = r.headers.get("Content-Type", "")
                fname = f"{safe_name(filename_no_ext)}{ext}" if filename_no_ext else None
                if not fname:
                    # 尝试从 Content-Disposition 取服务器文件名
                    cd = r.headers.get("Content-Disposition", "")
                    m = re.search(r"filename\*?=(?:UTF-8'')?[\"']?([^\"';]+)", cd, re.I)
                    if m:
                        fname = safe_name(m.group(1))
                if not fname:
                    fname = safe_name(f"file_{int(time.time())}") + (ext or ".bin")
                dest = os.path.join(dest_dir, fname)
                tmp = dest + ".part"
                size = 0
                with open(tmp, "wb") as f:
                    for chunk in r.iter_content(chunk_size=1 << 16):
                        if chunk:
                            f.write(chunk)
                            size += len(chunk)
                if size == 0:
                    last_err = "空文件(0 字节)"
                    raise RuntimeError(last_err)
                os.replace(tmp, dest)
                return dest, size
        except Exception as e:
            attempt += 1
            last_err = str(e)
            print(f"      [重试 {attempt}/{retries}] 下载失败: {last_err}")
            time.sleep(2 * attempt)
    raise RuntimeError(last_err)


# --------------------------------------------------------------------------- #
#  状态管理(断点续传)                                                           #
# --------------------------------------------------------------------------- #
def load_json(path, default):
    if path and os.path.isfile(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def save_json(path, obj):
    directory = os.path.dirname(os.path.abspath(path))
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
                "w", encoding="utf-8", dir=directory,
                prefix=os.path.basename(path) + ".", suffix=".tmp", delete=False) as f:
            temp_path = f.name
            json.dump(obj, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, path)
        temp_path = None
    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


# --------------------------------------------------------------------------- #
#  流程                                                                        #
# --------------------------------------------------------------------------- #
def do_check(cfg, session):
    user, summary = fetch_user(session)
    print("=" * 60)
    if user is None:
        print("[check] 账号校验失败！")
        print("  ", summary)
        print("  请重新登录网站后，再导出最新 Cookie 填入 config.json。")
        return 1
    print("[check] 账号有效:")
    print("  ", summary)
    uid = user.get("id") or user.get("userId")
    member = user.get("member_type")
    if member in (None, "", "0"):
        print("  [!] 警告：当前账号未识别到 VIP/会员类型，服务端可能拒绝下载。")
    print("=" * 60)
    return 0


def build_sources(cfg):
    """把 config 中的分类URL + 关键词扩展成源列表。"""
    sources = []
    for s in cfg.get("sources", []) or []:
        if isinstance(s, str):
            sources.append({"url": s})
        else:
            sources.append(dict(s))
    for kw in cfg.get("keywords", []) or []:
        sources.append({
            "url": f"{SITE}/search/val-{quote(kw)}-type-0.html",
            "name": f"搜索:{kw}",
        })
    return sources


def gather(cfg, session):
    """抓取所有源 -> 过滤 -> 返回 items"""
    include = cfg.get("include", []) or []
    exclude = cfg.get("exclude", []) or []
    max_pages = int(cfg.get("max_pages", 0) or 0)
    timeout = int(cfg.get("timeout", 30))
    delay = cfg.get("delay", 1.0)
    try:
        delay = float(delay)
        delay = (delay * 0.6, delay * 1.4)
    except Exception:
        delay = (0.5, 1.2)

    all_items = []
    for src in build_sources(cfg):
        url = src["url"]
        name = src.get("name") or url
        inc = src.get("include") or include
        exc = src.get("exclude") or exclude
        mp = int(src.get("max_pages") or max_pages or 0)
        print(f"[目录] 源: {name}")
        print(f"       {url}")
        items, pages, stop = crawl_source(session, url, max_pages=mp, delay=delay, timeout=timeout)
        print(f"       -> 抓到 {len(items)} 条 (翻 {pages} 页)  {stop}")
        kept = 0
        for it in items:
            if matches_filters(it, inc, exc):
                it["source"] = name
                all_items.append(it)
                kept += 1
        print(f"       过滤后保留 {kept} 条")
        time.sleep(random.uniform(*delay))
    return all_items


def do_list(cfg, session):
    items = gather(cfg, session)
    cfg.setdefault("out_dir", "./downloads")
    queue_path = os.path.join(cfg["out_dir"], "queue.json")
    os.makedirs(cfg["out_dir"], exist_ok=True)
    save_json(queue_path, items)
    print("=" * 60)
    print(f"[list] 共 {len(items)} 条候选已写入 {queue_path}")
    for it in items[:20]:
        print(f"   {it['proId']:>7}  {it['title'][:60]}")
    if len(items) > 20:
        print(f"   ... 其余 {len(items) - 20} 条略")
    return 0


def do_download(cfg, session):
    out_dir = cfg.get("out_dir", "./downloads")
    queue_path = cfg.get("queue") or os.path.join(out_dir, "queue.json")
    state_path = cfg.get("state") or os.path.join(out_dir, "state.json")
    limit = int(cfg.get("limit", 0) or 0)
    delay = float(cfg.get("delay", 2.0))
    timeout = int(cfg.get("timeout", 60))
    retries = int(cfg.get("retries", 3))

    if not os.path.isfile(queue_path):
        print("[download] 未找到 queue.json —— 请先运行 list 生成候选清单，")
        print("           或运行 `all` 一步抓取并下载。")
        return 2

    items = load_json(queue_path, [])
    if not items:
        print("[download] queue.json 为空。")
        return 1

    user, summary = fetch_user(session)
    if user is None:
        print("[download] 账号校验失败：")
        print("   ", summary)
        print("   请重新登录网站后导出最新 Cookie。")
        return 1
    print("[download] 账号：", summary)

    state = load_json(state_path, {"ok": {}, "fail": {}})
    ok_map, fail_map = state["ok"], state["fail"]

    sub_by_source = {}
    for it in items:
        sub_by_source.setdefault(it.get("source", "misc"), []).append(it)

    done_cnt = skip_cnt = fail_cnt = 0
    print(f"[download] 共 {len(items)} 条，分 {len(sub_by_source)} 个目录")
    for src_name, sub in sub_by_source.items():
        folder = safe_name(src_name, 60)
        dest_dir = os.path.join(out_dir, folder)
        for it in sub:
            pid = it["proId"]
            if pid in ok_map:
                skip_cnt += 1
                continue
            # 若文件已存在也算完成
            print(f"  [{pid}] {it['title'][:55]}")
            try:
                j = add_download(session, user, it, timeout=timeout,
                                 referer=it.get("detail") or SITE)
            except Exception as e:
                print(f"     [失败] addDownload 请求异常: {e}")
                fail_map[pid] = {"reason": str(e), "title": it["title"]}
                fail_cnt += 1
                save_json(state_path, state)
                continue

            if j.get("status") == 1:
                url = resolve_file_url(j.get("data") or {})
                if not url:
                    print(f"     [失败] 接口无文件地址: {json.dumps(j, ensure_ascii=False)[:200]}")
                    fail_map[pid] = {"reason": "no file url", "title": it["title"]}
                    fail_cnt += 1
                    save_json(state_path, state)
                    continue
                print(f"     文件: {url}")
                try:
                    dest, size = download_file(
                        session, url, dest_dir,
                        filename_no_ext=f"{pid}_{it['title']}",
                        referer=it.get("detail") or SITE,
                        timeout=timeout, retries=retries)
                    ok_map[pid] = {"file": os.path.relpath(dest, out_dir),
                                   "size": size, "title": it["title"]}
                    done_cnt += 1
                    print(f"     [完成] {os.path.basename(dest)}  ({size/1024:.0f} KB)")
                except Exception as e:
                    print(f"     [失败] 文件下载失败: {e}")
                    fail_map[pid] = {"reason": str(e), "title": it["title"]}
                    fail_cnt += 1
                save_json(state_path, state)
            else:
                msg = str(j.get("msg") or j.get("data") or j)
                lim = ""
                d = j.get("data")
                if isinstance(d, dict):
                    lim = str(d.get("downloadLimit") or "")
                print(f"     [被拒] {msg} {lim}")
                fail_map[pid] = {"reason": msg, "title": it["title"]}
                fail_cnt += 1
                save_json(state_path, state)
                # 服务器明确限制/需要会员 -> 礼貌停止，避免被封
                joined = (msg + " " + lim)
                if any(k in joined for k in ("频繁", "明天", "次数", "恶意", "限额", "每日")):
                    print("[download] 服务端提示下载受限(可能为当日次数/频率限制)，")
                    print("           已停止本轮。请稍后(建议次日)重跑本命令续传。")
                    break
                if any(k in joined for k in ("会员", "VIP", "权限", "登录", "过期")):
                    print("[download] 服务端提示权限不足(会员/登录/过期)。请检查账号后重试。")
                    break

            if limit and done_cnt >= limit:
                print(f"[download] 已达 limit={limit}，停止。")
                break
            time.sleep(max(0.5, delay + random.uniform(0, delay * 0.5)))
        else:
            continue
        break  # 受限/权限不足时跳出外层

    print("=" * 60)
    print(f"[download] 完成 {done_cnt}，跳过(已完成) {skip_cnt}，失败 {fail_cnt}")
    if fail_map:
        print("失败明细已记录在 state.json 的 fail 段，重跑即可续传失败项。")
    return 0


def do_all(cfg, session):
    items = gather(cfg, session)
    out_dir = cfg.get("out_dir", "./downloads")
    os.makedirs(out_dir, exist_ok=True)
    queue_path = os.path.join(out_dir, "queue.json")
    save_json(queue_path, items)
    cfg["queue"] = queue_path
    print(f"[all] 候选 {len(items)} 条 -> 开始下载")
    return do_download(cfg, session)


def main():
    ap = argparse.ArgumentParser(description="办公资源网批量下载器", add_help=True)
    ap.add_argument("mode", choices=["check", "list", "download", "all"],
                    help="check=校验账号; list=只抓目录生成 queue.json; "
                         "download=按 queue.json 下载(可续传); all=抓取+下载")
    ap.add_argument("-c", "--config", default="config.json", help="配置文件(默认 config.json)")
    ap.add_argument("--cookie-string", default="", help="浏览器 Cookie 头字符串")
    ap.add_argument("--cookie-file", default="", help="Netscape 格式 cookies.txt 文件")
    ap.add_argument("--out", default="", help="输出目录(覆盖配置)")
    ap.add_argument("--limit", type=int, default=0, help="最多下载 N 条(0=不限)")
    ap.add_argument("--max-pages", type=int, default=0, help="每个源最多翻 N 页(0=全部)")
    ap.add_argument("--delay", type=float, default=0, help="每次下载间隔秒数(默认读配置)")
    ap.add_argument("--include", default="", help="标题必须包含的关键词，逗号分隔(覆盖配置)")
    ap.add_argument("--exclude", default="", help="标题必须排除的关键词，逗号分隔(覆盖配置)")
    args = ap.parse_args()

    cfg = load_json(args.config, {}) if os.path.isfile(args.config) else {}
    if not cfg:
        print(f"[i] 未找到配置文件 {args.config}，使用命令行参数(可能缺少 cookie/目录)。")

    # 命令行覆盖配置
    if args.cookie_string:
        cfg["cookie_string"] = args.cookie_string
    if args.cookie_file:
        cfg["cookie_file"] = args.cookie_file
    if args.out:
        cfg["out_dir"] = args.out
    if args.limit:
        cfg["limit"] = args.limit
    if args.max_pages:
        cfg["max_pages"] = args.max_pages
    if args.delay:
        cfg["delay"] = args.delay
    if args.include:
        cfg["include"] = [k.strip() for k in args.include.split(",") if k.strip()]
    if args.exclude:
        cfg["exclude"] = [k.strip() for k in args.exclude.split(",") if k.strip()]

    cfg.setdefault("out_dir", "./downloads")
    if "cookies" in cfg:                     # 兼容键名 cookies
        cfg.setdefault("cookie_string", cfg.pop("cookies"))

    session = make_session(cfg.get("cookie_string", ""), cfg.get("cookie_file", ""))

    if args.mode == "check":
        return do_check(cfg, session)
    if args.mode == "list":
        return do_list(cfg, session)
    if args.mode == "download":
        return do_download(cfg, session)
    if args.mode == "all":
        return do_all(cfg, session)
    return 0


if __name__ == "__main__":
    sys.exit(main())
