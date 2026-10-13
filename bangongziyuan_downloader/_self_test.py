# -*- coding: utf-8 -*-
"""Quick sanity checks of pure functions (no network)."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
import bzy_downloader as b

fails = []

def eq(name, got, want):
    if got != want:
        fails.append(f"{name}: got={got!r} want={want!r}")
    else:
        print("ok:", name)

# page_url
eq("page_url p1", b.page_url("https://x.com/ppt/cat1-11-stu-0.html", 1),
   "https://x.com/ppt/cat1-11-stu-0.html")
eq("page_url p2", b.page_url("https://x.com/ppt/cat1-11-stu-0.html", 2),
   "https://x.com/ppt/cat1-11-stu-0-pn-2.html")
eq("page_url search p3", b.page_url("https://x.com/search/val-a-type-0.html", 3),
   "https://x.com/search/val-a-type-0-pn-3.html")
eq("page_url no .html", b.page_url("https://x.com/ppt/cat1-11-stu-0", 2),
   "https://x.com/ppt/cat1-11-stu-0-pn-2.html")

# resolve_file_url (mirror of site JS: downUrl + path minus http://image1 prefix)
eq("resolve rel", b.resolve_file_url({"path": "/files/p/1.zip"}),
   "https://image1.bangongziyuan.com/files/p/1.zip")
eq("resolve http abs", b.resolve_file_url({"path": "http://image1.bangongziyuan.com/files/p/1.zip"}),
   "https://image1.bangongziyuan.com/files/p/1.zip")
eq("resolve https abs", b.resolve_file_url({"path": "https://other.cdn/x.zip"}),
   "https://other.cdn/x.zip")
eq("resolve empty", b.resolve_file_url({"path": ""}), None)
eq("resolve alt key", b.resolve_file_url({"url": "/a/b.ppt"}), "https://image1.bangongziyuan.com/a/b.ppt")

# build_download_params
user = {"id": 123, "userBytype": "vip", "userByDownloadLimit": 50}
p = b.build_download_params(user, 999, "rid1", "ops1")
eq("params userId", p["userId"], "123")
eq("params proId", p["proId"], "999")
eq("params inviter", p["inviterLevel"], 0)
eq("params rid", p["request_id"], "rid1")
eq("params ops", p["ops_request_misc"], "ops1")
p2 = b.build_download_params(user, 1)
eq("params defaults", (p2["request_id"], p2["ops_request_misc"], p2["userByExpireDate"]), ("", "", ""))

# parse_items on a synthetic snippet
html = ('<div class="model-item"><a class="title" data-id="12" href="/ppt/detail-12.html">  '
        '述职 <b>报告</b> 模板</a></div>'
        '<div class="model-item"><a class="title"  data-id="34" href="/ppt/detail-34.html"> 汇报 </a></div>')
items = b.parse_items(html, "http://page")
eq("parse count", len(items), 2)
eq("parse id", items[0]["proId"], "12")
eq("parse title", items[0]["title"], "述职 报告 模板")
eq("parse id2", items[1]["proId"], "34")

# filters
it = {"title": "蓝色述职汇报PPT模板"}
eq("filter include", b.matches_filters(it, ["述职"], []), True)
eq("filter include miss", b.matches_filters(it, ["卡通"], []), False)
eq("filter exclude", b.matches_filters(it, [], ["汇报"]), False)
eq("filter both", b.matches_filters(it, ["述职"], ["卡通"]), True)

# safe_name
eq("safe win chars", b.safe_name('a/b\\c:d*e?"f<>g|h'), "a_b_c_d_e__f__g_h")
eq("safe trim", b.safe_name("  hello  ."), "hello")

# ext helpers
eq("ext url", b.ext_of_url("https://x/a.B.PPT?k=1"), ".ppt")
eq("ext cd", b.ext_from_headers({"Content-Disposition": 'attachment; filename="t.zip"'}, ""), ".zip")
eq("ext ct pptx", b.ext_from_headers({"Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation"}), ".pptx")

print()
if fails:
    print("FAILURES:")
    for f in fails:
        print("  -", f)
    sys.exit(1)
print("ALL OK")
