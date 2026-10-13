# -*- coding: utf-8 -*-
"""Integration test for do_download via monkeypatched network fns."""
import json, os, sys, tempfile

sys.path.insert(0, os.path.dirname(__file__))
import bzy_downloader as b

fake_user = {"id": 1, "userBytype": "vip", "userByDownloadLimit": 100}

class FakeSession:
    pass

def run(items, mode="ok"):
    tmp = tempfile.mkdtemp(prefix="bzy_it_")
    cfg = {"out_dir": tmp, "delay": 0.01, "timeout": 5, "retries": 1, "limit": 0}
    with open(os.path.join(tmp, "queue.json"), "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False)

    orig_fetch = b.fetch_user
    orig_add = b.add_download
    orig_dl = b.download_file

    def fake_fetch(session):
        return dict(fake_user), "fake user ok"

    def fake_add(session, user, item, timeout=30, referer=None):
        if mode == "reject":
            return {"status": 0, "msg": "今日下载次数已用完，请明天再来~", "data": {"downloadLimit": "0"}}
        return {"status": 1, "data": {"path": "/files/x/123.zip"}}

    def fake_dl(session, url, dest_dir, filename_no_ext="", referer=b.SITE, timeout=60, retries=3):
        os.makedirs(dest_dir, exist_ok=True)
        p = os.path.join(dest_dir, safe := (filename_no_ext or "x") + ".zip")
        with open(p, "wb") as f:
            f.write(b"PK")
        return p, 2

    b.fetch_user = fake_fetch
    b.add_download = fake_add
    b.download_file = fake_dl
    try:
        rc = b.do_download(cfg, FakeSession())
        st = json.load(open(os.path.join(tmp, "state.json"), encoding="utf-8"))
        return rc, st, tmp
    finally:
        b.fetch_user = orig_fetch
        b.add_download = orig_add
        b.download_file = orig_dl

items = [
    {"proId": "111", "title": "述职模板A", "detail": "https://x/detail-111.html", "source": "s1"},
    {"proId": "222", "title": "汇报模板B", "detail": "https://x/detail-222.html", "source": "s1"},
]

rc, st, tmp = run(items, mode="ok")
print("== ok mode rc:", rc, "== ok:", sorted(st["ok"]), "== fail:", st["fail"])
assert rc == 0 and set(st["ok"]) == {"111", "222"} and not st["fail"], "ok-mode failed"

# resume: second run should skip existing
rc, st2, tmp2 = run(items, mode="ok")
print("== resume ok:", sorted(st2["ok"]))
assert len(st2["ok"]) == 2

rc, st, tmp = run(items, mode="reject")
print("== reject mode rc:", rc, "== ok:", st["ok"], "== fail keys:", list(st["fail"]))
assert list(st["fail"]).pop(0) == "111", "reject-mode should record and stop"

print("INTEGRATION OK")
