# -*- coding: utf-8 -*-
"""Verify resume: second run in SAME out_dir skips already-downloaded items."""
import json, os, sys, tempfile

sys.path.insert(0, os.path.dirname(__file__))
import bzy_downloader as b

class FakeSession:
    pass

orig_fetch, orig_add, orig_dl = b.fetch_user, b.add_download, b.download_file
b.fetch_user = lambda s: ({"id": 1}, "fake")
b.add_download = lambda s, u, it, **k: {"status": 1, "data": {"path": "/f/x.zip"}}
def fake_dl(s, url, ddir, filename_no_ext="", **k):
    os.makedirs(ddir, exist_ok=True)
    p = os.path.join(ddir, filename_no_ext + ".zip")
    open(p, "wb").write(b"PK")
    return p, 2
b.download_file = fake_dl

tmp = tempfile.mkdtemp(prefix="bzy_resume_")
items = [{"proId": str(i), "title": f"T{i}", "detail": f"https://x/detail-{i}.html", "source": "s"} for i in (1, 2, 3)]
with open(os.path.join(tmp, "queue.json"), "w", encoding="utf-8") as f:
    json.dump(items, f, ensure_ascii=False)

cfg = {"out_dir": tmp, "delay": 0, "timeout": 5, "retries": 1, "limit": 0}
try:
    rc1 = b.do_download(cfg, FakeSession())   # downloads all 3
    files1 = sorted(os.listdir(os.path.join(tmp, "s")))
    rc2 = b.do_download(cfg, FakeSession())   # should skip all 3
    files2 = sorted(os.listdir(os.path.join(tmp, "s")))
    print("run1 files:", files1)
    print("run2 files:", files2)
    assert rc1 == rc2 == 0
    assert len(files1) == 3 and files1 == files2, "resume must not re-download or duplicate"
    print("RESUME OK (same 3 files, no duplicates)")
finally:
    b.fetch_user, b.add_download, b.download_file = orig_fetch, orig_add, orig_dl
