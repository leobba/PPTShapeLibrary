/* ============================================================
   app.js — 应用层：状态树、渲染、交互
   ============================================================ */
(function () {
    "use strict";

    // ---------------- 状态树 ----------------
    var state = {
        settings: {},
        categories: [],
        filter: "",
        category: "全部",
        sort: "最新",
        items: [],            // 已渲染条目
        filteredTotal: 0,
        total: 0,
        selected: {},         // id -> true（跨分页保留）
        selectedOrder: [],
        loading: false,
        adjustVisible: false,
        lastInsertedId: null
    };

    // ---------------- 工具 ----------------
    function $(sel) { return document.querySelector(sel); }
    function el(tag, cls, text) {
        var d = document.createElement(tag);
        if (cls) d.className = cls;
        if (text !== undefined) d.textContent = text;
        return d;
    }
    function ico(code) { return el("span", "ico", code); }

    function toast(msg, isError) {
        var wrap = $(".toast-wrap");
        if (!wrap) {
            wrap = el("div", "toast-wrap");
            document.body.appendChild(wrap);
        }
        var t = el("div", "toast" + (isError ? " error" : ""), msg);
        wrap.appendChild(t);
        setTimeout(function () { t.remove(); }, 2500);
    }

    // ---------------- 初始化 ----------------
    document.addEventListener("DOMContentLoaded", function () {
        buildSkeleton();
        bindEvents();
        Bridge.init().then(function (data) {
            state.settings = data;
            state.categoryCounts = data.categoryCounts || {};
            applyTheme(data.theme || "light");
            applyCardSize(data.cardSize || "compact");
            renderCategories(data.categories || []);
            reload();
        }).catch(function (err) {
            toast("初始化失败：" + err.message, true);
        });
    });

    // 订阅 C# 下行推送
    Bridge.on("refresh", function () { reload(); });
    Bridge.on("settingsChanged", function (s) {
        state.settings = s;
        applyTheme(s.theme || "light");
        applyCardSize(s.cardSize || "compact");
        reload();
    });

    // ---------------- 骨架 ----------------
    function buildSkeleton() {
        var app = $("#app");
        app.innerHTML = "";

        // ① 头部
        var header = el("div", "header");
        var hrow = el("div", "header-row");
        var sw = el("div", "search-wrap");
        var sico = ico("\uE721");
        sico.classList.add("ico-search");
        sw.appendChild(sico);
        var search = el("input");
        search.id = "searchInput";
        search.placeholder = "\uE721  搜索 名称 / 分类 / 标签";
        search.setAttribute("autocomplete", "off");
        sw.appendChild(search);
        hrow.appendChild(sw);

        var sort = el("select", "sort-select");
        ["最新", "最旧", "名称", "最近使用", "最常用", "自定义"].forEach(function (m) {
            var o = el("option", "", m); o.value = m; sort.appendChild(o);
        });
        sort.id = "sortSelect";
        hrow.appendChild(sort);
        header.appendChild(hrow);
        app.appendChild(header);

        // ② 分类条
        var catBar = el("div", "category-bar");
        catBar.id = "categoryBar";
        app.appendChild(catBar);

        // ③ 状态行
        var status = el("div", "status-row");
        var cnt = el("span"); cnt.id = "lblCount";
        status.appendChild(cnt);
        var themeBtn = el("button", "theme-toggle");
        themeBtn.id = "themeToggle";
        themeBtn.innerHTML = '<span class="ico">\uE793</span>';
        themeBtn.title = "切换深浅主题";
        status.appendChild(themeBtn);
        app.appendChild(status);

        // ④ 网格滚动区
        var scroll = el("div", "grid-scroll");
        scroll.id = "gridScroll";
        var grid = el("div", "card-grid");
        grid.id = "cardGrid";
        scroll.appendChild(grid);
        var sentinel = el("div"); sentinel.id = "sentinel";
        scroll.appendChild(sentinel);
        app.appendChild(scroll);

        // 空状态
        var empty = el("div", "empty-state");
        empty.id = "emptyState";
        empty.style.display = "none";
        empty.innerHTML = '<span class="ico">\uE8A5</span>';
        app.appendChild(empty);

        // ⑤ 快速调整条
        var adjust = el("div", "float-bar");
        adjust.id = "adjustBar";
        adjust.hidden = true;
        [
            ["\uE7E8", "rotateL", "左转15°"], ["\uE7E9", "rotateR", "右转15°"],
            ["透明-", "opacityDown", ""], ["透明+", "opacityUp", ""],
            ["←", "nudgeL", ""], ["→", "nudgeR", ""], ["↑", "nudgeU", ""], ["↓", "nudgeD", ""]
        ].forEach(function (cfg) {
            var b = el("button", "bar-btn", cfg[0]);
            b.title = cfg[2] || cfg[0];
            b.addEventListener("click", function () {
                Bridge.adjust(cfg[1]).catch(function (e) { toast(e.message, true); });
            });
            adjust.appendChild(b);
        });
        var done = el("button", "bar-btn primary", "完成");
        done.style.marginLeft = "auto";
        done.addEventListener("click", function () { adjust.hidden = true; });
        adjust.appendChild(done);
        app.appendChild(adjust);

        // 批量操作条
        var batch = el("div", "float-bar");
        batch.id = "batchBar";
        batch.hidden = true;
        var bt = el("div", "bar-text"); bt.id = "batchText";
        batch.appendChild(bt);
        var bCancel = el("button", "bar-btn", "取消");
        bCancel.addEventListener("click", clearSelection);
        batch.appendChild(bCancel);
        var bMove = el("button", "bar-btn", "移动分类");
        bMove.addEventListener("click", function () { batchMovePrompt(); });
        batch.appendChild(bMove);
        var bDel = el("button", "bar-btn danger", "删除");
        bDel.addEventListener("click", batchDelete);
        batch.appendChild(bDel);
        app.appendChild(batch);

        // Toast 容器
        app.appendChild(el("div", "toast-wrap"));
    }

    // ---------------- 事件绑定 ----------------
    function bindEvents() {
        $("#searchInput").addEventListener("input", debounce(function () {
            state.filter = this.value.trim();
            reload();
        }.bind($("#searchInput")), 150));

        $("#searchInput").addEventListener("keydown", function (e) {
            if (e.key === "Escape") { this.value = ""; state.filter = ""; reload(); }
        });

        $("#sortSelect").addEventListener("change", function () {
            state.sort = this.value;
            reload();
        });

        $("#themeToggle").addEventListener("click", function () {
            var next = (state.settings.theme === "dark") ? "light" : "dark";
            Bridge.saveSettings({ theme: next }).then(function (s) {
                state.settings = s;
                applyTheme(s.theme);
            }).catch(function (e) { toast(e.message, true); });
        });

        // 滚动分页
        $("#gridScroll").addEventListener("scroll", debounce(function () {
            var sc = this;
            if (sc.scrollTop + sc.clientHeight >= sc.scrollHeight - 240) {
                loadMore();
            }
        }.bind($("#gridScroll")), 80));
    }

    function debounce(fn, ms) {
        var t;
        return function () {
            clearTimeout(t);
            t = setTimeout(fn, ms);
        };
    }

    // ---------------- 数据加载 ----------------
    function reload() {
        state.items = [];
        state.loadedOffset = 0;
        renderGrid();       // 清空并显示骨架
        loadPage(0);
    }

    function loadMore() {
        if (state.loading) return;
        if (state.items.length >= state.filteredTotal && state.filteredTotal > 0) return;
        loadPage(state.items.length);
    }

    function loadPage(offset) {
        if (state.loading) return;
        state.loading = true;
        if (offset === 0) showSkeletons();
        Bridge.list(state.filter, state.category, state.sort, offset, getPageSize())
            .then(function (data) {
                state.loading = false;
                state.filteredTotal = data.filteredTotal;
                state.total = data.total;
                if (offset === 0) state.items = [];
                state.items = state.items.concat(data.items);
                renderGrid();
                updateStatus();
            })
            .catch(function (err) {
                state.loading = false;
                toast(err.message, true);
            });
    }

    function showSkeletons() {
        var grid = $("#cardGrid");
        if (!grid) return;
        grid.innerHTML = "";
        var empty = $("#emptyState");
        if (empty) empty.style.display = "none";
        for (var i = 0; i < 8; i++) {
            var sk = el("article", "card skeleton");
            var wrap = el("div");
            var skImg = el("div", "sk sk-img");
            var skLine = el("div", "sk sk-line");
            wrap.appendChild(skImg);
            wrap.appendChild(skLine);
            sk.appendChild(wrap);
            grid.appendChild(sk);
        }
    }

    function getPageSize() {
        var n = parseInt(state.settings.pageSize, 10);
        return isNaN(n) || n <= 0 ? 200 : n;
    }

    // ---------------- 渲染 ----------------
    function renderCategories(cats) {
        var bar = $("#categoryBar");
        if (!bar) return;
        bar.innerHTML = "";
        var counts = state.categoryCounts || {};
        var names = ["全部"].concat((cats || []).filter(function (c) { return c !== "全部"; }));
        names.forEach(function (c) {
            var label = c;
            if (c !== "全部" && counts[c]) label += " " + counts[c];
            var chip = el("button", "chip" + (c === state.category ? " active" : ""), label);
            chip.addEventListener("click", function () {
                state.category = c;
                bar.querySelectorAll(".chip").forEach(function (x) { x.classList.remove("active"); });
                chip.classList.add("active");
                reload();
            });
            bar.appendChild(chip);
        });
    }

    function updateStatus() {
        var cnt = $("#lblCount");
        if (!cnt) return;
        if (state.filter || state.category !== "全部") {
            cnt.textContent = "筛选 " + state.items.length + " / " + state.filteredTotal + " · 共 " + state.total;
        } else {
            cnt.textContent = "共 " + state.total + " 项";
        }
    }

    function renderGrid() {
        var grid = $("#cardGrid");
        if (!grid) return;
        grid.innerHTML = "";

        // 移除旧的"最近使用"区
        var oldRecent = $("#recentSection");
        if (oldRecent) oldRecent.remove();

        // 无筛选时：顶部展示"最近使用"横向条；已在横条中的条目不再进主网格，避免重复
        var recentExcluded = {};
        if (!state.filter && state.category === "全部" && state.sort !== "最近使用") {
            var recent = state.items.filter(function (m) { return m.LastUsedAt || m.lastUsedAt; })
                .sort(function (a, b) {
                    return (b.lastUsedAt || "").localeCompare(a.lastUsedAt || "");
                }).slice(0, 6);
            if (recent.length >= 2) {
                var sec = el("div");
                sec.id = "recentSection";
                var title = el("div", "section-title");
                title.appendChild(ico("\uE818"));
                title.appendChild(el("span", "", "最近使用"));
                var more = el("button", "more", "查看全部 ›");
                more.addEventListener("click", function () {
                    $("#sortSelect").value = "最近使用";
                    state.sort = "最近使用";
                    reload();
                });
                title.appendChild(more);
                sec.appendChild(title);

                var row = el("div", "card-grid");
                row.style.gridTemplateColumns = "repeat(auto-fill, minmax(84px, 1fr))";
                recent.forEach(function (m) {
                    recentExcluded[m.id] = true;
                    row.appendChild(createCard(m));
                });
                sec.appendChild(row);
                grid.parentElement.insertBefore(sec, grid);

                // 最近区与主网格之间加分隔
                var gap = el("div");
                gap.style.height = "14px";
                sec.appendChild(gap);
            }
        }

        if (state.items.length === 0) {
            var empty = $("#emptyState");
            empty.style.display = "";
            empty.innerHTML = '<span class="ico">\uE8A5</span>' +
                ((state.filter || state.category !== "全部")
                    ? "没有符合筛选条件的收藏"
                    : "还没有收藏<br><br>选中页面上的形状<br>点击工具栏「收藏形状」或「收藏样式」");
            updateStatus();
            return;
        }
        var empty = $("#emptyState");
        if (empty) empty.style.display = "none";

        state.items.forEach(function (meta) {
            if (recentExcluded[meta.id]) return;   // 已在"最近使用"横条展示
            grid.appendChild(createCard(meta));
        });
        updateSelectionUI();

        // 分页哨兵
        var sentinel = $("#sentinel");
        if (sentinel && state.items.length < state.filteredTotal) {
            observeSentinel(sentinel);
        }
    }

    var sentinelObserver = null;
    function observeSentinel(sentinel) {
        if (!("IntersectionObserver" in window)) return;
        if (sentinelObserver) sentinelObserver.disconnect();
        sentinelObserver = new IntersectionObserver(function (entries) {
            entries.forEach(function (en) {
                if (en.isIntersecting) loadMore();
            });
        }, { rootMargin: "240px" });
        sentinelObserver.observe(sentinel);
    }

    function createCard(meta) {
        var card = el("article", "card");
        card.dataset.id = meta.id;

        // 缩略图撑满整卡
        if (meta.previewUrl) {
            var img = el("img", "thumb");
            img.loading = "lazy";
            img.src = meta.previewUrl;
            img.onerror = function () {
                var fb = el("div", "thumb-fallback");
                fb.innerHTML = '<span class="ico">\uE8A5</span>';
                img.replaceWith(fb);
            };
            card.appendChild(img);
        } else {
            var fb2 = el("div", "thumb-fallback");
            fb2.innerHTML = '<span class="ico">\uE8A5</span>';
            card.appendChild(fb2);
        }

        // 类型色点（常显）
        card.appendChild(el("span", "badge " + (meta.kind === "Style" ? "style" : "shape")));

        // hover 底部遮罩：名称 + 类型
        var overlay = el("div", "card-overlay");
        var nameSpan = el("span", "", meta.name);
        overlay.appendChild(nameSpan);
        overlay.appendChild(el("span", "badge-text-in-overlay " + (meta.kind === "Style" ? "style" : "shape"),
            meta.kind === "Style" ? "样式" : "形状"));
        card.appendChild(overlay);

        // 删除按钮
        var del = el("button", "del-btn", "\uE711");
        del.title = "删除";
        del.addEventListener("click", function (e) {
            e.stopPropagation();
            confirmModal("删除收藏「" + meta.name + "」？", function () {
                Bridge.del([meta.id]).then(function () {
                    removeItemLocal(meta.id);
                    toast("已删除");
                }).catch(function (e2) { toast(e2.message, true); });
            });
        });
        card.appendChild(del);

        // tooltip
        var tip = meta.name + "  [分类: " + meta.category + "]";
        if (meta.tags && meta.tags.length) tip += "\n标签: " + meta.tags.join("、");
        if (meta.useCount > 0) tip += "\n使用 " + meta.useCount + " 次";
        card.title = tip;

        // 选中态
        if (state.selected[meta.id]) card.classList.add("selected");

        // 点击 / 多选 / 双击
        card.addEventListener("click", function (e) {
            if (e.ctrlKey) {
                toggleSelect(meta.id, card);
                return;
            }
            if (Object.keys(state.selected).length > 0) {
                clearSelection();
                return;
            }
            if (state.settings.clickToInsert === false) {
                showPreview(meta);
            } else {
                doInsert(meta, card);
            }
        });

        card.addEventListener("dblclick", function () {
            if (state.settings.clickToInsert === false) doInsert(meta, card);
        });

        // 右键菜单
        card.addEventListener("contextmenu", function (e) {
            e.preventDefault();
            showContextMenu(e.clientX, e.clientY, meta);
        });

        // 拖拽排序（自定义模式）
        setupDrag(card, meta);

        return card;
    }

    /** 插入成功的中央对勾动画 */
    function showInsertCheck(card) {
        if (!card) return;
        var wrap = el("div", "insert-check");
        var c = el("span", "ico", "\uE73E");
        wrap.appendChild(c);
        card.appendChild(wrap);
        setTimeout(function () { wrap.remove(); }, 600);
    }

    // ---------------- 插入 / 应用 ----------------
    var lastInsert = { id: null, t: 0 };
    function doInsert(meta, card) {
        // 双击防重：同一卡片 600ms 内只插一次
        var now = Date.now();
        if (meta.id === lastInsert.id && now - lastInsert.t < 600) return;
        lastInsert = { id: meta.id, t: now };

        Bridge.insert(meta.id).then(function () {
            flashCard(card);
            showInsertCheck(card);
            showAdjustBar();
        }).catch(function (e) { toast(e.message, true); });
    }

    // TouchUsed 由 C# insert 处理器内完成；此处仅为兼容保留
    function flashCard(card) {
        if (!card) return;
        card.classList.remove("flash");
        void card.offsetWidth;
        card.classList.add("flash");
    }

    function showAdjustBar() {
        var bar = $("#adjustBar");
        if (!bar) return;
        bar.hidden = false;
        if (state._adjustTimer) clearTimeout(state._adjustTimer);
        state._adjustTimer = setTimeout(function () { bar.hidden = true; }, 10000);
    }

    // ---------------- 多选 ----------------
    function toggleSelect(id, card) {
        if (state.selected[id]) {
            delete state.selected[id];
            card.classList.remove("selected");
            state.selectedOrder = state.selectedOrder.filter(function (x) { return x !== id; });
        } else {
            state.selected[id] = true;
            state.selectedOrder.push(id);
            card.classList.add("selected");
        }
        updateBatchUI();
    }

    function clearSelection() {
        state.selected = {};
        state.selectedOrder = [];
        document.querySelectorAll(".card.selected").forEach(function (c) { c.classList.remove("selected"); });
        updateBatchUI();
    }

    function updateBatchUI() {
        var n = Object.keys(state.selected).length;
        var bar = $("#batchBar");
        if (!bar) return;
        bar.hidden = n === 0;
        var txt = $("#batchText");
        if (txt) txt.textContent = "已选 " + n + " 项";
    }

    function batchDelete() {
        var ids = Object.keys(state.selected);
        if (ids.length === 0) return;
        confirmModal("确定删除选中的 " + ids.length + " 项收藏？", function () {
            Bridge.del(ids).then(function () {
                ids.forEach(removeItemLocal);
                clearSelection();
                toast("已删除 " + ids.length + " 项");
            }).catch(function (e) { toast(e.message, true); });
        });
    }

    function batchMovePrompt() {
        var ids = Object.keys(state.selected);
        if (ids.length === 0) return;
        inputModal("批量移动分类", "目标分类名称：", "", function (cat) {
            if (!cat) return;
            Bridge.batchMove(ids, cat).then(function () {
                ids.forEach(removeItemLocal);
                clearSelection();
                toast("已移动 " + ids.length + " 项到「" + cat + "」");
            }).catch(function (e) { toast(e.message, true); });
        });
    }

    /** 本地移除条目（不整页刷新，保持滚动位置） */
    function removeItemLocal(id) {
        state.items = state.items.filter(function (m) { return m.id !== id; });
        delete state.selected[id];
        state.filteredTotal = Math.max(0, state.filteredTotal - 1);
        state.total = Math.max(0, state.total - 1);
        var card = document.querySelector('.card[data-id="' + id + '"]');
        if (card) card.remove();
        updateStatus();
        checkEmpty();
    }

    function checkEmpty() {
        if (state.items.length === 0) reload();
    }

    // ---------------- 右键菜单 ----------------
    function showContextMenu(x, y, meta) {
        closeContextMenu();
        var menu = el("div", "ctx-menu");
        menu.id = "ctxMenu";

        function mi(text, fn) {
            var b = el("button", "mi", text);
            b.addEventListener("click", function () { closeContextMenu(); fn(); });
            menu.appendChild(b);
        }
        function sep() { menu.appendChild(el("div", "sep")); }

        mi("插入到当前页", function () { doInsert(meta, findCard(meta.id)); });
        mi("应用样式到选中形状", function () {
            Bridge.applyStyle(meta.id).then(function () {
                Bridge.touch(meta.id); toast("样式已应用");
            }).catch(function (e) { toast(e.message, true); });
        });
        mi("应用到本页同类形状", function () {
            Bridge.applyStyleSameType(meta.id).then(function () {
                Bridge.touch(meta.id);
            }).catch(function (e) { toast(e.message, true); });
        });
        mi("预览大图…", function () { showPreview(meta); });
        sep();
        mi("移动到分类…", function () {
            inputModal("移动到分类", "分类名称（输入新分类或已有分类）：", meta.category, function (v) {
                if (!v) return;
                Bridge.setCategory(meta.id, v).then(function () { reload(); }).catch(function (e) { toast(e.message, true); });
            });
        });
        mi("编辑标签…", function () {
            inputModal("编辑标签", "标签（多个用逗号分隔）：", (meta.tags || []).join(","), function (v) {
                Bridge.setTags(meta.id, v).then(function () { reload(); }).catch(function (e) { toast(e.message, true); });
            });
        });
        mi("重命名", function () {
            inputModal("重命名收藏", "名称：", meta.name, function (v) {
                if (!v) return;
                Bridge.rename(meta.id, v).then(function () { reload(); }).catch(function (e) { toast(e.message, true); });
            });
        });
        mi("删除", function () {
            confirmModal("确定删除收藏「" + meta.name + "」？", function () {
                Bridge.del([meta.id]).then(function () {
                    removeItemLocal(meta.id); toast("已删除");
                }).catch(function (e) { toast(e.message, true); });
            });
        });
        sep();
        mi("打开收藏文件夹", function () { Bridge.openFolder(meta.id); });

        document.body.appendChild(menu);

        // 视口边缘翻转
        var mw = menu.offsetWidth, mh = menu.offsetHeight;
        if (x + mw > window.innerWidth - 4) x = window.innerWidth - mw - 4;
        if (y + mh > window.innerHeight - 4) y = window.innerHeight - mh - 4;
        menu.style.left = x + "px";
        menu.style.top = y + "px";

        setTimeout(function () {
            document.addEventListener("click", closeContextMenu, { once: true });
            document.addEventListener("contextmenu", closeContextMenu, { once: true });
            window.addEventListener("scroll", closeContextMenu, { once: true, capture: true });
        }, 10);
    }

    function closeContextMenu() {
        var m = $("#ctxMenu");
        if (m) m.remove();
    }

    function findCard(id) {
        return document.querySelector('.card[data-id="' + id + '"]');
    }

    // ---------------- 大图预览 ----------------
    function showPreview(meta) {
        closeLightbox();
        var mask = el("div", "lightbox-mask");
        mask.id = "lightbox";
        var lb = el("div", "lightbox");
        var wrap = el("div", "lb-img-wrap");
        var img = el("img");
        img.src = meta.previewUrl || "";
        wrap.appendChild(img);
        lb.appendChild(wrap);
        var info = el("div", "lb-info");
        info.appendChild(el("div", "lb-name", meta.name));
        info.appendChild(el("div", "lb-kind", meta.kind === "Style" ? "样式收藏 · 右键卡片可应用" : "形状收藏"));
        lb.appendChild(info);
        var actions = el("div", "lb-actions");
        var insBtn = el("button", "bar-btn primary", "插入到当前页");
        insBtn.addEventListener("click", function () {
            closeLightbox();
            doInsert(meta, findCard(meta.id));
        });
        actions.appendChild(insBtn);
        var closeBtn = el("button", "bar-btn", "关闭");
        closeBtn.addEventListener("click", closeLightbox);
        actions.appendChild(closeBtn);
        lb.appendChild(actions);
        mask.appendChild(lb);
        mask.addEventListener("click", function (e) { if (e.target === mask) closeLightbox(); });
        document.body.appendChild(mask);
        document.addEventListener("keydown", lightboxEsc);
    }
    function lightboxEsc(e) { if (e.key === "Escape") closeLightbox(); }
    function closeLightbox() {
        var m = $("#lightbox");
        if (m) m.remove();
        document.removeEventListener("keydown", lightboxEsc);
    }

    // ---------------- 模态框 ----------------
    function inputModal(title, label, value, onOk) {
        openModal(title, function (body, ok) {
            var input = el("input", "modal-input");
            input.value = value || "";
            body.appendChild(input);
            setTimeout(function () { input.focus(); input.select(); }, 30);
            return function () {
                if (input.value.trim() !== "") onOk(input.value.trim());
            };
        });
    }

    function confirmModal(message, onOk) {
        openModal("确认操作", function (body, ok) {
            body.appendChild(el("div", "modal-msg", message));
            return function () { onOk(); };
        });
    }

    function openModal(title, buildBody) {
        closeModal();
        var mask = el("div", "modal-mask");
        mask.id = "modal";
        var modal = el("div", "modal");
        modal.appendChild(el("div", "modal-title", title));
        var body = el("div");
        modal.appendChild(body);
        var actions = el("div", "modal-actions");
        var ok = el("button", "bar-btn primary", "确定");
        var cancel = el("button", "bar-btn", "取消");
        actions.appendChild(ok);
        actions.appendChild(cancel);
        modal.appendChild(actions);
        mask.appendChild(modal);
        document.body.appendChild(mask);

        var commit = buildBody(body, ok) || function () {};
        ok.addEventListener("click", function () { commit(); closeModal(); });
        cancel.addEventListener("click", closeModal);
        mask.addEventListener("click", function (e) { if (e.target === mask) closeModal(); });
        document.addEventListener("keydown", modalEsc);
    }
    function modalEsc(e) { if (e.key === "Escape") closeModal(); }
    function closeModal() {
        var m = $("#modal");
        if (m) m.remove();
        document.removeEventListener("keydown", modalEsc);
    }

    // ---------------- 拖拽排序（自定义模式） ----------------
    function setupDrag(card, meta) {
        card.draggable = (state.sort === "自定义");
        card.addEventListener("dragstart", function (e) {
            e.dataTransfer.setData("text/plain", meta.id);
            e.dataTransfer.effectAllowed = "move";
            card.classList.add("dragging");
        });
        card.addEventListener("dragend", function () {
            card.classList.remove("dragging");
            document.querySelectorAll(".drag-over-top,.drag-over-bottom").forEach(function (c) {
                c.classList.remove("drag-over-top", "drag-over-bottom");
            });
        });
        card.addEventListener("dragover", function (e) {
            if (state.sort !== "自定义") return;
            e.preventDefault();
            e.dataTransfer.dropEffect = "move";
            var r = card.getBoundingClientRect();
            var before = (e.clientY - r.top) < r.height / 2;
            card.classList.toggle("drag-over-top", before);
            card.classList.toggle("drag-over-bottom", !before);
        });
        card.addEventListener("drop", function (e) {
            if (state.sort !== "自定义") return;
            e.preventDefault();
            var srcId = e.dataTransfer.getData("text/plain");
            if (!srcId || srcId === meta.id) return;
            var r = card.getBoundingClientRect();
            var before = (e.clientY - r.top) < r.height / 2;

            // 组装新顺序（基于已渲染子集）
            var ids = state.items.map(function (m) { return m.id; });
            var si = ids.indexOf(srcId);
            if (si < 0) return;
            ids.splice(si, 1);
            var di = ids.indexOf(meta.id);
            ids.splice(before ? di : di + 1, 0, srcId);

            Bridge.reorder(ids).then(function () {
                state.sort = "自定义";
                reload();
            }).catch(function (e2) { toast(e2.message, true); });
        });
    }

    // ---------------- 主题 / 卡片大小 ----------------
    function applyTheme(theme) {
        document.documentElement.dataset.theme = theme === "dark" ? "dark" : "light";
    }

    function applyCardSize(size) {
        document.documentElement.dataset.cardsize = size || "compact";
    }

    // ---------------- TouchUsed 兼容桥 ----------------
    // C# 的 insert/applyStyle 处理器已内置 TouchUsed，无需单独消息；
    // 此处补一个空实现避免 bridge.js 里 Bridge.touch 未定义报错。
    if (!Bridge.touch) {
        Bridge.touch = function () {};
    }
})();
