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
        kind: "全部",           // 类型过滤：全部/Shape/Style
        tag: "",                // 标签过滤：""=全部
        sort: "最新",
        items: [],            // 已渲染条目
        filteredTotal: 0,
        total: 0,
        selected: {},         // id -> true（跨分页保留）
        selectedOrder: [],
        loading: false,
        adjustVisible: false,
        lastInsertedId: null,
        recentTop: []         // 全库最近使用 Top6（init 下发）
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
    function getUrlParam(name) {
        var m = new RegExp("[?&]" + name + "=([^&]+)").exec(window.location.search);
        return m ? decodeURIComponent(m[1]) : "";
    }

    document.addEventListener("DOMContentLoaded", function () {
        var popupView = getUrlParam("view");   // 独立弹窗：?view=toolX / settings
        if (popupView) {
            buildPopupView(popupView);
            return;
        }
        buildSkeleton();
        bindEvents();
        Bridge.init().then(function (data) {
            state.settings = data;
            state.categoryCounts = data.categoryCounts || {};
            state.recentTop = data.recentTop || [];
            state._tagCounts = data.tagCounts || {};
            applyTheme(data.theme || "light");
            applyCardSize(data.cardSize || "compact");
            renderCategories(data.categories || []);
            renderTagSelect(data.tags || []);
            reload();
        }).catch(function (err) {
            toast("初始化失败：" + err.message, true);
        });
    });

    // 订阅 C# 下行推送
    Bridge.on("refresh", function () {
        // 数据变了：重拉元数据（recentTop/分类/计数/标签），再刷新列表
        Bridge.init().then(function (data) {
            state.recentTop = data.recentTop || [];
            state.categoryCounts = data.categoryCounts || {};
            state._tagCounts = data.tagCounts || {};
            renderCategories(data.categories || []);
            renderTagSelect(data.tags || []);
        }).catch(function () {});
        reload();
    });
    Bridge.on("settingsChanged", function (s) {
        state.settings = s;
        applyTheme(s.theme || "light");
        applyCardSize(s.cardSize || "compact");
        reload();
    });

    // C#（Ribbon 按钮/设置对话框）请求切换到某视图
    Bridge.on("switchView", function (payload) {
        var view = (payload && payload.view) || "library";
        if (view === "tools") view = "tool1";   // 旧入口（Ribbon「小工具」）→ 工具首项
        switchView(view);
        activateNavTab(view);
    });

    // ---------------- 骨架（素材库视图；右侧任务窗格只显示素材库） ----------------
    function buildSkeleton() {
        var app = $("#app");
        app.innerHTML = "";
        app.classList.remove("popup-mode");

        // 素材库容器
        var libView = el("div", "lib-view");
        libView.id = "libView";
        app.appendChild(libView);

        // ⓪ 素材库标题条：图标+名称（左） / 库管理+主题切换（右）
        var topBar = el("div", "lib-topbar");
        var tbBrand = el("div", "lib-title");
        var tbIc = el("span", "msym", "photo_library");
        tbIc.setAttribute("aria-hidden", "true");
        tbBrand.appendChild(tbIc);
        tbBrand.appendChild(el("span", "", "素材库"));
        topBar.appendChild(tbBrand);
        var tbRight = el("div", "lib-topbar-right");
        var mgmtBtn2 = el("button", "topbar-btn");
        mgmtBtn2.id = "mgmtToggle";
        var mgIc = el("span", "msym", "tune");
        mgIc.setAttribute("aria-hidden", "true");
        mgmtBtn2.appendChild(mgIc);
        mgmtBtn2.title = "库管理：标签管理 / 去重清理 / 导入素材包";
        mgmtBtn2.addEventListener("click", function () {
            showMgmtMenu(mgmtBtn2);
        });
        tbRight.appendChild(mgmtBtn2);
        var themeBtn2 = el("button", "topbar-btn");
        themeBtn2.id = "themeToggle2";
        var thIc = el("span", "msym", "dark_mode");
        thIc.setAttribute("aria-hidden", "true");
        themeBtn2.appendChild(thIc);
        themeBtn2.title = "切换深浅主题";
        themeBtn2.addEventListener("click", function () {
            var next = (state.settings.theme === "dark") ? "light" : "dark";
            Bridge.saveSettings({ theme: next }).then(function (s) {
                state.settings = s;
                applyTheme(s.theme);
                thIc.textContent = next === "dark" ? "light_mode" : "dark_mode";
            }).catch(function (e) { toast(e.message, true); });
        });
        tbRight.appendChild(themeBtn2);
        topBar.appendChild(tbRight);
        libView.appendChild(topBar);

        // ① 头部
        var header = el("div", "header");
        var hrow = el("div", "header-row");
        var sw = el("div", "search-wrap");
        var sico = el("span", "msym ico-search", "search");
        sico.setAttribute("aria-hidden", "true");
        sw.appendChild(sico);
        var search = el("input");
        search.id = "searchInput";
        search.placeholder = "搜索 名称 / 分类 / 标签";
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
        libView.appendChild(header);

        // ② 分类条
        var catBar = el("div", "category-bar");
        catBar.id = "categoryBar";
        app.appendChild(catBar);

        // ②.5 类型段选 + 标签下拉
        var filterRow = el("div", "filter-row");
        var kindWrap = el("div", "kind-tabs");
        ["全部", "Shape", "Style"].forEach(function (k) {
            var b = el("button", "chip" + (k === "全部" ? " active" : ""), k === "Shape" ? "形状" : k === "Style" ? "样式" : "全部");
            b.dataset.kind = k;
            b.addEventListener("click", function () {
                state.kind = k;
                kindWrap.querySelectorAll(".chip").forEach(function (x) { x.classList.remove("active"); });
                b.classList.add("active");
                reload();
            });
            kindWrap.appendChild(b);
        });
        filterRow.appendChild(kindWrap);
        var tagSel = el("select", "sort-select");
        tagSel.id = "tagSelect";
        tagSel.title = "按标签筛选";
        var opt = el("option", "", "全部标签"); opt.value = "";
        tagSel.appendChild(opt);
        tagSel.addEventListener("change", function () {
            state.tag = this.value;
            reload();
        });
        filterRow.appendChild(tagSel);
        app.appendChild(filterRow);

        // ③ 状态行（计数；管理/主题在顶部标题条）
        var status = el("div", "status-row");
        var cnt = el("span"); cnt.id = "lblCount";
        status.appendChild(cnt);
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
        var undoBtn = el("button", "bar-btn", "\uE7A7");
        undoBtn.title = "撤销上一步（插入/微调/应用样式）";
        undoBtn.addEventListener("click", function () {
            Bridge.undo().catch(function (e) { toast(e.message, true); });
        });
        adjust.appendChild(undoBtn);
        var sep1 = el("span", "bar-sep");
        adjust.appendChild(sep1);
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
        var bPack = el("button", "bar-btn", "导出素材包");
        bPack.addEventListener("click", function () {
            var ids = Object.keys(state.selected);
            if (ids.length === 0) return;
            exportPackPrompt(ids, "素材包_" + ids.length + "项");
        });
        batch.appendChild(bPack);
        var bDel = el("button", "bar-btn danger", "删除");
        bDel.addEventListener("click", batchDelete);
        batch.appendChild(bDel);
        app.appendChild(batch);

        // Toast 容器（放 #app 顶层，全视图可见）
        var toastWrap = el("div", "toast-wrap");
        toastWrap.className = "toast-wrap";
        app.appendChild(toastWrap);

        // 素材库内容已直接构建进 #libView；header..batch 若挂在 #app 则收进 #libView（toast 除外）
        var libViewBox = $("#libView");
        Array.prototype.slice.call(app.children).forEach(function (c) {
            if (c === libViewBox) return;
            if (c.classList && c.classList.contains("toast-wrap")) return;
            libViewBox.appendChild(c);
        });
    }

    // ---------------- 视图切换 ----------------
    var currentView = "library";

    // 顶部导航项（工具平铺）。icon 为 Material Symbols ligature 名。
    var NAV_ITEMS = [
        { view: "library", label: "素材库", icon: "photo_library" },
        { view: "tool1", label: "文字样式", icon: "text_fields" },
        { view: "tool2", label: "一键换字体", icon: "font_download" },
        { view: "tool3", label: "查找替换", icon: "manage_search" },
        { view: "tool4", label: "文字清理", icon: "cleaning_services" },
        { view: "tool5", label: "导出图片", icon: "image" },
        { view: "tool6", label: "清理瘦身", icon: "delete_sweep" },
        { view: "tool7", label: "统一尺寸", icon: "aspect_ratio" },
        { view: "tool8", label: "图片网格", icon: "grid_view" },
        { view: "settings", label: "设置", icon: "settings" }
    ];

    function activateNavTab(view) {
        var tabs = document.querySelectorAll("#appNavTabs .nav-tab");
        tabs.forEach(function (x) {
            x.classList.toggle("active", x.dataset.view === view);
        });
    }

    function switchView(name) {
        currentView = name;
        var lib = $("#libView");
        var tools = $("#toolsPanel");
        var set = $("#settingsPanel");
        if (lib) lib.style.display = (name === "library") ? "" : "none";
        if (tools) tools.style.display = (name === "settings") ? "none" : (name === "library") ? "none" : "";
        if (set) set.style.display = (name === "settings") ? "" : "none";
        if (tools) {
            // 工具视图内：只显示与 name 对应的工具 section
            var sections = tools.querySelectorAll(".tool-sec");
            sections.forEach(function (s) {
                s.style.display = (s.dataset.tool === name) ? "" : "none";
            });
        }
        if (name === "library") { reload(); }
        if (name === "settings") { renderSettingsForm(); }
        if (name && name.indexOf("tool") === 0) { onToolsPanelShown(); }
    }

    // ---------------- 独立弹窗模式（?view=toolX / settings） ----------------
    // 弹窗由 ToolPopupHostForm 以 WebView2 打开，仅渲染单个工具/设置，全窗口铺满，
    // 不出现素材库导航/网格，从而在小窗口内自适应布局、不遮挡。
    function buildPopupView(viewKey) {
        var app = $("#app");
        app.innerHTML = "";
        app.classList.add("popup-mode");

        var root = el("div", "popup-root");
        root.id = "popupRoot";
        app.appendChild(root);

        var head = el("div", "popup-head");
        var title = el("div", "popup-title");
        head.appendChild(title);
        // 无边框弹窗：head 承担标题栏（可拖动），右侧关闭钮
        head.classList.add("popup-head-drag");
        var closeBtn = el("button", "popup-close", "");
        var cx = el("span", "msym", "close");
        cx.setAttribute("aria-hidden", "true");
        closeBtn.appendChild(cx);
        closeBtn.title = "关闭";
        closeBtn.addEventListener("click", function () {
            try { window.chrome.webview.postMessage(JSON.stringify({ type: "popupClose" })); }
            catch (e) {}
        });
        head.appendChild(closeBtn);
        root.appendChild(head);

        // 按住标题栏拖动（无边框窗口移动由宿主处理）
        head.addEventListener("mousedown", function (e) {
            if (e.target === closeBtn || closeBtn.contains(e.target)) return;
            try { window.chrome.webview.postMessage(JSON.stringify({ type: "popupDrag" })); }
            catch (e2) {}
        });

        var body = el("div", "popup-body");
        body.id = "popupBody";
        root.appendChild(body);

        var titleMap = {
            tool1: "文字样式", tool2: "一键换字体", tool3: "查找 / 替换", tool4: "文字清理",
            tool5: "导出图片", tool6: "清理瘦身", tool7: "统一尺寸", tool8: "图片网格",
            settings: "设置"
        };
        var iconMap = {
            tool1: "text_fields", tool2: "font_download", tool3: "manage_search", tool4: "cleaning_services",
            tool5: "image", tool6: "delete_sweep", tool7: "aspect_ratio", tool8: "grid_view",
            settings: "settings"
        };
        if (iconMap[viewKey]) {
            var tIc = el("span", "msym", iconMap[viewKey]);
            tIc.setAttribute("aria-hidden", "true");
            title.appendChild(tIc);
        }
        title.appendChild(el("span", "", titleMap[viewKey] || "工具"));

        var content = null;
        switch (viewKey) {
            case "tool1": content = buildTextStyleControls(); break;
            case "tool2": content = buildFontReplaceControls(); break;
            case "tool3": content = buildFindReplaceControls(); break;
            case "tool4": content = buildCleanControls(); break;
            case "tool5": content = buildExportImgControls(); break;
            case "tool6": content = buildScanCleanControls(); break;
            case "tool7": content = buildResizeControls(); break;
            case "tool8": content = buildImageGridControls(); break;
            case "settings": content = null; break;   // 设置用专用渲染
            default: content = null;
        }
        if (viewKey === "settings") {
            // 复用设置表单（需 state.settings 就绪）
            Bridge.init().then(function (data) {
                state.settings = data;
                applyTheme(data.theme || "light");
                renderPopupSettings();
            }).catch(function (err) { toast("初始化失败：" + err.message, true); });
            return;
        }
        if (content) {
            var card = el("div", "tool-card");
            card.appendChild(content);
            body.appendChild(card);
        }
        Bridge.init().then(function (data) {
            state.settings = data;
            applyTheme(data.theme || "light");
            if (viewKey.indexOf("tool") === 0) onToolsPanelShown();
        }).catch(function () {});
    }

    // 弹窗内的设置表单（把原生弹窗所需字段放一张卡内）
    function renderPopupSettings() {
        var body = $("#popupBody");
        if (!body) return;
        var s = state.settings || {};
        var card = el("div", "tool-card");
        card.appendChild(el("div", "tool-card-title", "偏好设置"));

        var theme = makeSelect(["浅色", "深色"], s.theme === "dark" ? 1 : 0);
        var insert = makeSelect(["页面居中", "原始位置"], s.insertPosition === "original" ? 1 : 0);
        var cardSize = makeSelect(["紧凑", "标准", "大图"],
            s.cardSize === "large" ? 2 : s.cardSize === "standard" ? 1 : 0);
        var pageSize = makeInput("每页数量", s.pageSize || 200);
        var autoName = makeCheck("按形状类型自动命名", !!s.autoName);
        var clickToInsert = makeCheck("单击卡片直接插入（取消则单击预览、双击插入）", s.clickToInsert !== false);
        var themeAdapt = makeCheck("应用样式时把主色映射到当前文档主题色", !!s.themeAdaptStyle);
        var widthNum = makeInput("面板宽度", s.paneWidth || 300);

        var save = runBtn("保存设置", function () {
            return Bridge.saveSettings({
                theme: theme.selectedIndex === 1 ? "dark" : "light",
                insertPosition: insert.selectedIndex === 1 ? "original" : "center",
                cardSize: cardSize.selectedIndex === 2 ? "large" : cardSize.selectedIndex === 1 ? "standard" : "compact",
                pageSize: parseInt(pageSize.value, 10) || 200,
                autoName: autoName.querySelector("input").checked,
                clickToInsert: clickToInsert.querySelector("input").checked,
                themeAdaptStyle: themeAdapt.querySelector("input").checked,
                paneWidth: parseInt(widthNum.value, 10) || 300
            }).then(function (d) {
                state.settings = d;
                applyTheme(d.theme); applyCardSize(d.cardSize);
                toast("设置已保存");
            });
        });

        card.appendChild(fld("主题", theme));
        card.appendChild(fld("插入位置", insert));
        card.appendChild(fld("卡片大小", cardSize));
        card.appendChild(fld("分页数量", pageSize));
        card.appendChild(fld("面板宽度", widthNum));
        card.appendChild(fld("", autoName));
        card.appendChild(fld("", clickToInsert));
        card.appendChild(fld("", themeAdapt));
        card.appendChild(fld("", save));
        body.appendChild(card);
    }

    // ---------------- 工具面板 ----------------
    var toolState = { fonts: [], hasPres: false };

    function buildToolsPanel(app) {
        var panel = el("div", "view-panel");
        panel.id = "toolsPanel";
        panel.style.display = "none";

        var head = el("div", "view-panel-head");
        head.appendChild(el("div", "view-panel-title", "工具"));
        var backLib = el("button", "bar-btn", "\uE711 返回素材库");
        backLib.addEventListener("click", function () { switchView("library"); });
        head.appendChild(backLib);
        panel.appendChild(head);

        var body = el("div", "view-panel-body");
        body.id = "toolsBody";
        // 每个工具一个 .tool-sec（由顶部导航对应项切换显示）
        body.appendChild(buildToolSec("tool1", "批量修改文字样式", buildTextStyleControls()));
        body.appendChild(buildToolSec("tool2", "一键更换字体", buildFontReplaceControls()));
        body.appendChild(buildToolSec("tool3", "查找 / 替换（保留格式）", buildFindReplaceControls()));
        body.appendChild(buildToolSec("tool4", "文字清理", buildCleanControls()));
        body.appendChild(buildToolSec("tool5", "整稿导出 PNG/JPG", buildExportImgControls()));
        body.appendChild(buildToolSec("tool6", "清理瘦身", buildScanCleanControls()));
        body.appendChild(buildToolSec("tool7", "统一尺寸", buildResizeControls()));
        body.appendChild(buildToolSec("tool8", "图片网格插入", buildImageGridControls()));
        panel.appendChild(body);
        app.appendChild(panel);
    }

    function buildToolSec(toolId, title, content) {
        var sec = el("div", "tool-sec");
        sec.dataset.tool = toolId;
        sec.style.display = "none";
        sec.appendChild(buildToolCard(title, content));
        return sec;
    }

    function buildToolCard(title, content) {
        var card = el("div", "tool-card");
        card.appendChild(el("div", "tool-card-title", title));
        card.appendChild(content);
        return card;
    }

    /** 通用表单行：label + 控件 */
    function fld(label, ctl) {
        var row = el("div", "fld");
        row.appendChild(el("div", "fld-label", label));
        row.appendChild(ctl);
        return row;
    }
    function makeSelect(items, selIdx) {
        var s = el("select", "ctl");
        items.forEach(function (it, i) {
            var o = el("option", "", it); o.value = it; s.appendChild(o);
        });
        if (selIdx != null) s.selectedIndex = selIdx;
        return s;
    }
    function makeInput(ph, val) {
        var i = el("input", "ctl");
        i.placeholder = ph || "";
        if (val != null) i.value = val;
        return i;
    }
    function makeCheck(label, checked) {
        var l = el("label", "chk");
        var c = el("input"); c.type = "checkbox"; c.checked = !!checked;
        l.appendChild(c);
        l.appendChild(el("span", "", label));
        return l;
    }
    function runBtn(text, fn) {
        var b = el("button", "bar-btn primary run-btn", text);
        b.addEventListener("click", function () {
            var self = this;
            self.disabled = true;
            fn().catch(function (e) { toast(e.message, true); })
                .then(function () { self.disabled = false; });
        });
        return b;
    }

    // ① 文字样式
    function buildTextStyleControls() {
        var box = el("div", "tool-card-body");
        var scope = makeSelect(["选中形状", "当前页", "全部页面"], 1);
        var font = makeInput("字体名（留空不改）");
        var size = makeInput("字号（留空不改）");
        var bold = makeSelect(["加粗：不改", "加粗", "取消加粗"], 0);
        var align = makeSelect(["对齐：不改", "左对齐", "居中", "右对齐"], 0);
        var run = runBtn("应用", function () {
            return Bridge.applyTextStyle({
                scope: scope.selectedIndex === 0 ? "selection" : scope.selectedIndex === 1 ? "page" : "all",
                fontName: font.value.trim(),
                fontSize: parseFloat(size.value) || 0,
                bold: bold.selectedIndex === 1 ? true : bold.selectedIndex === 2 ? false : null,
                italic: null,
                color: 0,
                align: align.selectedIndex === 1 ? "left" : align.selectedIndex === 2 ? "center" : align.selectedIndex === 3 ? "right" : ""
            }).then(function (d) { toast(d && d.changed != null ? "已处理 " + d.changed + " 个文字形状。" : "完成"); });
        });
        box.appendChild(fld("作用范围", scope));
        box.appendChild(fld("字体", font));
        box.appendChild(fld("字号", size));
        box.appendChild(fld("加粗", bold));
        box.appendChild(fld("对齐", align));
        box.appendChild(fld("", run));
        return box;
    }

    // ② 一键换字体
    function buildFontReplaceControls() {
        var box = el("div", "tool-card-body");
        var from = el("select", "ctl"); from.id = "frFromFont";
        var to = el("select", "ctl"); to.id = "frToFont";
        var common = ["微软雅黑", "宋体", "黑体", "仿宋", "楷体", "等线", "Arial", "Times New Roman", "Calibri", "Segoe UI", "Georgia", "Verdana"];
        var scope = makeSelect(["整份演示文稿", "当前页"], 0);
        var run = runBtn("开始替换", function () {
            return Bridge.replaceFont({ fromFont: from.value, toFont: to.value, allSlides: scope.selectedIndex === 0 })
                .then(function (d) { toast(d && d.message ? d.message : "完成"); });
        });
        box.appendChild(fld("原字体", from));
        box.appendChild(fld("新字体", to));
        box.appendChild(fld("范围", scope));
        box.appendChild(fld("", run));
        return box;
    }

    // ③ 查找替换
    function buildFindReplaceControls() {
        var box = el("div", "tool-card-body");
        var find = makeInput("查找内容");
        var rep = makeInput("替换为");
        var scope = makeSelect(["整份演示文稿", "当前页"], 0);
        var matchCase = makeCheck("区分大小写", false);
        var run = runBtn("替换全部", function () {
            if (!find.value) { toast("请输入查找内容。", true); return Promise.resolve(); }
            return Bridge.findReplace({
                find: find.value, replace: rep.value,
                matchCase: matchCase.querySelector("input").checked,
                allSlides: scope.selectedIndex === 0
            }).then(function (d) { toast(d && d.message ? d.message : "完成"); });
        });
        box.appendChild(fld("查找", find));
        box.appendChild(fld("替换为", rep));
        box.appendChild(fld("范围", scope));
        box.appendChild(fld("", matchCase));
        box.appendChild(fld("", run));
        return box;
    }

    // ④ 文字清理
    function buildCleanControls() {
        var box = el("div", "tool-card-body");
        var ws = makeCheck("去多余空白（全角空格/连续空格/行首尾）", true);
        var punct = makeCheck("中文间西文标点转全角", false);
        var note = el("div", "tool-note", "作用于整份演示文稿。文本会按首字符格式统一。");
        var run = runBtn("开始清理", function () {
            return Bridge.cleanText({
                trimWhitespace: ws.querySelector("input").checked,
                unifyPunctuation: punct.querySelector("input").checked
            }).then(function (d) { toast(d && d.message ? d.message : "完成"); });
        });
        box.appendChild(fld("", ws));
        box.appendChild(fld("", punct));
        box.appendChild(note);
        box.appendChild(fld("", run));
        return box;
    }

    function onToolsPanelShown() {
        // 刷新字体下拉
        Bridge.toolGetFonts().then(function (d) {
            toolState.fonts = d.fonts || [];
            toolState.hasPres = !!d.hasPresentation;
            fillSelect($("#frFromFont"), toolState.fonts, true);
            fillSelect($("#frToFont"), ["微软雅黑", "宋体", "黑体", "仿宋", "楷体", "等线", "Arial", "Times New Roman", "Calibri", "Segoe UI", "Georgia", "Verdana"], false);
        }).catch(function () {});
    }

    // ================= 版式批处理工具（tool5-8） =================

    // ⑤ 整稿导出 PNG/JPG
    var exportImgState = { dir: "" };
    function buildExportImgControls() {
        var box = el("div", "tool-card-body");
        var fmt = makeSelect(["PNG", "JPG"], 0);
        var scale = makeSelect(["原始大小", "宽 1280", "宽 1920"], 0);
        var dirInput = makeInput("选择导出目录", exportImgState.dir || "");
        dirInput.id = "expDir";
        var dirBtn = el("button", "bar-btn", "浏览…");
        dirBtn.addEventListener("click", function () {
            Bridge.requestFolder("选择图片导出目录").then(function (d) {
                if (d && d.path) { exportImgState.dir = d.path; dirInput.value = d.path; }
            }).catch(function (e) { toast(e.message, true); });
        });
        var dirRow = el("div", "dir-row");
        dirRow.appendChild(dirInput);
        dirRow.appendChild(dirBtn);
        var run = runBtn("开始导出", function () {
            var dir = dirInput.value.trim();
            if (!dir) { toast("请选择导出目录。", true); return Promise.resolve(); }
            var w = scale.selectedIndex === 1 ? 1280 : scale.selectedIndex === 2 ? 1920 : 0;
            return Bridge.formatExportImages({ dir: dir, format: fmt.selectedIndex === 1 ? "JPG" : "PNG", width: w })
                .then(function (d) { toast(d && d.message ? d.message : "导出完成", !d || !d.ok); });
        });
        box.appendChild(fld("格式", fmt));
        box.appendChild(fld("分辨率", scale));
        var dirLbl = el("div", "fld-label", "目录");
        var dirWrap = el("div"); dirWrap.style.flex = "1";
        dirWrap.appendChild(dirInput);
        box.appendChild(dirLbl);
        var fullRow = el("div", "dir-row");   // 目录输入+浏览 同行
        fullRow.appendChild(dirWrap);
        fullRow.appendChild(dirBtn);
        box.appendChild(fullRow);
        box.appendChild(fld("", run));
        return box;
    }

    // ⑥ 清理瘦身（扫描 + 执行）
    var cleanOpts = { delEmpty: true, delHidden: false, delOffSlide: false, clearNotes: true, removeActions: false };
    function buildScanCleanControls() {
        var box = el("div", "tool-card-body");
        var cEmpty = makeCheck("删除空白幻灯片", cleanOpts.delEmpty);
        var cHidden = makeCheck("删除隐藏对象", cleanOpts.delHidden);
        var cOff = makeCheck("删除页面外对象（完全超出画布）", cleanOpts.delOffSlide);
        var cNotes = makeCheck("清空备注文字", cleanOpts.clearNotes);
        var cActions = makeCheck("移除形状动作/超链接", cleanOpts.removeActions);
        var preview = el("div", "tool-note", "（尚未扫描）"); preview.id = "cleanPreview";
        var scanBtn = runBtn("扫描统计", function () {
            cleanOpts.delEmpty = cEmpty.querySelector("input").checked;
            cleanOpts.delHidden = cHidden.querySelector("input").checked;
            cleanOpts.delOffSlide = cOff.querySelector("input").checked;
            cleanOpts.clearNotes = cNotes.querySelector("input").checked;
            cleanOpts.removeActions = cActions.querySelector("input").checked;
            return Bridge.formatScanClean(cleanOpts).then(function (d) {
                if (preview) preview.textContent = d && d.message ? d.message : "扫描完成";
                toast(d && d.ok ? "扫描完成" : (d && d.message || "扫描失败"), !d || !d.ok);
            });
        });
        var runBtn2 = runBtn("执行清理", function () {
            if (!window.confirm("执行后将不可恢复（建议先保存文档），确定继续？")) {
                return Promise.resolve();
            }
            cleanOpts.delEmpty = cEmpty.querySelector("input").checked;
            cleanOpts.delHidden = cHidden.querySelector("input").checked;
            cleanOpts.delOffSlide = cOff.querySelector("input").checked;
            cleanOpts.clearNotes = cNotes.querySelector("input").checked;
            cleanOpts.removeActions = cActions.querySelector("input").checked;
            return Bridge.formatRunClean(cleanOpts).then(function (d) {
                toast(d && d.message ? d.message : "清理完成", !d || !d.ok);
            });
        });
        box.appendChild(fld("", cEmpty));
        box.appendChild(fld("", cHidden));
        box.appendChild(fld("", cOff));
        box.appendChild(fld("", cNotes));
        box.appendChild(fld("", cActions));
        box.appendChild(fld("", scanBtn));
        box.appendChild(fld("", runBtn2));
        box.appendChild(preview);
        return box;
    }

    // ⑦ 统一尺寸
    function buildResizeControls() {
        var box = el("div", "tool-card-body");
        var scope = makeSelect(["选中形状", "当前页全部形状", "全部页面形状"], 0);
        var w = makeInput("宽度（磅，可留空）");
        var h = makeInput("高度（磅，可留空）");
        var forceRatio = makeCheck("图片强制等比", false);
        var run = runBtn("应用尺寸", function () {
            var wv = parseFloat(w.value) || 0;
            var hv = parseFloat(h.value) || 0;
            if (!wv && !hv) { toast("请填写宽度或高度。", true); return Promise.resolve(); }
            return Bridge.formatResize({
                scope: scope.selectedIndex === 0 ? "selection" : scope.selectedIndex === 1 ? "page" : "all",
                width: wv, height: hv,
                forceImageRatio: forceRatio.querySelector("input").checked
            }).then(function (d) { toast(d && d.message ? d.message : "完成"); });
        });
        box.appendChild(fld("范围", scope));
        box.appendChild(fld("宽度", w));
        box.appendChild(fld("高度", h));
        box.appendChild(fld("", forceRatio));
        box.appendChild(fld("", run));
        return box;
    }

    // ⑧ 图片网格插入
    var gridFiles = [];
    function buildImageGridControls() {
        var box = el("div", "tool-card-body");
        var cols = el("input", "ctl"); cols.type = "number"; cols.min = "1"; cols.max = "8"; cols.value = "3";
        var insertMode = makeSelect(["插入到新幻灯片（排满自动翻页）", "插入到当前页"], 0);
        var filesNote = el("div", "tool-note", "尚未选择图片。"); filesNote.id = "gridFilesNote";
        var pickBtn = el("button", "bar-btn", "选择图片…");
        pickBtn.addEventListener("click", function () {
            Bridge.requestFiles("选择图片", "图片文件 (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp")
                .then(function (d) {
                    gridFiles = (d && d.files) || [];
                    if (filesNote) filesNote.textContent = gridFiles.length > 0 ? "已选 " + gridFiles.length + " 张：" + gridFiles[0] + (gridFiles.length > 1 ? " 等" : "") : "尚未选择图片。";
                }).catch(function (e) { toast(e.message, true); });
        });
        var run = runBtn("生成网格", function () {
            if (!gridFiles.length) { toast("请先选择图片。", true); return Promise.resolve(); }
            return Bridge.formatImageGrid({
                files: gridFiles,
                cols: parseInt(cols.value, 10) || 3,
                toNewPage: insertMode.selectedIndex === 0
            }).then(function (d) { toast(d && d.message ? d.message : "完成"); });
        });
        box.appendChild(fld("每行张数", cols));
        box.appendChild(fld("插入方式", insertMode));
        box.appendChild(fld("", filesNote));
        box.appendChild(fld("", pickBtn));
        box.appendChild(fld("", run));
        return box;
    }

    function fillSelect(sel, items, keepEmpty) {
        if (!sel) return;
        sel.innerHTML = "";
        if (keepEmpty) {
            var o = el("option", "", "（文档中出现的字体）"); o.value = "";
            sel.appendChild(o);
        }
        items.forEach(function (it) {
            var o = el("option", "", it); o.value = it; sel.appendChild(o);
        });
        if (sel.options.length > 0 && !keepEmpty) sel.selectedIndex = 0;
    }

    // ---------------- 设置面板 ----------------
    function buildSettingsPanel(app) {
        var panel = el("div", "view-panel");
        panel.id = "settingsPanel";
        panel.style.display = "none";
        var head = el("div", "view-panel-head");
        head.appendChild(el("div", "view-panel-title", "\uE713 设置"));
        var backLib2 = el("button", "bar-btn", "\uE711 返回素材库");
        backLib2.addEventListener("click", function () { switchView("library"); });
        head.appendChild(backLib2);
        panel.appendChild(head);
        var body = el("div", "view-panel-body");
        body.id = "settingsBody";
        panel.appendChild(body);
        app.appendChild(panel);
    }

    function renderSettingsForm() {
        var body = $("#settingsBody");
        if (!body) return;
        var s = state.settings || {};
        body.innerHTML = "";

        var card = el("div", "tool-card");
        card.appendChild(el("div", "tool-card-title", "偏好设置"));

        var theme = makeSelect(["浅色", "深色"], s.theme === "dark" ? 1 : 0);
        var insert = makeSelect(["页面居中", "原始位置"], s.insertPosition === "original" ? 1 : 0);
        var cardSize = makeSelect(["紧凑", "标准", "大图"],
            s.cardSize === "large" ? 2 : s.cardSize === "standard" ? 1 : 0);
        var pageSize = makeInput("每页数量", s.pageSize || 200);
        var autoName = makeCheck("按形状类型自动命名", !!s.autoName);
        var clickToInsert = makeCheck("单击卡片直接插入（取消则单击预览、双击插入）", s.clickToInsert !== false);
        var themeAdapt = makeCheck("应用样式时把主色映射到当前文档主题色", !!s.themeAdaptStyle);
        var widthNum = makeInput("面板宽度", s.paneWidth || 300);

        var save = runBtn("保存设置", function () {
            return Bridge.saveSettings({
                theme: theme.selectedIndex === 1 ? "dark" : "light",
                insertPosition: insert.selectedIndex === 1 ? "original" : "center",
                cardSize: cardSize.selectedIndex === 2 ? "large" : cardSize.selectedIndex === 1 ? "standard" : "compact",
                pageSize: parseInt(pageSize.value, 10) || 200,
                autoName: autoName.querySelector("input").checked,
                clickToInsert: clickToInsert.querySelector("input").checked,
                themeAdaptStyle: themeAdapt.querySelector("input").checked,
                paneWidth: parseInt(widthNum.value, 10) || 300
            }).then(function (d) {
                state.settings = d;
                applyTheme(d.theme); applyCardSize(d.cardSize);
                toast("设置已保存");
            });
        });

        card.appendChild(fld("主题", theme));
        card.appendChild(fld("插入位置", insert));
        card.appendChild(fld("卡片大小", cardSize));
        card.appendChild(fld("分页数量", pageSize));
        card.appendChild(fld("面板宽度", widthNum));
        card.appendChild(fld("", autoName));
        card.appendChild(fld("", clickToInsert));
        card.appendChild(fld("", themeAdapt));
        card.appendChild(fld("", save));
        body.appendChild(card);

        var libCard = el("div", "tool-card");
        libCard.appendChild(el("div", "tool-card-title", "库管理"));
        var mgmtNote = el("div", "tool-note", "标签管理 / 去重清理 / 导入素材包 / 备份与导出见 Ribbon「素材库」选项卡。");
        libCard.appendChild(mgmtNote);
        var goLib = el("button", "bar-btn", "标签 / 去重等管理工具");
        goLib.addEventListener("click", function () {
            switchView("library");
            var btn = $("#mgmtToggle");
            if (btn) { closeContextMenu(); showMgmtMenu(btn); }
        });
        libCard.appendChild(goLib);
        body.appendChild(libCard);
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
        Bridge.list(state.filter, state.category, state.kind, state.tag, state.sort, offset, getPageSize())
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
        if (!bar) return;        bar.innerHTML = "";
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

    /** 填充"标签筛选"下拉（tags 来自 init 的 tags 数组），保留当前选中值。 */
    function renderTagSelect(tags) {
        var sel = $("#tagSelect");
        if (!sel) return;
        var cur = state.tag;
        sel.innerHTML = "";
        var opt = el("option", "", "全部标签"); opt.value = "";
        sel.appendChild(opt);
        (tags || []).forEach(function (t) {
            var o = el("option", "", t);
            o.value = t;
            sel.appendChild(o);
        });
        // 还原选中（若当前标签仍存在）
        sel.value = (cur && tags.indexOf(cur) >= 0) ? cur : "";
        state.tag = sel.value;
    }

    function updateStatus() {
        var cnt = $("#lblCount");
        if (!cnt) return;
        if (state.filter || state.category !== "全部" || state.kind !== "全部" || state.tag !== "") {
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

        // 无筛选时：顶部展示"最近使用"横向条（数据来自 init 的全库 Top6，A6）
        var recentExcluded = {};
        if (!state.filter && state.category === "全部" && state.kind === "全部" && state.tag === ""
            && state.sort !== "最近使用"
            && state.recentTop && state.recentTop.length >= 2) {
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
            state.recentTop.forEach(function (m) {
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

        if (state.items.length === 0) {
            var empty = $("#emptyState");
            empty.style.display = "";
            empty.innerHTML = '' +
                '<div class="empty-ico msym" aria-hidden="true">' + (state.filter || state.category !== "全部" ? "search_off" : "photo_library") + '</div>' +
                (state.filter || state.category !== "全部" || state.kind !== "全部" || state.tag !== ""
                    ? "<div class=\"empty-main\">没有符合筛选条件的素材</div><div class=\"empty-sub\">试试清空搜索词、切换分类/类型/标签</div>"
                    : "<div class=\"empty-main\">素材库还是空的</div><div class=\"empty-sub\">在演示文稿中选中形状，点击顶部「素材库」选项卡里的「收藏形状」「收藏整页」或「收藏样式」<br>也可以从「收藏剪贴板」「收藏整稿」快速批量收录</div>");
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

        var isStyle = meta.kind === "Style";

        // 形状卡：插入到当前页；样式卡：应用样式到选中/本页同类
        if (!isStyle) {
            mi("插入到当前页", function () { doInsert(meta, findCard(meta.id)); });
        } else {
            mi("应用样式到选中形状", function () {
                Bridge.applyStyle(meta.id).then(function () {
                    toast("样式已应用");
                }).catch(function (e) { toast(e.message, true); });
            });
            mi("应用到本页同类形状", function () {
                Bridge.applyStyleSameType(meta.id).then(function () {
                    toast("已应用到本页同类形状");
                }).catch(function (e) { toast(e.message, true); });
            });
            mi("应用到整份演示文稿同类型形状", function () {
                Bridge.applyStyleSameTypeAllSlides(meta.id).then(function () {
                    toast("已应用到整份演示文稿同类型形状");
                }).catch(function (e) { toast(e.message, true); });
            });
        }
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
        mi("导出到素材包…", function () { exportPackPrompt([meta.id], meta.name); });

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
        info.appendChild(el("div", "lb-kind", meta.kind === "Style" ? "样式素材 · 右键卡片可应用" : "形状素材"));
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

    // ---------------- 库管理（… 菜单） ----------------
    function showMgmtMenu(anchor) {
        closeMgmtMenu();
        var menu = el("div", "ctx-menu");
        menu.id = "mgmtMenu";
        menu.style.position = "fixed";
        function mi(text, iconName, fn) {
            var b = el("button", "mi");
            if (iconName) {
                var ic = el("span", "msym", iconName);
                ic.setAttribute("aria-hidden", "true");
                b.appendChild(ic);
            }
            b.appendChild(el("span", "", text));
            b.addEventListener("click", function () { closeMgmtMenu(); fn(); });
            menu.appendChild(b);
        }
        mi("标签管理…", "label", showTagManager);
        mi("去重清理…", "delete_sweep", showDuplicateCleaner);
        mi("导入素材包…", "package_2", importPackPrompt);
        document.body.appendChild(menu);

        var r = anchor.getBoundingClientRect();
        menu.style.left = Math.max(4, r.right - menu.offsetWidth) + "px";
        menu.style.top = (r.bottom + 4) + "px";
        setTimeout(function () {
            document.addEventListener("click", closeMgmtMenu, { once: true });
        }, 10);
    }
    function closeMgmtMenu() {
        var m = $("#mgmtMenu");
        if (m) m.remove();
    }

    // ---------------- 标签管理器（Web） ----------------
    function showTagManager() {
        Bridge.init().then(function (data) {
            state._tagCounts = data.tagCounts || {};
            openTagManagerModal();
        }).catch(function (e) { toast(e.message, true); });
    }

    function openTagManagerModal() {
        closeModal();
        var mask = el("div", "modal-mask");
        mask.id = "modal";
        var modal = el("div", "modal");
        modal.style.width = "300px";
        modal.appendChild(el("div", "modal-title", "标签管理"));
        var body = el("div");
        var list = el("div", "tag-mgr-list");
        list.id = "tagList";
        body.appendChild(list);
        modal.appendChild(body);
        var actions = el("div", "modal-actions");
        var closeBtn = el("button", "bar-btn primary", "关闭");
        actions.appendChild(closeBtn);
        modal.appendChild(actions);
        mask.appendChild(modal);
        document.body.appendChild(mask);
        closeBtn.addEventListener("click", function () { closeModal(); refreshAll(); });
        mask.addEventListener("click", function (e) { if (e.target === mask) closeModal(); });

        renderTagList(list);
    }

    function renderTagList(list) {
        list.innerHTML = "";
        state._tagCounts = state._tagCounts || {};
        var tags = Object.keys(state._tagCounts).sort();
        if (tags.length === 0) {
            list.appendChild(el("div", "modal-msg", "暂无标签。\n右键卡片「编辑标签」即可添加。"));
            return;
        }
        tags.forEach(function (tag) {
            var row = el("div", "tag-mgr-row");
            row.appendChild(el("span", "tag-mgr-name", tag + "（" + state._tagCounts[tag] + " 项）"));
            var rn = el("button", "tag-mgr-btn", "重命名");
            rn.addEventListener("click", function () {
                closeModal();
                inputModal("重命名标签", "新的标签名：", tag, function (v) {
                    if (!v || v === tag) return;
                    Bridge.renameTag(tag, v).then(function () {
                        refreshAll(); toast("已重命名标签");
                    }).catch(function (e) { toast(e.message, true); });
                });
            });
            var dl = el("button", "tag-mgr-btn danger", "删除");
            dl.addEventListener("click", function () {
                closeModal();
                confirmModal("从所有收藏中删除标签「" + tag + "」？", function () {
                    Bridge.deleteTag(tag).then(function () {
                        refreshAll(); toast("已删除标签");
                    }).catch(function (e) { toast(e.message, true); });
                });
            });
            row.appendChild(rn);
            row.appendChild(dl);
            list.appendChild(row);
        });
    }

    // ---------------- 去重清理 ----------------
    function showDuplicateCleaner() {
        closeModal();
        Bridge.scanDuplicates().then(function (data) {
            var dups = data.duplicates || [];
            var mask = el("div", "modal-mask");
            mask.id = "modal";
            var modal = el("div", "modal");
            modal.style.width = "320px";
            modal.appendChild(el("div", "modal-title", "去重清理"));
            var body = el("div");
            if (dups.length === 0) {
                body.appendChild(el("div", "modal-msg", "未发现完全重复的收藏。\n（同类型、同数量、同尺寸判定）"));
            } else {
                body.appendChild(el("div", "modal-msg", "发现 " + dups.length + " 条可能完全重复（保留最早）："));
                var list = el("div", "dup-list");
                dups.forEach(function (d) {
                    var row = el("div", "tag-mgr-row");
                    var lbl = el("label", "dup-item");
                    var cb = el("input");
                    cb.type = "checkbox";
                    cb.checked = true;
                    cb.dataset.id = d.id;
                    lbl.appendChild(cb);
                    lbl.appendChild(el("span", "", d.name + "（" + (d.kind === "Style" ? "样式" : "形状") + "）"));
                    row.appendChild(lbl);
                    list.appendChild(row);
                });
                body.appendChild(list);
            }
            modal.appendChild(body);
            var actions = el("div", "modal-actions");
            if (dups.length > 0) {
                var delBtn = el("button", "bar-btn danger", "删除所选");
                delBtn.addEventListener("click", function () {
                    var ids = Array.prototype.map.call(list.querySelectorAll("input:checked"), function (c) { return c.dataset.id; });
                    if (ids.length === 0) return;
                    Bridge.del(ids).then(function () {
                        closeModal(); refreshAll(); toast("已删除 " + ids.length + " 条重复项");
                    }).catch(function (e) { toast(e.message, true); });
                });
                actions.appendChild(delBtn);
            }
            var closeBtn = el("button", "bar-btn", "关闭");
            closeBtn.addEventListener("click", function () { closeModal(); refreshAll(); });
            actions.appendChild(closeBtn);
            modal.appendChild(actions);
            mask.appendChild(modal);
            document.body.appendChild(mask);
            mask.addEventListener("click", function (e) { if (e.target === mask) { closeModal(); refreshAll(); } });
        }).catch(function (e) { toast(e.message, true); });
    }

    // ---------------- 素材包 ----------------
    // 素材包文件存放在 %APPDATA%\PPTShapeLibrary\packs\ 目录（纯文件名输入，C# 侧拼路径）
    function exportPackPrompt(ids, hintName) {
        var name = (hintName || "素材包").replace(/[\\/:*?"<>|]/g, "_");
        inputModal("导出素材包", "素材包文件名（将保存到收藏库 packs 目录）：", name, function (v) {
            if (!v) return;
            Bridge.exportPack(ids, v).then(function () {
                toast("已导出素材包到 packs 目录");
            }).catch(function (e) { toast(e.message, true); });
        });
    }

    function importPackPrompt() {
        inputModal("导入素材包", "素材包文件名（从收藏库 packs 目录导入，如 MyPack）：", "", function (v) {
            if (!v) return;
            Bridge.importPack(v).then(function (d) {
                refreshAll();
                var conflicts = d.conflicts || 0;
                toast("已导入 " + (d.imported || 0) + " 个收藏" +
                      (conflicts > 0 ? "（其中 " + conflicts + " 个与本库 id 冲突已自动改名导入）" : ""));
            }).catch(function (e) { toast(e.message, true); });
        });
    }

    function refreshAll() {
        Bridge.init().then(function (data) {
            state.settings = data;
            state.recentTop = data.recentTop || [];
            state.categoryCounts = data.categoryCounts || {};
            state._tagCounts = data.tagCounts || {};
            renderCategories(data.categories || []);
            renderTagSelect(data.tags || []);
            reload();
        }).catch(function () { reload(); });
    }
})();
