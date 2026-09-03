/* ============================================================
   bridge.js — C#↔JS 通信层
   reqId 管理、pending 超时兜底、下行推送订阅
   ============================================================ */
(function () {
    "use strict";

    var reqSeq = 0;
    var pending = {};          // reqId -> {resolve, reject, timer}
    var listeners = {};        // type -> [fn]
    var TIMEOUT = 8000;

    function post(msg) {
        try {
            window.chrome.webview.postMessage(JSON.stringify(msg));
        } catch (e) {
            console.error("postMessage failed (not hosted in WebView2?)", e);
            // 非宿主环境（浏览器直接打开）下的降级：直接失败回调
            return false;
        }
        return true;
    }

    /** 发送请求并等待 result（返回 Promise） */
    function send(type, payload) {
        return new Promise(function (resolve, reject) {
            var reqId = "r" + (++reqSeq);
            pending[reqId] = {
                resolve: resolve,
                reject: reject,
                timer: setTimeout(function () {
                    if (pending[reqId]) {
                        delete pending[reqId];
                        reject(new Error("timeout: " + type));
                    }
                }, TIMEOUT)
            };
            if (!post({ type: type, reqId: reqId, payload: payload || {} })) {
                clearTimeout(pending[reqId].timer);
                delete pending[reqId];
                reject(new Error("no host"));
            }
        });
    }

    /** 订阅 C# 下行推送（refresh / settingsChanged） */
    function on(type, fn) {
        (listeners[type] = listeners[type] || []).push(fn);
    }

    // 接收 C# 消息
    window.chrome && window.chrome.webview && window.chrome.webview.addEventListener &&
    window.chrome.webview.addEventListener("message", function (e) {
        var msg;
        try { msg = typeof e.data === "string" ? JSON.parse(e.data) : e.data; }
        catch (err) { return; }
        if (!msg || !msg.type) return;

        if (msg.type === "result") {
            var p = pending[msg.reqId];
            if (p) {
                clearTimeout(p.timer);
                delete pending[msg.reqId];
                if (msg.ok) p.resolve(msg.data || {});
                else p.reject(new Error(msg.error || "unknown error"));
            }
            return;
        }

        // 下行推送
        var fns = listeners[msg.type];
        if (fns) {
            fns.forEach(function (fn) {
                try { fn(msg.payload || {}); } catch (err) { console.error(err); }
            });
        }
    });

    // 全局未捕获异常 → 写 C# 日志（P1）
    window.addEventListener("error", function (e) {
        try {
            post({ type: "log", payload: { level: "error", message: "JS uncaught: " + e.message + " @" + e.filename + ":" + e.lineno } });
        } catch (_) {}
    });

    window.Bridge = {
        send: send,
        on: on,
        /** 快捷方法 */
        init: function () { return send("init"); },
        list: function (filter, category, kind, tag, sort, offset, limit) {
            return send("list", { filter: filter, category: category, kind: kind, tag: tag, sort: sort, offset: offset, limit: limit });
        },
        insert: function (id) { return send("insert", { id: id }); },
        applyStyle: function (id) { return send("applyStyle", { id: id }); },
        applyStyleSameType: function (id) { return send("applyStyleSameType", { id: id }); },
        applyStyleSameTypeAllSlides: function (id) { return send("applyStyleSameTypeAllSlides", { id: id }); },
        collectClipboard: function () { return send("collectClipboard", {}); },
        collectAllSlides: function () { return send("collectAllSlides", {}); },
        adjust: function (action) { return send("adjust", { action: action }); },
        del: function (ids) { return send("delete", { ids: ids }); },
        rename: function (id, name) { return send("rename", { id: id, name: name }); },
        setTags: function (id, tags) { return send("setTags", { id: id, tags: tags }); },
        setCategory: function (id, category) { return send("setCategory", { id: id, category: category }); },
        batchMove: function (ids, category) { return send("batchMove", { ids: ids, category: category }); },
        reorder: function (orderedIds) { return send("reorder", { orderedIds: orderedIds }); },
        openFolder: function (id) { return send("openFolder", { id: id }); },
        getSettings: function () { return send("getSettings"); },
        saveSettings: function (patch) { return send("saveSettings", { patch: patch }); },
        undo: function () { return send("undo", {}); },
        scanDuplicates: function () { return send("scanDuplicates", {}); },
        renameTag: function (oldTag, newTag) { return send("renameTag", { oldTag: oldTag, newTag: newTag }); },
        deleteTag: function (tag) { return send("deleteTag", { tag: tag }); },
        exportPack: function (ids, saveAs) { return send("exportPack", { ids: ids, saveAs: saveAs }); },
        importPack: function (openFrom) { return send("importPack", { openFrom: openFrom }); },
        // 工具
        toolGetFonts: function () { return send("toolGetFonts", {}); },
        applyTextStyle: function (p) { return send("applyTextStyle", p || {}); },
        replaceFont: function (p) { return send("replaceFont", p || {}); },
        findReplace: function (p) { return send("findReplace", p || {}); },
        cleanText: function (p) { return send("cleanText", p || {}); },
        // 版式批处理工具
        requestFolder: function (title) { return send("requestFolder", { title: title || "" }); },
        requestFiles: function (title, filter) { return send("requestFiles", { title: title || "", filter: filter || "" }); },
        formatExportImages: function (p) { return send("formatExportImages", p || {}); },
        formatScanClean: function (p) { return send("formatScanClean", p || {}); },
        formatRunClean: function (p) { return send("formatRunClean", p || {}); },
        formatResize: function (p) { return send("formatResize", p || {}); },
        formatImageGrid: function (p) { return send("formatImageGrid", p || {}); },
        // 学术/学习工具
        academicFooter: function (p) { return send("academicFooter", p || {}); },
        academicExportPdf: function (p) { return send("academicExportPdf", p || {}); },
        academicExportOutline: function (p) { return send("academicExportOutline", p || {}); },
        academicBeautifyChart: function (p) { return send("academicBeautifyChart", p || {}); },
        academicBuildRefs: function (p) { return send("academicBuildRefs", p || {}); },
        log: function (level, message) { return send("log", { level: level, message: message }); }
    };
})();
