// C#↔JS 通信层（TS 化，协议与旧 bridge.js 完全一致）
let reqSeq = 0;
const pending = new Map<string, { resolve: (d: any) => void; reject: (e: Error) => void; timer: number }>();
const listeners: Record<string, Array<(payload: any) => void>> = {};
const TIMEOUT = 30000; // 默认请求超时 30s：首次插入/导出等 COM 冷启动可能超过 8s，过短会误报 timeout

function post(msg: object): boolean {
  const wv = (window as any).chrome?.webview;
  try {
    wv.postMessage(JSON.stringify(msg));
    return true;
  } catch (e) {
    console.error("postMessage failed (not hosted in WebView2?)", e);
    return false;
  }
}

function send(type: string, payload: any = {}, timeoutMs: number = TIMEOUT): Promise<any> {
  return new Promise((resolve, reject) => {
    const reqId = "r" + ++reqSeq;
    pending.set(reqId, {
      resolve,
      reject,
      timer:
        timeoutMs > 0
          ? window.setTimeout(() => {
              if (pending.has(reqId)) {
                pending.delete(reqId);
                reject(new Error("timeout: " + type));
              }
            }, timeoutMs)
          : 0, // timeoutMs<=0: 不超时（模态对话框类：用户在系统对话框停留多久都等）
    });
    if (!post({ type, reqId, payload })) {
      const p = pending.get(reqId);
      if (p && p.timer) clearTimeout(p.timer);
      pending.delete(reqId);
      reject(new Error("no host"));
    }
  });
}

function on(type: string, fn: (payload: any) => void) {
  (listeners[type] = listeners[type] || []).push(fn);
}

const wv = (window as any).chrome?.webview;
if (wv?.addEventListener) {
  wv.addEventListener("message", (e: any) => {
    let msg: any;
    try {
      msg = typeof e.data === "string" ? JSON.parse(e.data) : e.data;
    } catch {
      return;
    }
    if (!msg || !msg.type) return;
    if (msg.type === "result") {
      const p = pending.get(msg.reqId);
      if (p) {
        clearTimeout(p.timer);
        pending.delete(msg.reqId);
        if (msg.ok) p.resolve(msg.data || {});
        else p.reject(new Error(msg.error || "unknown error"));
      }
      return;
    }
    (listeners[msg.type] || []).forEach((fn) => {
      try {
        fn(msg.payload || {});
      } catch (err) {
        console.error(err);
      }
    });
  });
}

window.addEventListener("error", (e) => {
  try {
    post({ type: "log", payload: { level: "error", message: "JS uncaught: " + e.message } });
  } catch {
    /* noop */
  }
});

export const bridge = {
  send,
  on,
  init: () => send("init"),
  list: (filter: string, category: string, kind: string, group: string, sort: string, offset: number, limit: number) =>
    send("list", { filter, category, kind, group, sort, offset, limit }),
  insert: (id: string) => send("insert", { id }),
  insertPage: (id: string) => send("insertPage", { id }),
  applyStyle: (id: string) => send("applyStyle", { id }),
  applyStyleSameType: (id: string) => send("applyStyleSameType", { id }),
  applyStyleSameTypeAllSlides: (id: string) => send("applyStyleSameTypeAllSlides", { id }),
  adjust: (action: string) => send("adjust", { action }),
  del: (ids: string[]) => send("delete", { ids }),
  rename: (id: string, name: string) => send("rename", { id, name }),
  setCategory: (id: string, category: string) => send("setCategory", { id, category }),
  batchMove: (ids: string[], category: string) => send("batchMove", { ids, category }),
  setGroup: (id: string, group: string) => send("setGroup", { id, group }),
  batchMoveGroup: (ids: string[], group: string) => send("batchMoveGroup", { ids, group }),
  createGroup: (name: string) => send("createGroup", { name }),
  renameGroup: (oldName: string, newName: string) => send("renameGroup", { oldName, newName }),
  deleteGroup: (name: string) => send("deleteGroup", { name }),
  reorder: (orderedIds: string[]) => send("reorder", { orderedIds }),
  openFolder: (id: string) => send("openFolder", { id }),
  getSettings: () => send("getSettings"),
  saveSettings: (patch: object) => send("saveSettings", { patch }),
  undo: () => send("undo", {}),
  redo: () => send("redo", {}),
  trashList: () => send("trashList", {}),
  trashRestore: (id: string) => send("trashRestore", { id }),
  trashPurge: () => send("trashPurge", {}),
  trashPurgeOne: (id: string) => send("trashPurgeOne", { id }),
  rebuildPreviews: () => send("rebuildPreviews", {}, 120000), // 逐个打开文稿重建，可能较慢
  setPaneDock: (dock: string) => send("setPaneDock", { dock }),
  scanDuplicates: () => send("scanDuplicates", {}),
  createGroupCategory: (group: string, category: string) => send("createGroupCategory", { group, category }),
  renameGroupCategory: (group: string, oldCategory: string, newCategory: string) => send("renameGroupCategory", { group, oldCategory, newCategory }),
  deleteGroupCategory: (group: string, category: string) => send("deleteGroupCategory", { group, category }),
  exportPack: (ids: string[], saveAs: string) => send("exportPack", { ids, saveAs }, 60000),
  importPack: (openFrom: string) => send("importPack", { openFrom }, 60000),
  importPackFile: () => send("importPackFile", {}, 0), // 弹文件选择器选 zip（模态，不超时）
  importPictures: () => send("importPictures", {}, 0), // 模态选文件对话框，不设超时
  // 工具
  toolGetFonts: () => send("toolGetFonts", {}),
  applyTextStyle: (p: object) => send("applyTextStyle", p),
  replaceFont: (p: object) => send("replaceFont", p),
  findReplace: (p: object) => send("findReplace", p),
  cleanText: (p: object) => send("cleanText", p),
  requestFolder: (title: string) => send("requestFolder", { title }, 0), // 模态目录对话框：不设超时
  requestFiles: (title: string, filter: string) => send("requestFiles", { title, filter }, 0),
  formatExportImages: (p: object) => send("formatExportImages", p),
  formatScanClean: (p: object) => send("formatScanClean", p),
  formatRunClean: (p: object) => send("formatRunClean", p),
  formatResize: (p: object) => send("formatResize", p),
  formatImageGrid: (p: object) => send("formatImageGrid", p),
  academicFooter: (p: object) => send("academicFooter", p),
  academicExportPdf: (p: object) => send("academicExportPdf", p),
  academicExportOutline: (p: object) => send("academicExportOutline", p),
  academicBeautifyChart: (p: object) => send("academicBeautifyChart", p),
  insertTemplatePack: (p: object) => send("insertTemplatePack", p),
  log: (level: string, message: string) => send("log", { level, message }),
};

export type ListItem = {
  id: string;
  name: string;
  kind: "Shape" | "Style";
  category: string;
  group: string;
  shapeCount?: number;
  useCount?: number;
  previewUrl: string;
  /** true=整页快照（含背景版式，插入以整页为单位） */
  snapshotPage?: boolean;
};

export type InitData = {
  theme: string;
  cardSize: string;
  groups: string[];
  groupCounts: Record<string, number>;
  groupCategories: Record<string, string[]>;
  categories: string[];
  total: number;
  recentTop: ListItem[];
  [k: string]: unknown;
};
