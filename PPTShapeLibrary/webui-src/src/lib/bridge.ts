// C#↔JS 通信层（TS 化，协议与旧 bridge.js 完全一致）
let reqSeq = 0;
const pending = new Map<string, { resolve: (d: any) => void; reject: (e: Error) => void; timer: number }>();
const listeners: Record<string, Array<(payload: any) => void>> = {};
const TIMEOUT = 8000;

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

function send(type: string, payload: any = {}): Promise<any> {
  return new Promise((resolve, reject) => {
    const reqId = "r" + ++reqSeq;
    pending.set(reqId, {
      resolve,
      reject,
      timer: window.setTimeout(() => {
        if (pending.has(reqId)) {
          pending.delete(reqId);
          reject(new Error("timeout: " + type));
        }
      }, TIMEOUT),
    });
    if (!post({ type, reqId, payload })) {
      clearTimeout(pending.get(reqId)!.timer);
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
  list: (filter: string, category: string, kind: string, tag: string, group: string, sort: string, offset: number, limit: number) =>
    send("list", { filter, category, kind, tag, group, sort, offset, limit }),
  insert: (id: string) => send("insert", { id }),
  applyStyle: (id: string) => send("applyStyle", { id }),
  applyStyleSameType: (id: string) => send("applyStyleSameType", { id }),
  applyStyleSameTypeAllSlides: (id: string) => send("applyStyleSameTypeAllSlides", { id }),
  adjust: (action: string) => send("adjust", { action }),
  del: (ids: string[]) => send("delete", { ids }),
  rename: (id: string, name: string) => send("rename", { id, name }),
  setTags: (id: string, tags: string) => send("setTags", { id, tags }),
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
  scanDuplicates: () => send("scanDuplicates", {}),
  renameTag: (oldTag: string, newTag: string) => send("renameTag", { oldTag, newTag }),
  deleteTag: (tag: string) => send("deleteTag", { tag }),
  exportPack: (ids: string[], saveAs: string) => send("exportPack", { ids, saveAs }),
  importPack: (openFrom: string) => send("importPack", { openFrom }),
  importPictures: () => send("importPictures", {}),
  // 工具
  toolGetFonts: () => send("toolGetFonts", {}),
  applyTextStyle: (p: object) => send("applyTextStyle", p),
  replaceFont: (p: object) => send("replaceFont", p),
  findReplace: (p: object) => send("findReplace", p),
  cleanText: (p: object) => send("cleanText", p),
  requestFolder: (title: string) => send("requestFolder", { title }),
  requestFiles: (title: string, filter: string) => send("requestFiles", { title, filter }),
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
  tags: string[];
  shapeCount?: number;
  useCount?: number;
  previewUrl: string;
};

export type InitData = {
  theme: string;
  cardSize: string;
  groups: string[];
  groupCounts: Record<string, number>;
  groupCategories: Record<string, string[]>;
  tags: string[];
  tagCounts: Record<string, number>;
  categories: string[];
  total: number;
  recentTop: ListItem[];
  [k: string]: unknown;
};
