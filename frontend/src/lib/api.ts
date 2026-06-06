// 後端 API 封裝。dev 透過 vite proxy,prod 同源。

export type Settings = Record<string, any>;

async function post<T = any>(url: string, body?: any): Promise<T> {
  const r = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: body ? JSON.stringify(body) : undefined,
  });
  return r.json();
}
async function get<T = any>(url: string): Promise<T> {
  const r = await fetch(url);
  return r.json();
}

export const api = {
  getSettings: () => get<Settings>("/api/settings"),
  updateSettings: (patch: Settings) => post("/api/settings", { patch }),

  pick: (kind: string) => post<{ path: string; cancelled: boolean }>("/api/pick", { kind }),

  excelSheets: (path: string) => post<{ sheets: string[]; error?: string }>("/api/excel/sheets", { path }),
  excelColumns: (path: string, sheet = "", header_row = 1) =>
    post<{ columns: string[]; rows: string[][]; total_rows: number; error?: string }>(
      "/api/excel/columns", { path, sheet, header_row }),

  wordVars: (path: string) => post<{ variables: string[]; error?: string }>("/api/template/word-vars", { path }),
  wordText: (path: string) => post<{ paragraphs: string[]; count: number; error?: string }>("/api/template/word-text", { path }),
  excelVars: (path: string, sheet = "") => post<{ variables: string[]; error?: string }>("/api/template/excel-vars", { path, sheet }),
  pptxVars: (path: string) => post<{ variables: string[]; error?: string }>("/api/template/pptx-vars", { path }),
  excelGrid: (path: string, sheet = "") =>
    post<{ cells: string[][]; col_labels: string[]; sheet: string; sheets: string[]; error?: string }>(
      "/api/template/excel-grid", { path, sheet }),

  insertTag: (anchor: string, variable: string, position = "after") =>
    post("/api/template/insert-tag", { anchor, variable, position }),
  renameTag: (oldName: string, newName: string) => post("/api/template/rename-tag", { old: oldName, new: newName }),
  excelInsertTag: (cell: string, tag: string) => post("/api/template/excel-insert-tag", { cell, tag }),
  suggestMappings: () => post("/api/template/suggest-mappings"),

  validate: (mode: "word" | "excel" | "pptx") => post<{ missing_in_excel: string[]; extra_in_excel: string[]; passed: boolean; error?: string }>("/api/validate", { mode }),

  generateCancel: () => post("/api/generate/cancel"),
  openOutput: () => post("/api/open-output"),

  transferAutoMatch: () => post<{ value: Record<string, string>; count: number }>("/api/transfer/auto-match"),
  transferRun: () => post("/api/transfer/run"),

  aiModels: () => get<{ available: boolean; text_models: string[]; vision_models: string[]; error?: string }>("/api/ai/models"),
  aiBudget: () => get<{ planner_used: number; planner_limit: number; reviewer_used: number; reviewer_limit: number }>("/api/ai/budget"),
  aiBudgetReset: () => post("/api/ai/budget/reset"),

  renderDocx: (path: string, dpi = 120, max_pages = 4) => post("/api/render/docx", { path, dpi, max_pages }),
  fileUrl: (path: string) => `/api/file?path=${encodeURIComponent(path)}`,

  // 圖片資料夾 → 欄位
  imagesList: (folder: string) => post<{ images: { name: string; path: string }[] }>("/api/images/list", { folder }),
  imagesMatch: (image_names: string[], targets: string[], use_ai = false, hint = "") =>
    post<{ mapping: Record<string, string>; ai?: boolean; error?: string }>("/api/images/match", { image_names, targets, use_ai, hint }),
  imagesApplyStatic: (mapping: Record<string, string>, width_mm = 80) =>
    post<{ applied: any[]; failed: any[]; applied_count: number; error?: string }>("/api/images/apply-static", { mapping, width_mm }),
  imagesFillExcel: (folder: string, key_column: string, image_column: string, image_to_key?: Record<string, string>) =>
    post<{ output_path: string; matched: number; total: number; unmatched_rows: string[]; error?: string }>(
      "/api/images/fill-excel", { folder, key_column, image_column, image_to_key }),
  imagesAutoVisual: (folder: string, width_mm = 0, use_com = true) =>
    post<any>("/api/images/auto-visual", { folder, width_mm, use_com }),
  reportFill: (template_path: string, photo_root: string, output_path = "", dry_run = false, match_mode = "auto") =>
    post<any>("/api/report/fill-from-folders", { template_path, photo_root, output_path, dry_run, match_mode }),
};

// SSE 產出進度
export function streamGenerate(
  mode: "word" | "excel" | "pptx",
  onEvent: (e: any) => void,
  onDone: () => void,
) {
  const es = new EventSource(`/api/generate/stream?mode=${mode}`);
  es.onmessage = (ev) => {
    try {
      const data = JSON.parse(ev.data);
      onEvent(data);
      if (data.type === "done" || data.type === "error") {
        es.close();
        onDone();
      }
    } catch {}
  };
  es.onerror = () => { es.close(); onDone(); };
  return () => es.close();
}
