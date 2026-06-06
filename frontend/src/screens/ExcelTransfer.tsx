import { useEffect, useState } from "react";
import { ArrowLeftRight, Wand2, Play, FolderOpen, Database } from "lucide-react";
import { PathPicker } from "../components/PathPicker";
import { Card, Field, SectionTitle, Empty } from "../components/ui";
import { api } from "../lib/api";
import { useSettings } from "../lib/store";

const MODES = [
  { id: "append", label: "附加", desc: "寫到既有資料下方" },
  { id: "overwrite", label: "覆寫", desc: "清空標題列以下再寫" },
  { id: "fresh", label: "全新", desc: "新建目標檔" },
];

export default function ExcelTransfer() {
  const { settings, set } = useSettings();
  const [srcCols, setSrcCols] = useState<string[]>([]);
  const [tgtCols, setTgtCols] = useState<string[]>([]);
  const [srcSheets, setSrcSheets] = useState<string[]>([]);
  const [tgtSheets, setTgtSheets] = useState<string[]>([]);
  const [result, setResult] = useState<any>(null);
  const [busy, setBusy] = useState(false);

  const map: Record<string, string> = settings.transfer_column_map || {};
  const up = (patch: any) => set(patch);

  useEffect(() => {
    if (settings.excel_path) {
      api.excelSheets(settings.excel_path).then((r) => setSrcSheets(r.sheets || []));
      api.excelColumns(settings.excel_path, settings.sheet_name || "", settings.header_row || 1).then((r) => setSrcCols(r.columns || []));
    }
  }, [settings.excel_path, settings.sheet_name, settings.header_row]);

  useEffect(() => {
    if (settings.transfer_target_path) {
      api.excelSheets(settings.transfer_target_path).then((r) => setTgtSheets(r.sheets || []));
      api.excelColumns(settings.transfer_target_path, settings.transfer_target_sheet || "", settings.transfer_target_header_row || 1).then((r) => setTgtCols(r.columns || []));
    }
  }, [settings.transfer_target_path, settings.transfer_target_sheet, settings.transfer_target_header_row]);

  const autoMatch = async () => { const r = await api.transferAutoMatch(); await set({ transfer_column_map: r.value || {} }); };
  const setMapping = (src: string, tgt: string) => {
    const m = { ...map };
    if (tgt) m[src] = tgt; else delete m[src];
    up({ transfer_column_map: m });
  };
  const run = async () => { setBusy(true); setResult(await api.transferRun()); setBusy(false); };

  return (
    <div>
      <h1 className="mb-1 text-xl font-bold">Excel 搬移</h1>
      <p className="mb-5 text-sm text-slate-500 dark:text-slate-400">把來源 Excel 的欄位,依對應關係搬到目標 Excel。</p>

      <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
        <Card>
          <SectionTitle icon={<Database size={18} />} title="來源" />
          <Field label="來源 Excel"><PathPicker kind="excel" value={settings.excel_path} onPick={(p) => up({ excel_path: p })} /></Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="工作表"><select className="input" value={settings.sheet_name || ""} onChange={(e) => up({ sheet_name: e.target.value })}><option value="">（第一個）</option>{srcSheets.map((s) => <option key={s}>{s}</option>)}</select></Field>
            <Field label="標題列"><input type="number" min={1} className="input" value={settings.header_row || 1} onChange={(e) => up({ header_row: Number(e.target.value) || 1 })} /></Field>
          </div>
        </Card>
        <Card>
          <SectionTitle icon={<FolderOpen size={18} />} title="目標" />
          <Field label="目標 Excel"><PathPicker kind="save_excel" value={settings.transfer_target_path} onPick={(p) => up({ transfer_target_path: p })} /></Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="工作表"><select className="input" value={settings.transfer_target_sheet || ""} onChange={(e) => up({ transfer_target_sheet: e.target.value })}><option value="">（第一個 / 新建）</option>{tgtSheets.map((s) => <option key={s}>{s}</option>)}</select></Field>
            <Field label="標題列"><input type="number" min={1} className="input" value={settings.transfer_target_header_row || 1} onChange={(e) => up({ transfer_target_header_row: Number(e.target.value) || 1 })} /></Field>
          </div>
        </Card>
      </div>

      <Card className="mt-5">
        <div className="mb-3 flex items-center justify-between">
          <SectionTitle icon={<ArrowLeftRight size={18} />} title="欄位對應" desc="來源欄位 → 目標欄位" />
          <button className="btn-outline" onClick={autoMatch}><Wand2 size={16} /> 自動對應同名</button>
        </div>
        {srcCols.length === 0 ? <Empty icon={<Database size={26} />} text="先選好來源 Excel" /> : (
          <div className="grid grid-cols-1 gap-2 md:grid-cols-2">
            {srcCols.map((c) => (
              <div key={c} className="flex items-center gap-2 rounded-lg border border-slate-200 px-3 py-2 dark:border-slate-800">
                <span className="w-1/3 truncate text-sm font-medium" title={c}>{c}</span>
                <ArrowLeftRight size={14} className="shrink-0 text-slate-400" />
                <input list="tgtcols" className="input !py-1 text-xs" placeholder="（不搬）" value={map[c] || ""} onChange={(e) => setMapping(c, e.target.value)} />
              </div>
            ))}
            <datalist id="tgtcols">{tgtCols.map((t) => <option key={t} value={t} />)}</datalist>
          </div>
        )}
      </Card>

      <Card className="mt-5">
        <SectionTitle title="寫入模式" />
        <div className="grid grid-cols-3 gap-3">
          {MODES.map((m) => (
            <button key={m.id} onClick={() => up({ transfer_mode: m.id })}
              className={`rounded-lg border p-3 text-left transition-colors cursor-pointer ${settings.transfer_mode === m.id ? "border-brand-500 bg-brand-50 dark:bg-brand-900/20" : "border-slate-200 hover:border-brand-300 dark:border-slate-800"}`}>
              <div className="text-sm font-semibold">{m.label}</div>
              <div className="text-xs text-slate-400">{m.desc}</div>
            </button>
          ))}
        </div>
        <button className="btn-primary mt-4" onClick={run} disabled={busy}><Play size={16} /> {busy ? "搬移中…" : "開始搬移"}</button>
        {result && (
          <div className={`mt-4 rounded-lg p-3 text-sm ${result.error ? "bg-rose-50 text-rose-700 dark:bg-rose-900/30 dark:text-rose-300" : "bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300"}`}>
            {result.error ? <>失敗:{result.error}</> : <>完成!寫入 {result.rows_written}/{result.total} 列 → {result.sheet}（{result.mode}）</>}
          </div>
        )}
      </Card>
    </div>
  );
}
