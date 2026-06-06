import { useCallback, useEffect, useMemo, useState } from "react";
import clsx from "clsx";
import {
  FileText, FileSpreadsheet, Database, FolderOutput, Tags, ShieldCheck,
  Rocket, Sparkles, Play, X, FolderOpen, CheckCircle2, AlertTriangle,
  ArrowRight, ArrowLeft, Wand2, Keyboard, MousePointerClick, Presentation, ImagePlus,
} from "lucide-react";
import { Stepper, type Step } from "../components/Stepper";
import { PathPicker } from "../components/PathPicker";
import { Card, Field, Badge, Empty, SectionTitle } from "../components/ui";
import { api, streamGenerate } from "../lib/api";
import { useSettings } from "../lib/store";

const STEPS: Step[] = [
  { key: "config", label: "設定來源", hint: "範本 + 資料" },
  { key: "map", label: "對應標籤", hint: "{{欄位}} 對應" },
  { key: "validate", label: "驗證", hint: "檢查對齊" },
  { key: "generate", label: "產出", hint: "批次生成" },
  { key: "review", label: "完成 / 審查", hint: "結果與檢查" },
];

type ReportType = "word" | "excel" | "pptx";

export default function ReportWizard() {
  const { settings, set } = useSettings();
  const [step, setStep] = useState(0);
  const [maxReached, setMaxReached] = useState(0);
  const [reportType, setReportType] = useState<ReportType>("word");

  const up = useCallback((patch: Record<string, any>) => set(patch), [set]);
  const go = (i: number) => { setStep(i); setMaxReached((m) => Math.max(m, i)); };
  const next = () => go(Math.min(step + 1, STEPS.length - 1));
  const prev = () => go(Math.max(step - 1, 0));

  return (
    <div>
      <div className="mb-1 flex items-center justify-between">
        <h1 className="text-xl font-bold">報告精靈</h1>
        <ReportTypeToggle value={reportType} onChange={setReportType} />
      </div>
      <p className="mb-5 text-sm text-slate-500 dark:text-slate-400">
        一份範本 + 一張 Excel,批次產出每一列一份報告。跟著下面 5 步走即可。
      </p>

      <Card className="mb-5">
        <Stepper steps={STEPS} current={step} onJump={go} maxReached={maxReached} />
      </Card>

      <div className="animate-fade-in">
        {step === 0 && <StepConfigure reportType={reportType} settings={settings} up={up} />}
        {step === 1 && <StepMap reportType={reportType} settings={settings} />}
        {step === 2 && <StepValidate reportType={reportType} />}
        {step === 3 && <StepGenerate reportType={reportType} settings={settings} />}
        {step === 4 && <StepReview settings={settings} />}
      </div>

      <div className="mt-6 flex items-center justify-between">
        <button className="btn-ghost" onClick={prev} disabled={step === 0}>
          <ArrowLeft size={16} /> 上一步
        </button>
        <span className="text-xs text-slate-400">步驟 {step + 1} / {STEPS.length}</span>
        <button className="btn-primary" onClick={next} disabled={step === STEPS.length - 1}>
          下一步 <ArrowRight size={16} />
        </button>
      </div>
    </div>
  );
}

function ReportTypeToggle({ value, onChange }: { value: ReportType; onChange: (v: ReportType) => void }) {
  return (
    <div className="flex rounded-lg border border-slate-200 bg-white p-1 dark:border-slate-700 dark:bg-slate-900">
      {([["word", "Word 報告", FileText], ["excel", "Excel 報告", FileSpreadsheet], ["pptx", "PPTX 簡報", Presentation]] as const).map(
        ([id, label, Icon]) => (
          <button
            key={id}
            onClick={() => onChange(id)}
            className={clsx(
              "flex items-center gap-1.5 rounded-md px-3 py-1.5 text-sm font-medium transition-colors cursor-pointer",
              value === id
                ? "bg-brand-600 text-white"
                : "text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
            )}
          >
            <Icon size={15} /> {label}
          </button>
        )
      )}
    </div>
  );
}

/* ---------------- Step 0: Configure ---------------- */
function StepConfigure({ reportType, settings, up }: any) {
  const [sheets, setSheets] = useState<string[]>([]);
  const [tplSheets, setTplSheets] = useState<string[]>([]);
  const [preview, setPreview] = useState<{ columns: string[]; rows: string[][]; total_rows: number } | null>(null);

  const dataPath = settings.excel_path;
  const headerRow = settings.header_row || 1;
  const dataSheet = settings.sheet_name || "";

  useEffect(() => {
    if (dataPath) api.excelSheets(dataPath).then((r) => setSheets(r.sheets || []));
  }, [dataPath]);
  useEffect(() => {
    if (dataPath) api.excelColumns(dataPath, dataSheet, headerRow).then((r) => !r.error && setPreview(r));
  }, [dataPath, dataSheet, headerRow]);
  useEffect(() => {
    if (reportType === "excel" && settings.excel_template_path)
      api.excelGrid(settings.excel_template_path).then((r) => setTplSheets(r.sheets || []));
  }, [reportType, settings.excel_template_path]);

  return (
   <>
    <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
      <Card>
        <SectionTitle icon={<FileText size={18} />} title="① 範本" desc="批次產出時要套用的版型,含 {{欄位}} 標籤" />
        {reportType === "word" && (
          <Field label="Word 範本 (.docx)">
            <PathPicker kind="word" value={settings.word_path} placeholder="選擇 .docx 範本" onPick={(p) => up({ word_path: p })} />
          </Field>
        )}
        {reportType === "excel" && (
          <>
            <Field label="Excel 範本 (.xlsx)">
              <PathPicker kind="excel" value={settings.excel_template_path} placeholder="選擇 .xlsx 範本" onPick={(p) => up({ excel_template_path: p })} />
            </Field>
            {tplSheets.length > 0 && (
              <Field label="範本工作表">
                <select className="input" value={settings.excel_template_sheet || ""} onChange={(e) => up({ excel_template_sheet: e.target.value })}>
                  <option value="">（第一個）</option>
                  {tplSheets.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </Field>
            )}
          </>
        )}
        {reportType === "pptx" && (
          <Field label="PPTX 範本 (.pptx)" hint="在 PowerPoint 的文字框/表格放入 {{欄位}} 標籤">
            <PathPicker kind="pptx" value={settings.pptx_template_path} placeholder="選擇 .pptx 範本" onPick={(p) => up({ pptx_template_path: p })} />
          </Field>
        )}
      </Card>

      <Card>
        <SectionTitle icon={<Database size={18} />} title="② 資料來源" desc="每一列會產出一份報告" />
        <Field label="來源 Excel (.xlsx)">
          <PathPicker kind="excel" value={settings.excel_path} placeholder="選擇資料 Excel" onPick={(p) => up({ excel_path: p })} />
        </Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="工作表">
            <select className="input" value={dataSheet} onChange={(e) => up({ sheet_name: e.target.value })}>
              <option value="">（第一個）</option>
              {sheets.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </Field>
          <Field label="標題列" hint="欄位名稱在第幾列">
            <input type="number" min={1} className="input" value={headerRow} onChange={(e) => up({ header_row: Number(e.target.value) || 1 })} />
          </Field>
        </div>
      </Card>

      <Card>
        <SectionTitle icon={<FolderOutput size={18} />} title="③ 輸出" desc="產出檔案的存放位置與命名" />
        <Field label="輸出資料夾">
          <PathPicker kind="directory" value={settings.output_dir} placeholder="Generated_Reports" onPick={(p) => up({ output_dir: p })} />
        </Field>
        <Field label="檔名規則" hint="可用 {index} 與任意欄位名,如 {客戶}_{index}">
          <input className="input font-mono text-xs"
            value={(reportType === "word" ? settings.filename_template : reportType === "excel" ? settings.excel_filename_template : settings.pptx_filename_template) || ""}
            onChange={(e) => up(reportType === "word" ? { filename_template: e.target.value } : reportType === "excel" ? { excel_filename_template: e.target.value } : { pptx_filename_template: e.target.value })} />
        </Field>
        {reportType === "word" && (
          <Field label="圖片寬度 (mm)" hint="Excel 欄位若是圖片路徑會自動嵌入">
            <input type="number" min={1} className="input w-32" value={settings.image_width_mm || 80} onChange={(e) => up({ image_width_mm: Number(e.target.value) || 80 })} />
          </Field>
        )}
      </Card>

      <Card>
        <SectionTitle icon={<Database size={18} />} title="資料預覽" desc={preview ? `共 ${preview.total_rows} 列 · ${preview.columns.length} 欄` : "選好來源 Excel 後顯示"} />
        {preview ? (
          <div className="overflow-auto rounded-lg border border-slate-200 dark:border-slate-800" style={{ maxHeight: 260 }}>
            <table className="w-full text-xs">
              <thead className="sticky top-0 bg-slate-50 dark:bg-slate-800">
                <tr>{preview.columns.map((c) => <th key={c} className="whitespace-nowrap px-2 py-1.5 text-left font-semibold">{c}</th>)}</tr>
              </thead>
              <tbody>
                {preview.rows.map((row, i) => (
                  <tr key={i} className="border-t border-slate-100 dark:border-slate-800">
                    {row.map((v, j) => <td key={j} className="whitespace-nowrap px-2 py-1 text-slate-600 dark:text-slate-400">{v}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : <Empty icon={<Database size={28} />} text="尚未選擇資料來源" />}
      </Card>
    </div>

    <PerRowImageCard up={up} columns={preview?.columns || []} reportType={reportType} />
   </>
  );
}

/* 選用:每列不同的圖片(把資料夾照片依檔名配到每列,寫進資料來源) */
function PerRowImageCard({ up, columns, reportType }: any) {
  const [open, setOpen] = useState(false);
  const [folder, setFolder] = useState("");
  const [imgs, setImgs] = useState<number>(0);
  const [keyCol, setKeyCol] = useState("");
  const [imgCol, setImgCol] = useState("");
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<any>(null);

  useEffect(() => { if (folder) api.imagesList(folder).then((r) => setImgs((r.images || []).length)); }, [folder]);

  const run = async () => {
    if (!folder) return alert("請先選圖片資料夾");
    if (!keyCol) return alert("請選 key 欄位(檔名要對到的欄)");
    if (!imgCol) return alert("請填圖片欄位名(範本裡的圖片標籤)");
    setBusy(true);
    const r = await api.imagesFillExcel(folder, keyCol, imgCol);
    setBusy(false); setResult(r);
    if (r.output_path) up({ excel_path: r.output_path }); // 資料來源切到含圖片路徑的副本
  };

  return (
    <Card className="mt-5">
      <button onClick={() => setOpen((o) => !o)} className="flex w-full items-center justify-between text-left">
        <SectionTitle icon={<ImagePlus size={18} />} title="每列圖片(選用)"
          desc="若每份報告要放各自的照片(例:每個客戶一張店面照),在這裡設定;不需要就略過。" />
        <span className="text-sm text-slate-400">{open ? "▲ 收合" : "▼ 展開"}</span>
      </button>
      {open && (
        <div className="mt-3 space-y-3 border-t border-slate-100 pt-4 dark:border-slate-800">
          <Field label="① 圖片資料夾" hint="照片檔名要對得上下方 key 欄位的值(例:檔名「宏達科技.jpg」對到客戶名稱=宏達科技)">
            <PathPicker kind="directory" value={folder} placeholder="選含照片的資料夾" onPick={setFolder} />
            {folder && <p className="mt-1 text-xs text-slate-400">偵測到 {imgs} 張圖片</p>}
          </Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="② key 欄位(檔名對到的欄)">
              <select className="input" value={keyCol} onChange={(e) => setKeyCol(e.target.value)}>
                <option value="">（請選)</option>
                {columns.map((c: string) => <option key={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="③ 圖片欄位(範本裡的圖片標籤名)" hint="如 店面照,要跟範本 {{店面照}} 一致">
              <input className="input" value={imgCol} onChange={(e) => setImgCol(e.target.value)} placeholder="例:店面照" />
            </Field>
          </div>
          <button className="btn-outline" onClick={run} disabled={busy || !folder}>
            <Wand2 size={15} /> {busy ? "比對寫入中…" : "比對並寫入(更新資料來源)"}
          </button>
          {result && (
            <div className={clsx("rounded-lg p-3 text-sm", result.error ? "bg-rose-50 text-rose-700 dark:bg-rose-900/30 dark:text-rose-300" : "bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300")}>
              {result.error ? <>失敗:{result.error}</> : (
                <>對應 {result.matched}/{result.total} 列,資料來源已更新。{reportType === "word" ? "記得在 Word 範本放 " : reportType === "pptx" ? "記得在 PPTX 範本放 " : "記得在 Excel 範本放 "}<span className="font-mono">{`{{${imgCol}}}`}</span> 才會嵌圖。
                {result.unmatched_rows?.length > 0 && <div className="mt-1 text-amber-700 dark:text-amber-300">未對到:{result.unmatched_rows.join("、")}</div>}</>
              )}
            </div>
          )}
        </div>
      )}
    </Card>
  );
}

/* ---------------- Step 1: Map tags ---------------- */
function StepMap({ reportType, settings }: any) {
  const [columns, setColumns] = useState<string[]>([]);
  const [picked, setPicked] = useState<string | null>(null);
  const [vars, setVars] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);

  const refreshVars = useCallback(() => {
    if (reportType === "word" && settings.word_path)
      api.wordVars(settings.word_path).then((r) => setVars(r.variables || []));
    else if (reportType === "excel" && settings.excel_template_path)
      api.excelVars(settings.excel_template_path, settings.excel_template_sheet || "").then((r) => setVars(r.variables || []));
    else if (reportType === "pptx" && settings.pptx_template_path)
      api.pptxVars(settings.pptx_template_path).then((r) => setVars(r.variables || []));
  }, [reportType, settings.word_path, settings.excel_template_path, settings.excel_template_sheet, settings.pptx_template_path]);

  useEffect(() => {
    if (settings.excel_path) api.excelColumns(settings.excel_path, settings.sheet_name || "", settings.header_row || 1).then((r) => setColumns(r.columns || []));
  }, [settings.excel_path, settings.sheet_name, settings.header_row]);
  useEffect(() => { refreshVars(); }, [refreshVars]);

  const mapped = useMemo(() => new Set(vars), [vars]);

  return (
    <div className="grid grid-cols-1 gap-5 lg:grid-cols-3">
      {/* 欄位調色盤 */}
      <Card className="lg:col-span-1">
        <SectionTitle icon={<Tags size={18} />} title="Excel 欄位" desc="① 點一個欄位 → ② 點右側位置插入" />
        {columns.length === 0 ? (
          <Empty icon={<Database size={26} />} text="先在上一步選好來源 Excel" />
        ) : (
          <div className="flex flex-wrap gap-2">
            {columns.map((c) => (
              <button key={c} onClick={() => setPicked(c)}
                className={clsx("chip border transition-colors cursor-pointer",
                  picked === c ? "border-brand-500 bg-brand-600 text-white"
                    : mapped.has(c) ? "border-emerald-300 bg-emerald-50 text-emerald-700 dark:border-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300"
                      : "border-slate-300 bg-white text-slate-700 hover:border-brand-400 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300")}>
                {mapped.has(c) && <CheckCircle2 size={12} />} {c}
              </button>
            ))}
          </div>
        )}
        <div className="mt-4 rounded-lg bg-slate-50 p-3 text-xs text-slate-500 dark:bg-slate-800/50 dark:text-slate-400">
          <div className="mb-1 flex items-center gap-1.5 font-semibold text-slate-600 dark:text-slate-300"><MousePointerClick size={13} /> 目前選取</div>
          {picked ? <Badge tone="brand">{`{{ ${picked} }}`}</Badge> : "（尚未選欄位）"}
        </div>
        {reportType === "word" && (
          <button className="btn-outline mt-3 w-full" disabled={busy}
            onClick={async () => { setBusy(true); const r: any = await api.suggestMappings(); setBusy(false); refreshVars(); alert(r?.error ? `AI 建議失敗:${r.error}` : `AI 建議:改名 ${r?.renames?.length || 0} 筆、插入 ${r?.inserts?.length || 0} 筆。\n請於下方檢視範本變數。`); }}>
            <Wand2 size={16} /> {busy ? "AI 分析中…" : "AI 建議對應"}
          </button>
        )}
        <div className="mt-3 flex items-start gap-2 rounded-lg border border-amber-200 bg-amber-50 p-2.5 text-xs text-amber-800 dark:border-amber-900/50 dark:bg-amber-900/20 dark:text-amber-300">
          <Keyboard size={14} className="mt-0.5 shrink-0" />
          <span>進階:在真正的 {reportType === "word" ? "Word" : "Excel"} 視窗中按 <b>Ctrl+Shift+M</b> 也能即時插入標籤(需啟動熱鍵橋接)。</span>
        </div>
      </Card>

      {/* 範本編輯區 */}
      <Card className="lg:col-span-2">
        {reportType === "word" && <WordTagEditor settings={settings} picked={picked} onInserted={refreshVars} mapped={mapped} />}
        {reportType === "excel" && <ExcelTagEditor settings={settings} picked={picked} onInserted={refreshVars} />}
        {reportType === "pptx" && <PptxTagPanel settings={settings} vars={vars} columns={columns} mapped={mapped} />}
      </Card>
    </div>
  );
}

function WordTagEditor({ settings, picked, onInserted }: any) {
  const [paragraphs, setParagraphs] = useState<string[]>([]);
  const [hoverIdx, setHoverIdx] = useState<number | null>(null);

  useEffect(() => {
    if (settings.word_path) api.wordText(settings.word_path).then((r) => setParagraphs(r.paragraphs || []));
  }, [settings.word_path]);

  async function insertAfter(text: string) {
    if (!picked) { alert("請先點左側一個 Excel 欄位"); return; }
    const r: any = await api.insertTag(text, picked, "after");
    if (r?.error) alert(`插入失敗:${r.error}`);
    else { onInserted(); if (settings.word_path) api.wordText(settings.word_path).then((rr) => setParagraphs(rr.paragraphs || [])); }
  }

  if (!settings.word_path) return <Empty icon={<FileText size={28} />} text="先在上一步選好 Word 範本" />;

  return (
    <>
      <SectionTitle icon={<FileText size={18} />} title="Word 範本內容" desc="點某段落的「＋ 插入」把選取的欄位標籤放到該段之後" />
      <div className="space-y-1 overflow-auto rounded-lg border border-slate-200 p-2 dark:border-slate-800" style={{ maxHeight: 420 }}>
        {paragraphs.length === 0 && <Empty text="範本沒有可顯示的文字段落" />}
        {paragraphs.map((p, i) => (
          <div key={i} onMouseEnter={() => setHoverIdx(i)} onMouseLeave={() => setHoverIdx(null)}
            className="group flex items-center gap-2 rounded-md px-2 py-1.5 hover:bg-brand-50 dark:hover:bg-brand-900/20">
            <span className="flex-1 text-sm text-slate-700 dark:text-slate-300">{p || <span className="text-slate-300">（空白段）</span>}</span>
            <button onClick={() => insertAfter(p)}
              className={clsx("btn-primary !px-2 !py-1 text-xs transition-opacity", hoverIdx === i ? "opacity-100" : "opacity-0")}>
              ＋ 插入
            </button>
          </div>
        ))}
      </div>
    </>
  );
}

function ExcelTagEditor({ settings, picked, onInserted }: any) {
  const [grid, setGrid] = useState<{ cells: string[][]; col_labels: string[] } | null>(null);
  const [sel, setSel] = useState<string | null>(null);

  const reload = useCallback(() => {
    if (settings.excel_template_path)
      api.excelGrid(settings.excel_template_path, settings.excel_template_sheet || "").then((r) => !r.error && setGrid(r));
  }, [settings.excel_template_path, settings.excel_template_sheet]);
  useEffect(() => { reload(); }, [reload]);

  async function insertInto(coord: string) {
    if (!picked) { alert("請先點左側一個 Excel 欄位"); return; }
    const r: any = await api.excelInsertTag(coord, picked);
    if (r?.error) alert(`插入失敗:${r.error}`);
    else { setSel(coord); onInserted(); reload(); }
  }

  if (!settings.excel_template_path) return <Empty icon={<FileSpreadsheet size={28} />} text="先在上一步選好 Excel 範本" />;
  if (!grid) return <Empty text="載入範本中…" />;

  return (
    <>
      <SectionTitle icon={<FileSpreadsheet size={18} />} title="Excel 範本網格" desc="點某個儲存格把選取的欄位標籤插入該格" />
      <div className="overflow-auto rounded-lg border border-slate-200 dark:border-slate-800" style={{ maxHeight: 420 }}>
        <table className="border-collapse text-xs">
          <thead>
            <tr>
              <th className="sticky left-0 top-0 z-10 bg-slate-100 px-2 py-1 dark:bg-slate-800"></th>
              {grid.col_labels.map((c) => <th key={c} className="bg-slate-100 px-3 py-1 font-semibold dark:bg-slate-800">{c}</th>)}
            </tr>
          </thead>
          <tbody>
            {grid.cells.map((row, r) => (
              <tr key={r}>
                <td className="sticky left-0 bg-slate-100 px-2 py-1 text-center font-semibold text-slate-500 dark:bg-slate-800">{r + 1}</td>
                {row.map((v, c) => {
                  const coord = `${grid.col_labels[c]}${r + 1}`;
                  const hasTag = /\{\{.*\}\}/.test(v);
                  return (
                    <td key={c} onClick={() => insertInto(coord)}
                      className={clsx("min-w-[90px] cursor-pointer border border-slate-200 px-2 py-1 transition-colors hover:bg-brand-50 dark:border-slate-800 dark:hover:bg-brand-900/20",
                        sel === coord && "ring-2 ring-brand-500",
                        hasTag && "bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300")}>
                      {v}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

function PptxTagPanel({ settings, vars, columns, mapped }: any) {
  if (!settings.pptx_template_path) return <Empty icon={<Presentation size={28} />} text="先在上一步選好 PPTX 範本" />;
  const colSet = new Set(columns);
  return (
    <>
      <SectionTitle icon={<Presentation size={18} />} title="PPTX 範本標籤" desc="PPTX 標籤請在 PowerPoint 編輯;這裡核對標籤與 Excel 欄位是否對得上" />
      <div className="mb-4 flex items-start gap-2 rounded-lg border border-amber-200 bg-amber-50 p-2.5 text-xs text-amber-800 dark:border-amber-900/50 dark:bg-amber-900/20 dark:text-amber-300">
        <Keyboard size={14} className="mt-0.5 shrink-0" />
        <span>用 PowerPoint 打開 .pptx,在文字框 / 表格輸入 <span className="font-mono">{`{{欄位名}}`}</span>(欄位名要對到 Excel 欄位),存檔後回來驗證。圖片欄位:把某形狀文字設成單一 <span className="font-mono">{`{{圖片欄}}`}</span>,值是圖片路徑就會就地嵌圖。</span>
      </div>
      <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
        <div>
          <div className="label">範本目前的標籤(共 {vars.length})</div>
          {vars.length === 0 ? <p className="text-sm text-slate-400">尚未偵測到 {`{{標籤}}`}</p> :
            <div className="flex flex-wrap gap-1.5">{vars.map((v: string) =>
              <Badge key={v} tone={colSet.has(v) ? "green" : "red"}>{colSet.has(v) ? <CheckCircle2 size={12} /> : <AlertTriangle size={12} />} {v}</Badge>)}</div>}
        </div>
        <div>
          <div className="label">Excel 欄位(共 {columns.length})</div>
          <div className="flex flex-wrap gap-1.5">{columns.map((c: string) =>
            <Badge key={c} tone={mapped.has(c) ? "brand" : "slate"}>{c}</Badge>)}</div>
        </div>
      </div>
    </>
  );
}

/* ---------------- Step 2: Validate ---------------- */
function StepValidate({ reportType }: any) {
  const [res, setRes] = useState<any>(null);
  const [busy, setBusy] = useState(false);

  const run = async () => { setBusy(true); setRes(await api.validate(reportType)); setBusy(false); };
  useEffect(() => { run(); /* eslint-disable-next-line */ }, []);

  return (
    <Card>
      <SectionTitle icon={<ShieldCheck size={18} />} title="驗證範本與資料對齊" desc="檢查範本用到的標籤是否都能在 Excel 找到對應欄位" />
      <button className="btn-primary mb-4" onClick={run} disabled={busy}>{busy ? "檢查中…" : "重新檢查"}</button>
      {res && (res.error ? (
        <div className="flex items-center gap-2 rounded-lg bg-rose-50 p-3 text-sm text-rose-700 dark:bg-rose-900/30 dark:text-rose-300">
          <AlertTriangle size={16} /> {res.error}
        </div>
      ) : (
        <div className="space-y-4">
          <div className={clsx("flex items-center gap-2 rounded-lg p-3 text-sm font-medium",
            res.passed ? "bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300"
              : "bg-amber-50 text-amber-700 dark:bg-amber-900/30 dark:text-amber-300")}>
            {res.passed ? <CheckCircle2 size={18} /> : <AlertTriangle size={18} />}
            {res.passed ? "全部對齊,可以產出!" : `有 ${res.missing_in_excel.length} 個標籤在 Excel 找不到對應欄位`}
          </div>
          <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
            <div>
              <div className="label">範本有、Excel 缺(會留空)</div>
              {res.missing_in_excel.length === 0 ? <p className="text-sm text-slate-400">無</p> :
                <div className="flex flex-wrap gap-1.5">{res.missing_in_excel.map((c: string) => <Badge key={c} tone="red">{c}</Badge>)}</div>}
            </div>
            <div>
              <div className="label">Excel 有、範本沒用到</div>
              {res.extra_in_excel.length === 0 ? <p className="text-sm text-slate-400">無</p> :
                <div className="flex flex-wrap gap-1.5">{res.extra_in_excel.map((c: string) => <Badge key={c} tone="slate">{c}</Badge>)}</div>}
            </div>
          </div>
        </div>
      ))}
    </Card>
  );
}

/* ---------------- Step 3: Generate ---------------- */
function StepGenerate({ reportType, settings }: any) {
  const [running, setRunning] = useState(false);
  const [progress, setProgress] = useState({ produced: 0, total: 0 });
  const [files, setFiles] = useState<string[]>([]);
  const [result, setResult] = useState<any>(null);

  const start = () => {
    setRunning(true); setFiles([]); setResult(null); setProgress({ produced: 0, total: 0 });
    streamGenerate(reportType,
      (e) => {
        if (e.type === "progress") { setProgress({ produced: e.produced, total: e.total }); setFiles((f) => [e.file, ...f].slice(0, 50)); }
        else if (e.type === "done") setResult(e.result);
        else if (e.type === "error") setResult({ error: e.error });
      },
      () => setRunning(false));
  };
  const cancel = async () => { await api.generateCancel(); };
  const pct = progress.total ? Math.round((progress.produced / progress.total) * 100) : 0;

  return (
    <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
      <Card>
        <SectionTitle icon={<Rocket size={18} />} title="批次產出" desc={settings.enable_review ? "已啟用 AI 審查,產出後會自動評分" : "直接產出(未啟用審查)"} />
        <div className="mb-4 space-y-1.5 text-sm">
          <Row label="範本" value={reportType === "word" ? settings.word_path : reportType === "excel" ? settings.excel_template_path : settings.pptx_template_path} />
          <Row label="資料" value={settings.excel_path} />
          <Row label="輸出" value={settings.output_dir || "Generated_Reports"} />
        </div>
        {!running ? (
          <button className="btn-primary w-full" onClick={start}><Play size={16} /> 開始產出</button>
        ) : (
          <button className="btn-danger w-full" onClick={cancel}><X size={16} /> 取消產出</button>
        )}
        <div className="mt-4">
          <div className="mb-1 flex justify-between text-xs text-slate-500">
            <span>{running ? "產出中…" : result ? "完成" : "待命"}</span>
            <span>{progress.produced}/{progress.total} ({pct}%)</span>
          </div>
          <div className="h-2.5 overflow-hidden rounded-full bg-slate-200 dark:bg-slate-800">
            <div className="h-full rounded-full bg-brand-600 transition-all duration-200" style={{ width: `${pct}%` }} />
          </div>
        </div>
        {result && (
          <div className={clsx("mt-4 rounded-lg p-3 text-sm", result.error ? "bg-rose-50 text-rose-700 dark:bg-rose-900/30 dark:text-rose-300" : "bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300")}>
            {result.error ? <>失敗:{result.error}</> : (
              <>
                <div className="flex items-center gap-2 font-medium"><CheckCircle2 size={16} /> 產出 {result.produced}/{result.total} 份</div>
                {typeof result.failed_count === "number" && result.failed_count > 0 && (
                  <div className="mt-1 text-amber-700 dark:text-amber-300">審查不合格 {result.failed_count} 份(已移至 Failed_Reports)</div>
                )}
                <button className="btn-outline mt-2" onClick={() => api.openOutput()}><FolderOpen size={15} /> 開啟輸出資料夾</button>
              </>
            )}
          </div>
        )}
      </Card>

      <Card>
        <SectionTitle icon={<FileText size={18} />} title="即時產出清單" desc="最近產出的檔案" />
        {files.length === 0 ? <Empty icon={<FileText size={26} />} text="尚未開始" /> : (
          <div className="space-y-1 overflow-auto font-mono text-xs" style={{ maxHeight: 360 }}>
            {files.map((f, i) => (
              <div key={i} className="flex items-center gap-2 rounded px-2 py-1 text-slate-600 dark:text-slate-400">
                <CheckCircle2 size={13} className="text-emerald-500" /> {f}
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

/* ---------------- Step 4: Review ---------------- */
function StepReview({ settings }: any) {
  return (
    <Card>
      <SectionTitle icon={<ShieldCheck size={18} />} title="完成 / 審查" desc="產出結果檢視" />
      <p className="text-sm text-slate-500 dark:text-slate-400">
        {settings.enable_review
          ? "已啟用 AI 審查:產出階段會逐份用視覺模型評分,不合格的自動移到 Failed_Reports 資料夾。可至「AI 引擎」調整抽樣比例與評分標準。"
          : "目前未啟用 AI 審查。若要逐份自動檢查產出品質,請到「AI 引擎」開啟審查並選好視覺模型。"}
      </p>
      <div className="mt-4 flex gap-2">
        <button className="btn-outline" onClick={() => api.openOutput()}><FolderOpen size={15} /> 開啟輸出資料夾</button>
      </div>
      <div className="mt-5 flex items-center gap-2 rounded-lg bg-brand-50 p-4 text-sm text-brand-800 dark:bg-brand-900/20 dark:text-brand-200">
        <Sparkles size={18} /> 完成!整個流程:設定來源 → 對應標籤 → 驗證 → 產出 → 審查,一條龍跑完。
      </div>
    </Card>
  );
}

function Row({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex gap-2">
      <span className="w-12 shrink-0 text-slate-400">{label}</span>
      <span className="truncate font-mono text-xs text-slate-600 dark:text-slate-400" title={value}>{value || "—"}</span>
    </div>
  );
}
