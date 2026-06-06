import { useCallback, useEffect, useState } from "react";
import clsx from "clsx";
import {
  Images, Wand2, Play, Image as ImageIcon, ArrowRight,
  CheckCircle2, FileText, Eye, Sparkles, Table2, FolderOpen, AlertTriangle,
} from "lucide-react";
import { PathPicker } from "../components/PathPicker";
import { Card, Field, SectionTitle, Empty } from "../components/ui";
import { api } from "../lib/api";
import { useSettings } from "../lib/store";

type Mode = "static" | "visual" | "structured";

export default function ImageMapping() {
  const { settings, set } = useSettings();
  const [mode, setMode] = useState<Mode>("structured");
  const [folder, setFolder] = useState("");
  const [images, setImages] = useState<{ name: string; path: string }[]>([]);

  const loadImages = useCallback(async (f: string) => {
    if (!f) return;
    const r = await api.imagesList(f);
    setImages(r.images || []);
  }, []);
  useEffect(() => { if (folder) loadImages(folder); }, [folder, loadImages]);

  return (
    <div>
      <h1 className="mb-1 text-xl font-bold">圖片對應</h1>
      <p className="mb-3 text-sm text-slate-500 dark:text-slate-400">
        這頁把照片塞進<b>「一份現成的 Word 檔」</b>。<br />
        💡 若你是「從範本批次產報告、每列放各自的圖」,請改用 <b>報告精靈 → ① 設定來源 → 每列圖片(選用)</b>,不在這裡。
      </p>

      {/* 模式說明 */}
      <div className="mb-5 grid grid-cols-1 gap-2 md:grid-cols-3">
        {([
          ["structured", "制式範本(資料夾→各表)", "範本已含照片、位置固定(如測試報告)。多個子資料夾(01,02…)各對一個照片表,依就近標籤換掉照片。", Table2],
          ["static", "範本共用(配 Word)", "Word 範本裡有 {{圖片欄}} 標籤,把固定一組圖依檔名換進去。例:公司 logo、固定圖表。", FileText],
          ["visual", "視覺自動(Word/PPT·無標註)", "Word/PPT 沒有標籤時,AI 看版面/投影片自己決定每張圖貼哪。", Eye],
        ] as const).map(([id, title, desc, Icon]) => (
          <button key={id} onClick={() => setMode(id)}
            className={clsx("rounded-xl border p-3 text-left transition-colors cursor-pointer",
              mode === id ? "border-brand-500 bg-brand-50 dark:bg-brand-900/20" : "border-slate-200 hover:border-brand-300 dark:border-slate-800")}>
            <div className="mb-1 flex items-center gap-1.5 text-sm font-semibold"><Icon size={15} /> {title}</div>
            <div className="text-xs text-slate-500 dark:text-slate-400">{desc}</div>
          </button>
        ))}
      </div>

      {mode !== "structured" && (
      <Card className="mb-5">
        <SectionTitle icon={<Images size={18} />} title="① 圖片資料夾" desc="選一個裝著要放進報告的照片的資料夾" />
        <PathPicker kind="directory" value={folder} placeholder="點「瀏覽」選含圖片的資料夾" onPick={setFolder} />
        {images.length > 0 && (
          <div className="mt-3">
            <div className="mb-1 text-xs text-slate-400">偵測到 {images.length} 張圖片:</div>
            <div className="flex flex-wrap gap-2">
              {images.map((im) => (
                <span key={im.name} className="chip bg-slate-100 text-slate-600 dark:bg-slate-800 dark:text-slate-300">
                  <ImageIcon size={12} /> {im.name}
                </span>
              ))}
            </div>
          </div>
        )}
        {folder && images.length === 0 && <Empty icon={<Images size={26} />} text="此資料夾沒有圖片(png/jpg/...)" />}
      </Card>
      )}

      {mode === "structured" && <StructuredMode settings={settings} set={set} />}
      {mode === "static" && <StaticMode settings={settings} set={set} images={images} />}
      {mode === "visual" && <VisualMode settings={settings} set={set} folder={folder} images={images} />}
    </div>
  );
}

/* 制式範本:多個子資料夾 → 各照片表,結構化換照片(media 位元組置換) */
function StructuredMode({ settings, set }: any) {
  const [photoRoot, setPhotoRoot] = useState("");
  const [output, setOutput] = useState("");
  const [matchMode, setMatchMode] = useState("auto");
  const [busy, setBusy] = useState<"" | "dry" | "run">("");
  const [preview, setPreview] = useState<any>(null);
  const [result, setResult] = useState<any>(null);

  const tpl = settings.word_path;
  const dry = async () => {
    if (!tpl) return alert("請先選制式範本 (.docx)");
    if (!photoRoot) return alert("請先選照片根目錄(底下有 01,02… 子資料夾)");
    setBusy("dry"); setResult(null);
    setPreview(await api.reportFill(tpl, photoRoot, "", true, matchMode));
    setBusy("");
  };
  const run = async () => {
    setBusy("run"); setResult(null);
    setResult(await api.reportFill(tpl, photoRoot, output, false, matchMode));
    setBusy("");
  };

  return (
    <Card>
      <SectionTitle icon={<Table2 size={18} />} title="② 結構化填圖(制式範本)"
        desc="範本已含照片、位置固定。子資料夾排序後(01,02…)各對一個照片表,依就近標籤(如 Top side/Bottom corner)換掉照片;XML 不動、不錯位。" />
      <Field label="制式範本 (.docx / .pptx)" hint="已內含照片、位置固定的報告:docx 每表=一個樣本、pptx 每張投影片=一個樣本">
        <PathPicker kind="any" value={tpl} placeholder="選擇制式範本 (.docx 或 .pptx)" onPick={(p) => set({ word_path: p })} />
      </Field>
      <Field label="照片根目錄" hint="底下要有 01、02、03… 子資料夾,每個子資料夾 = 一個樣本(docx 對一個表 / pptx 對一張投影片)">
        <PathPicker kind="directory" value={photoRoot} placeholder="選含 01~05 子資料夾的根目錄" onPick={setPhotoRoot} />
      </Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="配對方式" hint="auto=檔名語意優先、看不懂再用視覺;text=只看檔名;vlm=只看圖片內容">
          <select className="input" value={matchMode} onChange={(e) => setMatchMode(e.target.value)}>
            <option value="auto">auto(推薦)</option>
            <option value="text">text(只看檔名)</option>
            <option value="vlm">vlm(只看圖內容·慢)</option>
          </select>
        </Field>
        <Field label="輸出檔(選填)" hint="留空=範本同目錄 *_filled_agent.docx">
          <input className="input font-mono text-xs" value={output} onChange={(e) => setOutput(e.target.value)} placeholder="（留空自動命名）" />
        </Field>
      </div>

      <div className="mt-2 flex gap-2">
        <button className="btn-outline" onClick={dry} disabled={busy !== ""}><Eye size={15} /> {busy === "dry" ? "推導中…" : "① 先預覽對應(dry-run)"}</button>
        <button className="btn-primary" onClick={run} disabled={busy !== "" || !preview || preview?.error || preview?.need_review}>
          <Play size={15} /> {busy === "run" ? "產出中…" : "② 正式產出"}
        </button>
      </div>

      {/* dry-run 對應預覽 */}
      {preview && (
        <div className={clsx("mt-4 rounded-lg p-3 text-sm", preview.error ? "bg-rose-50 text-rose-700 dark:bg-rose-900/30 dark:text-rose-300" : "bg-slate-50 dark:bg-slate-800/50")}>
          {preview.error ? <>失敗:{preview.error}</> : (
            <>
              <div className="flex items-center gap-2 font-medium">
                {preview.need_review ? <AlertTriangle size={16} className="text-amber-500" /> : <CheckCircle2 size={16} className="text-emerald-500" />}
                偵測 {preview.photo_tables} 個照片表 · 對應 {preview.total_matched}/{preview.total_slots} 格
                {preview.need_review && <span className="text-amber-600">(有對不齊,請檢查)</span>}
              </div>
              <div className="mt-2 space-y-1 text-xs text-slate-600 dark:text-slate-300">
                {preview.samples?.map((s: any) => (
                  <div key={s.folder}>資料夾 <b>{s.folder}</b> → 表{s.table_index}:{s.matched}/{s.slots} 格
                    {s.unmatched_files?.length > 0 && <span className="text-amber-600"> · 未配:{s.unmatched_files.join("、")}</span>}</div>
                ))}
              </div>
              {!preview.need_review && <div className="mt-2 text-emerald-700 dark:text-emerald-300">對應正確,可按「② 正式產出」。</div>}
            </>
          )}
        </div>
      )}

      {/* 正式產出結果 */}
      {result && (
        <div className={clsx("mt-3 rounded-lg p-3 text-sm", result.error ? "bg-rose-50 text-rose-700 dark:bg-rose-900/30 dark:text-rose-300" : "bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300")}>
          {result.error ? <>失敗:{result.error}</> : (
            <>
              <div className="flex items-center gap-2 font-medium"><CheckCircle2 size={16} /> 完成!置換 {result.replaced}/{result.expected} 張照片</div>
              <div className="mt-1 font-mono text-xs opacity-80">輸出 → {result.output}</div>
              <button className="btn-outline mt-2" onClick={() => api.openOutput()}><FolderOpen size={14} /> 開啟輸出資料夾</button>
            </>
          )}
        </div>
      )}
    </Card>
  );
}

function Row({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex gap-2">
      <span className="w-16 shrink-0 text-slate-400">{label}</span>
      <span className="truncate font-mono text-xs text-slate-600 dark:text-slate-400" title={value}>{value || "—"}</span>
    </div>
  );
}

/* 視覺自動:Word 無標註,AI 看版面決定位置,pywin32 一次貼上 */
function VisualMode({ settings, set, folder, images }: any) {
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<any>(null);

  const run = async () => {
    if (!settings.word_path) return alert("請先選 Word 報告檔");
    if (!folder) return alert("請先選圖片資料夾");
    if (!settings.gemini_reviewer_model && !settings.ollama_reviewer_model)
      return alert("請先到「AI 引擎」選好 reviewer(視覺)模型");
    setBusy(true); setResult(null);
    setResult(await api.imagesAutoVisual(folder, settings.image_width_mm || 80, true));
    setBusy(false);
  };

  return (
    <Card>
      <SectionTitle icon={<Eye size={18} />} title="② 視覺驅動全自動配圖" desc="Word/PPT 沒有任何標籤也能做 —— AI 看版面/投影片判斷位置後貼上" />
      <Field label="報告檔 (.docx / .pptx)" hint="要被填入照片的報告(不需任何標籤);AI 會看版面決定每張圖貼哪。Word→段落定位、PPT→投影片">
        <PathPicker kind="any" value={settings.word_path} placeholder="選擇要配圖的 Word/PPT 報告" onPick={(p) => set({ word_path: p })} />
      </Field>
      <div className="mb-4 space-y-2 rounded-lg bg-brand-50 p-4 text-sm text-brand-800 dark:bg-brand-900/20 dark:text-brand-200">
        <div className="flex items-center gap-2 font-semibold"><Sparkles size={16} /> 運作流程(全自動)</div>
        <ol className="ml-5 list-decimal space-y-0.5 text-[13px]">
          <li>把 Word 報告每頁渲染成圖(需本機 Word)</li>
          <li>把「要放的圖片 + 每頁版面 + 段落文字」一起給視覺模型</li>
          <li>模型依檔名關鍵字 + 實際版面,判斷各圖該放第幾頁、緊鄰哪段</li>
          <li>用 pywin32(Word COM)一次把所有圖貼到定位點並存檔</li>
        </ol>
      </div>
      <div className="mb-3 space-y-1.5 text-sm">
        <Row label="報告" value={settings.word_path} />
        <Row label="圖片夾" value={folder} />
        <Row label="視覺模型" value={settings.gemini_reviewer_model || settings.ollama_reviewer_model || "（未選）"} />
        <Row label="圖片數" value={String(images.length)} />
      </div>
      <button className="btn-primary" onClick={run} disabled={busy || !folder}>
        <Eye size={16} /> {busy ? "AI 視覺配圖中…(每張圖一次視覺呼叫,請稍候)" : "開始視覺自動配圖"}
      </button>

      {result && (
        <div className={clsx("mt-4 rounded-lg p-3 text-sm", result.error ? "bg-rose-50 text-rose-700 dark:bg-rose-900/30 dark:text-rose-300" : "bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300")}>
          {result.error ? <>失敗:{result.error}</> : (
            <>
              <div className="flex items-center gap-2 font-medium">
                <CheckCircle2 size={16} /> 貼上 {result.placed_count}/{result.total_images} 張({result.method}{result.fallback ? ` · 退回 ${result.fallback}` : ""})
              </div>
              {result.placements?.length > 0 && (
                <div className="mt-2 space-y-1 text-xs text-slate-600 dark:text-slate-300">
                  {result.placements.map((p: any, i: number) => (
                    <div key={i} className="rounded bg-white/60 px-2 py-1 dark:bg-slate-900/40">
                      <b>{p.image}</b> → {p.anchor ? `第 ${p.page} 頁「${p.anchor.slice(0, 20)}…」` : "（未定位)"}
                      {p.reason && <span className="text-slate-400"> · {p.reason}</span>}
                    </div>
                  ))}
                </div>
              )}
              {result.placed?.length > 0 && (
                <div className="mt-2 space-y-1 text-xs text-slate-600 dark:text-slate-300">
                  {result.placed.map((p: any, i: number) => (
                    <div key={i} className="rounded bg-white/60 px-2 py-1 dark:bg-slate-900/40">
                      <b>{p.file}</b> → 投影片 {p.slide}「{(p.caption || "").slice(0, 16)}」({p.where})
                    </div>
                  ))}
                </div>
              )}
              {result.output && <div className="mt-1 font-mono text-xs opacity-80">輸出 → {result.output}</div>}
            </>
          )}
        </div>
      )}
    </Card>
  );
}

/* 靜態:檔名 → Word 範本 {{欄位}},直接把標籤換成圖片 */
function StaticMode({ settings, set, images }: any) {
  const [vars, setVars] = useState<string[]>([]);
  const [mapping, setMapping] = useState<Record<string, string>>({}); // image_name -> field
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<any>(null);

  useEffect(() => {
    const p = settings.word_path || "";
    if (!p) return;
    const low = p.toLowerCase();
    const f = low.endsWith(".pptx") ? api.pptxVars(p)
      : low.endsWith(".xlsx") || low.endsWith(".xls") ? api.excelVars(p)
        : api.wordVars(p);
    f.then((r: any) => setVars(r.variables || []));
  }, [settings.word_path]);

  const autoMatch = async (useAi: boolean) => {
    if (images.length === 0) return alert("先選圖片資料夾");
    if (vars.length === 0) return alert("範本沒有 {{欄位}} 標籤");
    setBusy(true);
    const r = await api.imagesMatch(images.map((i: any) => i.name), vars, useAi, "圖片放進報告範本的圖片欄位");
    setBusy(false);
    if (r.error && !r.mapping) return alert(r.error);
    setMapping(r.mapping || {});
    if (r.error) alert(`AI 不可用,已用檔名比對。(${r.error})`);
  };

  const apply = async () => {
    const nameToPath: Record<string, string> = Object.fromEntries(images.map((i: any) => [i.name, i.path]));
    const fieldToPath: Record<string, string> = {};
    for (const [img, field] of Object.entries(mapping)) if (field) fieldToPath[field] = nameToPath[img];
    if (Object.keys(fieldToPath).length === 0) return alert("沒有任何對應");
    setBusy(true);
    const r = await api.imagesApplyStatic(fieldToPath, settings.image_width_mm || 80);
    setBusy(false); setResult(r);
  };

  return (
    <Card>
      <SectionTitle icon={<FileText size={18} />} title="② 圖片 → 範本欄位" desc="每張圖配一個 {{欄位}},套用後把該標籤換成圖片(Word/PPT/Excel 皆可)" />
      <Field label="範本檔 (.docx / .pptx / .xlsx)" hint="含 {{圖片欄}} 標籤的範本;會把標籤換成對應圖片">
        <PathPicker kind="any" value={settings.word_path} placeholder="選擇含圖片標籤的範本" onPick={(p) => set({ word_path: p })} />
      </Field>
      {!settings.word_path ? <Empty icon={<FileText size={26} />} text="先選範本檔" /> : (
      <>
      <div className="mb-3 flex items-center justify-end gap-2">
        <button className="btn-outline" onClick={() => autoMatch(false)} disabled={busy}><Wand2 size={15} /> 檔名比對</button>
        <button className="btn-outline" onClick={() => autoMatch(true)} disabled={busy}><Wand2 size={15} /> AI 比對</button>
      </div>

      {images.length === 0 ? <Empty icon={<Images size={26} />} text="先選圖片資料夾" /> : (
        <div className="space-y-2">
          {images.map((im: any) => (
            <div key={im.name} className="flex items-center gap-3 rounded-lg border border-slate-200 px-3 py-2 dark:border-slate-800">
              <span className="flex w-1/2 items-center gap-2 truncate text-sm" title={im.name}><ImageIcon size={14} className="shrink-0 text-slate-400" /> {im.name}</span>
              <ArrowRight size={14} className="shrink-0 text-slate-400" />
              <select className="input !py-1" value={mapping[im.name] || ""} onChange={(e) => setMapping((m) => ({ ...m, [im.name]: e.target.value }))}>
                <option value="">（不放）</option>
                {vars.map((v) => <option key={v}>{v}</option>)}
              </select>
            </div>
          ))}
        </div>
      )}

      <button className="btn-primary mt-4" onClick={apply} disabled={busy || images.length === 0}>
        <Play size={16} /> {busy ? "套用中…" : "套用到範本"}
      </button>

      {result && (
        <div className={clsx("mt-4 rounded-lg p-3 text-sm", result.error ? "bg-rose-50 text-rose-700 dark:bg-rose-900/30 dark:text-rose-300" : "bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300")}>
          {result.error ? <>失敗:{result.error}</> : (
            <>
              <div className="flex items-center gap-2 font-medium"><CheckCircle2 size={16} /> 已套用 {result.applied_count} 張圖到範本</div>
              {result.failed?.length > 0 && <div className="mt-1 text-amber-700 dark:text-amber-300">失敗 {result.failed.length} 筆(欄位可能不存在於範本)</div>}
            </>
          )}
        </div>
      )}
      </>
      )}
    </Card>
  );
}
