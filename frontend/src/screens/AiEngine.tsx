import { useCallback, useEffect, useState } from "react";
import { Cpu, RefreshCw, Gauge, ShieldCheck } from "lucide-react";
import { Card, Field, SectionTitle, Badge } from "../components/ui";
import { api } from "../lib/api";
import { useSettings } from "../lib/store";

export default function AiEngine() {
  const { settings, set } = useSettings();
  const up = (patch: any) => set(patch);
  const [models, setModels] = useState<{ available: boolean; text_models: string[]; vision_models: string[]; error?: string }>({ available: false, text_models: [], vision_models: [] });
  const [budget, setBudget] = useState<any>(null);
  const [loading, setLoading] = useState(false);

  const loadModels = useCallback(async () => { setLoading(true); setModels(await api.aiModels()); setLoading(false); }, []);
  const loadBudget = useCallback(async () => setBudget(await api.aiBudget()), []);
  useEffect(() => { loadModels(); loadBudget(); }, [loadModels, loadBudget]);

  const provider = settings.llm_provider || "Gemini";

  return (
    <div>
      <h1 className="mb-1 text-xl font-bold">AI 引擎</h1>
      <p className="mb-5 text-sm text-slate-500 dark:text-slate-400">設定產報告與審查用的模型,以及呼叫次數預算。</p>

      <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
        <Card>
          <SectionTitle icon={<Cpu size={18} />} title="供應商與模型" />
          <Field label="供應商">
            <div className="flex gap-2">
              {["Gemini", "Ollama"].map((p) => (
                <button key={p} onClick={() => up({ llm_provider: p })}
                  className={`flex-1 rounded-lg border py-2 text-sm font-medium transition-colors cursor-pointer ${provider === p ? "border-brand-500 bg-brand-50 text-brand-700 dark:bg-brand-900/20 dark:text-brand-300" : "border-slate-200 dark:border-slate-800"}`}>{p}</button>
              ))}
            </div>
          </Field>

          {provider === "Ollama" && (
            <Field label="Ollama Endpoint"><input className="input" value={settings.ollama_endpoint || ""} onChange={(e) => up({ ollama_endpoint: e.target.value })} /></Field>
          )}

          <div className="mb-2 flex items-center justify-between">
            <span className="text-xs text-slate-400">
              {loading ? "連線中…" : models.available ? <Badge tone="green">已連線 · {models.text_models.length} 個模型</Badge> : <Badge tone="red">{models.error || "未連線"}</Badge>}
            </span>
            <button className="btn-ghost !py-1 text-xs" onClick={loadModels}><RefreshCw size={14} /> 測試連線</button>
          </div>

          <Field label="Planner 模型（產報告 / 對應建議 / Agent）">
            <select className="input" value={provider === "Gemini" ? settings.gemini_planner_model || "" : settings.ollama_planner_model || ""}
              onChange={(e) => up(provider === "Gemini" ? { gemini_planner_model: e.target.value } : { ollama_planner_model: e.target.value })}>
              <option value="">（未選）</option>
              {models.text_models.map((m) => <option key={m}>{m}</option>)}
            </select>
          </Field>
          <Field label="Reviewer 模型（視覺審查產出品質）">
            <select className="input" value={provider === "Gemini" ? settings.gemini_reviewer_model || "" : settings.ollama_reviewer_model || ""}
              onChange={(e) => up(provider === "Gemini" ? { gemini_reviewer_model: e.target.value } : { ollama_reviewer_model: e.target.value })}>
              <option value="">（未選）</option>
              {models.vision_models.map((m) => <option key={m}>{m}</option>)}
            </select>
          </Field>
        </Card>

        <div className="space-y-5">
          <Card>
            <SectionTitle icon={<ShieldCheck size={18} />} title="產出審查" desc="產報告後逐份用視覺模型評分" />
            <label className="mb-3 flex items-center gap-2 text-sm cursor-pointer">
              <input type="checkbox" checked={!!settings.enable_review} onChange={(e) => up({ enable_review: e.target.checked })} className="h-4 w-4 rounded" />
              啟用審查
            </label>
            <Field label={`抽樣比例:${settings.review_sampling_percent ?? 100}%`}>
              <input type="range" min={1} max={100} value={settings.review_sampling_percent ?? 100} onChange={(e) => up({ review_sampling_percent: Number(e.target.value) })} className="w-full" />
            </Field>
            <Field label="評分標準 (rubric)">
              <textarea className="input h-28 resize-none font-mono text-xs" value={settings.review_rubric || ""} onChange={(e) => up({ review_rubric: e.target.value })} />
            </Field>
          </Card>

          <Card>
            <SectionTitle icon={<Gauge size={18} />} title="呼叫預算" desc="避免失控的 API 用量上限" />
            {budget && (
              <div className="mb-3 grid grid-cols-2 gap-3 text-sm">
                <div className="rounded-lg bg-slate-50 p-3 dark:bg-slate-800/50">
                  <div className="text-xs text-slate-400">Planner</div>
                  <div className="text-lg font-bold">{budget.planner_used} / {budget.planner_limit}</div>
                </div>
                <div className="rounded-lg bg-slate-50 p-3 dark:bg-slate-800/50">
                  <div className="text-xs text-slate-400">Reviewer</div>
                  <div className="text-lg font-bold">{budget.reviewer_used} / {budget.reviewer_limit}</div>
                </div>
              </div>
            )}
            <div className="grid grid-cols-2 gap-3">
              <Field label="Planner 上限"><input type="number" className="input" value={settings.max_planner_calls || 50} onChange={(e) => up({ max_planner_calls: Number(e.target.value) || 50 })} /></Field>
              <Field label="Reviewer 上限"><input type="number" className="input" value={settings.max_reviewer_calls || 100} onChange={(e) => up({ max_reviewer_calls: Number(e.target.value) || 100 })} /></Field>
            </div>
            <Field label="視覺配速 RPM(每分鐘上限,0=依模型自動)" hint="主動冷卻:達『上限-1』就自動等視窗清出再續,絕不撞 429。慢一點但穩定不中斷。">
              <input type="number" min={0} className="input w-40" value={settings.vision_rpm ?? 0} onChange={(e) => up({ vision_rpm: Number(e.target.value) || 0 })} />
            </Field>
            <Field label="VLM 看圖配對方式" hint="整批送=一次送多張(快,需多圖能力強的模型);單張送=逐張描述(慢,弱模型較準)">
              <select className="input w-48" value={settings.vlm_match_strategy || "batched"} onChange={(e) => up({ vlm_match_strategy: e.target.value })}>
                <option value="batched">整批送(快)</option>
                <option value="describe">單張送(較準)</option>
              </select>
            </Field>
            <button className="btn-outline" onClick={async () => { await api.aiBudgetReset(); loadBudget(); }}><RefreshCw size={14} /> 重置計數</button>
          </Card>
        </div>
      </div>
    </div>
  );
}
