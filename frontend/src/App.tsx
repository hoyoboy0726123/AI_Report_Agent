import { useEffect, useState } from "react";
import clsx from "clsx";
import {
  FileText, ArrowLeftRight, Bot, Cpu, Settings as SettingsIcon,
  Moon, Sun, Sparkles, Images,
} from "lucide-react";
import { useSettings } from "./lib/store";
import ReportWizard from "./screens/ReportWizard";
import ExcelTransfer from "./screens/ExcelTransfer";
import ImageMapping from "./screens/ImageMapping";
import AiEngine from "./screens/AiEngine";
import AgentChat from "./screens/AgentChat";
import SettingsScreen from "./screens/Settings";

type Tab = "wizard" | "transfer" | "images" | "ai" | "agent" | "settings";

const NAV: { id: Tab; label: string; icon: any; desc: string }[] = [
  { id: "wizard", label: "報告精靈", icon: FileText, desc: "範本 + 資料 → 批次報告" },
  { id: "images", label: "圖片對應", icon: Images, desc: "資料夾圖片 → 欄位" },
  { id: "transfer", label: "Excel 搬移", icon: ArrowLeftRight, desc: "欄位對應搬移資料" },
  { id: "agent", label: "AI 助手", icon: Bot, desc: "用自然語言操作" },
  { id: "ai", label: "AI 引擎", icon: Cpu, desc: "模型與審查設定" },
  { id: "settings", label: "設定", icon: SettingsIcon, desc: "預設值與外觀" },
];

export default function App() {
  const { ready } = useSettings();
  const [tab, setTab] = useState<Tab>("wizard");
  const [dark, setDark] = useState(
    () => window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false
  );

  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
  }, [dark]);

  return (
    <div className="flex h-full">
      {/* Sidebar */}
      <aside className="flex w-60 shrink-0 flex-col border-r border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-900">
        <div className="flex items-center gap-2 px-5 py-5">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-brand-600 text-white">
            <Sparkles size={20} />
          </div>
          <div>
            <div className="text-sm font-bold leading-tight">AI Report Agent</div>
            <div className="text-[11px] text-slate-400">辦公自動化工作台</div>
          </div>
        </div>

        <nav className="flex-1 space-y-1 px-3">
          {NAV.map((n) => {
            const active = tab === n.id;
            const Icon = n.icon;
            return (
              <button
                key={n.id}
                onClick={() => setTab(n.id)}
                className={clsx(
                  "flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-left transition-colors duration-150 cursor-pointer",
                  active
                    ? "bg-brand-50 text-brand-700 dark:bg-brand-900/30 dark:text-brand-300"
                    : "text-slate-600 hover:bg-slate-100 dark:text-slate-300 dark:hover:bg-slate-800"
                )}
              >
                <Icon size={18} className="shrink-0" />
                <div className="min-w-0">
                  <div className="text-sm font-medium">{n.label}</div>
                  <div className="truncate text-[11px] text-slate-400">{n.desc}</div>
                </div>
              </button>
            );
          })}
        </nav>

        <div className="px-3 pb-4">
          <button onClick={() => setDark((d) => !d)} className="btn-ghost w-full justify-start">
            {dark ? <Sun size={18} /> : <Moon size={18} />}
            {dark ? "淺色模式" : "深色模式"}
          </button>
        </div>
      </aside>

      {/* Main */}
      <main className="flex-1 overflow-y-auto">
        {!ready ? (
          <div className="flex h-full items-center justify-center text-slate-400">載入中…</div>
        ) : (
          <div className="mx-auto w-full max-w-[1700px] px-10 py-8 animate-fade-in">
            {tab === "wizard" && <ReportWizard />}
            {tab === "images" && <ImageMapping />}
            {tab === "transfer" && <ExcelTransfer />}
            {tab === "ai" && <AiEngine />}
            {tab === "agent" && <AgentChat />}
            {tab === "settings" && <SettingsScreen />}
          </div>
        )}
      </main>
    </div>
  );
}
