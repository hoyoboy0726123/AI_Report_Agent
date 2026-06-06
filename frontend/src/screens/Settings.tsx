import { Settings as SettingsIcon, FolderOutput, Keyboard, Image } from "lucide-react";
import { PathPicker } from "../components/PathPicker";
import { Card, Field, SectionTitle } from "../components/ui";
import { useSettings } from "../lib/store";

export default function SettingsScreen() {
  const { settings, set } = useSettings();
  const up = (patch: any) => set(patch);

  return (
    <div>
      <h1 className="mb-1 text-xl font-bold">設定</h1>
      <p className="mb-5 text-sm text-slate-500 dark:text-slate-400">這些是各功能共用的預設值,會立即儲存。</p>

      <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
        <Card>
          <SectionTitle icon={<FolderOutput size={18} />} title="輸出預設" />
          <Field label="預設輸出資料夾"><PathPicker kind="directory" value={settings.output_dir} onPick={(p) => up({ output_dir: p })} /></Field>
          <Field label="Word 報告檔名規則"><input className="input font-mono text-xs" value={settings.filename_template || ""} onChange={(e) => up({ filename_template: e.target.value })} /></Field>
          <Field label="Excel 報告檔名規則"><input className="input font-mono text-xs" value={settings.excel_filename_template || ""} onChange={(e) => up({ excel_filename_template: e.target.value })} /></Field>
        </Card>

        <Card>
          <SectionTitle icon={<Image size={18} />} title="圖片嵌入" />
          <Field label="Word 圖片寬度 (mm)"><input type="number" min={1} className="input w-40" value={settings.image_width_mm || 80} onChange={(e) => up({ image_width_mm: Number(e.target.value) || 80 })} /></Field>
          <Field label="Excel 圖片寬度 (px)"><input type="number" min={50} className="input w-40" value={settings.excel_image_width_px || 320} onChange={(e) => up({ excel_image_width_px: Number(e.target.value) || 320 })} /></Field>
        </Card>

        <Card>
          <SectionTitle icon={<Keyboard size={18} />} title="熱鍵橋接" desc="Ctrl+Shift+M 在真正的 Office 視窗插入標籤(需另啟動橋接程式)" />
          <Field label="熱鍵寫入目標">
            <select className="input" value={settings.hotkey_target || "word"} onChange={(e) => up({ hotkey_target: e.target.value })}>
              <option value="word">Word 範本游標位置</option>
              <option value="excel_template">Excel 範本選取儲存格</option>
            </select>
          </Field>
          <p className="text-xs text-slate-400">提示:熱鍵橋接需要本機 Office (COM)。在報告精靈的「對應標籤」步驟也能用滑鼠完成相同的事,不必依賴熱鍵。</p>
        </Card>

        <Card>
          <SectionTitle icon={<SettingsIcon size={18} />} title="關於" />
          <p className="text-sm text-slate-500 dark:text-slate-400">AI Report Agent — 由原 customtkinter 桌面工具重構的本機 Web 工作台。後端 FastAPI + 既有報告引擎,前端 React。</p>
        </Card>
      </div>
    </div>
  );
}
