import { FolderOpen } from "lucide-react";
import { api } from "../lib/api";

export function PathPicker({
  value, kind, placeholder, onPick,
}: {
  value: string;
  kind: "word" | "excel" | "save_excel" | "pptx" | "directory" | "image" | "any";
  placeholder?: string;
  onPick: (path: string) => void;
}) {
  async function browse() {
    const r = await api.pick(kind);
    if (r.path) onPick(r.path);
  }
  return (
    <div className="flex gap-2">
      <input
        className="input font-mono text-xs"
        value={value || ""}
        placeholder={placeholder}
        onChange={(e) => onPick(e.target.value)}
      />
      <button className="btn-outline shrink-0" onClick={browse} title="瀏覽">
        <FolderOpen size={16} /> 瀏覽
      </button>
    </div>
  );
}
