import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import remarkCjkFriendly from "remark-cjk-friendly";

// 模型常輸出簡單 LaTeX(如 $\rightarrow$、$\times$),沒裝 katex 就轉成 unicode,避免顯示原始碼。
const LATEX_MAP: Record<string, string> = {
  "\\rightarrow": "→", "\\to": "→", "\\leftarrow": "←", "\\Rightarrow": "⇒",
  "\\times": "×", "\\div": "÷", "\\pm": "±", "\\cdot": "·",
  "\\leq": "≤", "\\geq": "≥", "\\neq": "≠", "\\approx": "≈",
  "\\alpha": "α", "\\beta": "β", "\\Delta": "Δ", "\\checkmark": "✓",
};

function delatex(s: string): string {
  // 把 $...$ / \(...\) 內的簡單符號換成 unicode;其餘去掉錢字號
  return s
    .replace(/\$([^$]+)\$/g, (_m, inner) => stripTokens(inner))
    .replace(/\\\(([^)]*)\\\)/g, (_m, inner) => stripTokens(inner))
    .replace(/\\\[([^\]]*)\\\]/g, (_m, inner) => stripTokens(inner));
}
function stripTokens(s: string): string {
  let out = s;
  for (const [k, v] of Object.entries(LATEX_MAP)) out = out.split(k).join(v);
  return out.replace(/\\(text|mathrm|mathbf)\{([^}]*)\}/g, "$2").replace(/[{}]/g, "").trim();
}

export function Markdown({ children, invert = false }: { children: string; invert?: boolean }) {
  const text = delatex(children || "");
  const linkCls = invert ? "underline text-white/90" : "text-brand-600 underline dark:text-brand-400";
  const codeCls = invert
    ? "rounded bg-white/20 px-1 py-0.5 font-mono text-[0.85em]"
    : "rounded bg-slate-200 px-1 py-0.5 font-mono text-[0.85em] text-rose-600 dark:bg-slate-700 dark:text-rose-300";
  return (
    <div className="space-y-2 text-sm leading-relaxed [&_p]:m-0">
      <ReactMarkdown
        remarkPlugins={[remarkGfm, remarkCjkFriendly]}
        components={{
          p: ({ children }) => <p className="whitespace-pre-wrap">{children}</p>,
          strong: ({ children }) => <strong className="font-semibold">{children}</strong>,
          em: ({ children }) => <em className="italic">{children}</em>,
          ul: ({ children }) => <ul className="ml-4 list-disc space-y-0.5">{children}</ul>,
          ol: ({ children }) => <ol className="ml-4 list-decimal space-y-0.5">{children}</ol>,
          li: ({ children }) => <li className="leading-snug">{children}</li>,
          a: ({ href, children }) => <a href={href} target="_blank" rel="noreferrer" className={linkCls}>{children}</a>,
          h1: ({ children }) => <h3 className="text-base font-bold">{children}</h3>,
          h2: ({ children }) => <h3 className="text-base font-bold">{children}</h3>,
          h3: ({ children }) => <h4 className="font-semibold">{children}</h4>,
          code: ({ className, children }) => {
            const block = (className || "").includes("language-");
            if (block) return <code className="font-mono text-xs">{children}</code>;
            return <code className={codeCls}>{children}</code>;
          },
          pre: ({ children }) => (
            <pre className={invert ? "overflow-x-auto rounded-lg bg-black/25 p-2.5 text-xs" : "overflow-x-auto rounded-lg bg-slate-900 p-2.5 text-xs text-slate-100 dark:bg-slate-950"}>{children}</pre>
          ),
          table: ({ children }) => <div className="overflow-x-auto"><table className="border-collapse text-xs">{children}</table></div>,
          th: ({ children }) => <th className="border border-slate-300 px-2 py-1 font-semibold dark:border-slate-600">{children}</th>,
          td: ({ children }) => <td className="border border-slate-200 px-2 py-1 dark:border-slate-700">{children}</td>,
          blockquote: ({ children }) => <blockquote className="border-l-2 border-slate-300 pl-3 italic opacity-80 dark:border-slate-600">{children}</blockquote>,
          hr: () => <hr className="border-slate-200 dark:border-slate-700" />,
        }}
      >
        {text}
      </ReactMarkdown>
    </div>
  );
}
