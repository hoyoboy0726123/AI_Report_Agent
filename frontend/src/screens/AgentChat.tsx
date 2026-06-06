import { useEffect, useRef, useState } from "react";
import clsx from "clsx";
import { Bot, User, Wrench, Send, RotateCcw, CornerDownLeft } from "lucide-react";
import { Card, Empty } from "../components/ui";
import { Markdown } from "../components/Markdown";

interface Msg { role: string; text: string; tool_name?: string; tool_calls?: { name: string; arguments: any }[]; }
interface Interaction { token: string; kind: string; question?: string; choices?: string[]; prompt?: string; }

export default function AgentChat() {
  const [msgs, setMsgs] = useState<Msg[]>([]);
  const [input, setInput] = useState("");
  const [connected, setConnected] = useState(false);
  const [thinking, setThinking] = useState(false);
  const [interaction, setInteraction] = useState<Interaction | null>(null);
  const wsRef = useRef<WebSocket | null>(null);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const proto = location.protocol === "https:" ? "wss" : "ws";
    const ws = new WebSocket(`${proto}://${location.host}/api/agent/ws`);
    wsRef.current = ws;
    ws.onopen = () => setConnected(true);
    ws.onclose = () => setConnected(false);
    ws.onmessage = (ev) => {
      const d = JSON.parse(ev.data);
      if (d.type === "message") {
        if (d.role === "system" && !d.text) return;
        setMsgs((m) => [...m, { role: d.role, text: d.text, tool_name: d.tool_name, tool_calls: d.tool_calls }]);
      } else if (d.type === "turn_done") setThinking(false);
      else if (d.type === "interaction") setInteraction(d);
      else if (d.type === "error") { setMsgs((m) => [...m, { role: "system", text: `[錯誤] ${d.text}` }]); setThinking(false); }
      else if (d.type === "reset_ok") setMsgs([]);
    };
    return () => ws.close();
  }, []);

  useEffect(() => { scrollRef.current?.scrollTo(0, scrollRef.current.scrollHeight); }, [msgs, thinking]);

  const send = () => {
    const t = input.trim();
    if (!t || !connected || thinking) return;
    setMsgs((m) => [...m, { role: "user", text: t }]);
    wsRef.current?.send(JSON.stringify({ type: "user", text: t }));
    setInput(""); setThinking(true);
  };
  const reset = () => { wsRef.current?.send(JSON.stringify({ type: "reset" })); setMsgs([]); };
  const reply = (answer: any) => {
    wsRef.current?.send(JSON.stringify({ type: "interaction_reply", token: interaction!.token, answer }));
    setInteraction(null);
  };

  return (
    <div>
      <div className="mb-1 flex items-center justify-between">
        <h1 className="text-xl font-bold">AI 助手</h1>
        <div className="flex items-center gap-3">
          <span className={clsx("chip", connected ? "bg-emerald-100 text-emerald-700 dark:bg-emerald-900/40 dark:text-emerald-300" : "bg-rose-100 text-rose-700")}>
            {connected ? "已連線" : "未連線"}
          </span>
          <button className="btn-ghost !py-1 text-xs" onClick={reset}><RotateCcw size={14} /> 新對話</button>
        </div>
      </div>
      <p className="mb-4 text-sm text-slate-500 dark:text-slate-400">用自然語言操作:例如「幫我把範本標籤對好」「全部產出報告」。需先到「AI 引擎」選好 Planner 模型。</p>

      <Card className="flex flex-col p-0" >
        <div ref={scrollRef} className="flex-1 space-y-3 overflow-y-auto p-5" style={{ height: 460 }}>
          {msgs.length === 0 && <Empty icon={<Bot size={30} />} text="開始對話吧。試試:「現在的設定是什麼?」" />}
          {msgs.map((m, i) => <MsgBubble key={i} msg={m} />)}
          {thinking && <div className="flex items-center gap-2 text-sm text-slate-400"><Bot size={16} /> 思考中…</div>}
        </div>

        {interaction && (
          <div className="border-t border-amber-200 bg-amber-50 p-4 dark:border-amber-900/50 dark:bg-amber-900/20">
            <div className="mb-2 text-sm font-medium text-amber-800 dark:text-amber-200">{interaction.question || interaction.prompt || "需要你的回覆"}</div>
            {interaction.kind === "ask_user" && interaction.choices ? (
              <div className="flex flex-wrap gap-2">
                {interaction.choices.map((c) => <button key={c} className="btn-outline !py-1 text-sm" onClick={() => reply({ answer: c })}>{c}</button>)}
                <button className="btn-ghost !py-1 text-sm" onClick={() => reply({ cancelled: true })}>取消</button>
              </div>
            ) : interaction.kind === "ask_user" ? (
              <InlineInput onSubmit={(v) => reply({ answer: v })} />
            ) : (
              <div className="text-xs text-amber-700 dark:text-amber-300">此互動需在本機視窗選檔(請改用報告精靈手動選),先取消。<button className="btn-ghost !py-1" onClick={() => reply({ cancelled: true })}>取消</button></div>
            )}
          </div>
        )}

        <div className="border-t border-slate-200 p-3 dark:border-slate-800">
          <div className="flex gap-2">
            <textarea className="input h-12 resize-none" placeholder="輸入訊息… (Ctrl+Enter 送出)" value={input}
              onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); send(); } }} />
            <button className="btn-primary shrink-0" onClick={send} disabled={!connected || thinking}><Send size={16} /></button>
          </div>
          <div className="mt-1 flex items-center gap-1 text-[11px] text-slate-400"><CornerDownLeft size={11} /> Ctrl + Enter 送出</div>
        </div>
      </Card>
    </div>
  );
}

function InlineInput({ onSubmit }: { onSubmit: (v: string) => void }) {
  const [v, setV] = useState("");
  return (
    <div className="flex gap-2">
      <input className="input !py-1 text-sm" value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => e.key === "Enter" && onSubmit(v)} autoFocus />
      <button className="btn-primary !py-1 text-sm" onClick={() => onSubmit(v)}>送出</button>
    </div>
  );
}

function MsgBubble({ msg }: { msg: Msg }) {
  if (msg.role === "tool") {
    return (
      <div className="flex gap-2 text-xs">
        <Wrench size={14} className="mt-0.5 shrink-0 text-slate-400" />
        <div className="min-w-0 flex-1 rounded-lg bg-slate-100 px-3 py-1.5 font-mono text-slate-500 dark:bg-slate-800 dark:text-slate-400">
          <span className="font-semibold">{msg.tool_name}</span> → <span className="break-all">{msg.text.slice(0, 300)}</span>
        </div>
      </div>
    );
  }
  if (msg.role === "system") return <div className="text-center text-xs text-rose-500">{msg.text}</div>;
  const isUser = msg.role === "user";
  return (
    <div className={clsx("flex gap-2", isUser && "flex-row-reverse")}>
      <div className={clsx("flex h-7 w-7 shrink-0 items-center justify-center rounded-full", isUser ? "bg-brand-600 text-white" : "bg-slate-200 text-slate-600 dark:bg-slate-700 dark:text-slate-200")}>
        {isUser ? <User size={15} /> : <Bot size={15} />}
      </div>
      <div className={clsx("max-w-[78%] rounded-2xl px-4 py-2.5 text-sm", isUser ? "bg-brand-600 text-white" : "bg-slate-100 text-slate-800 dark:bg-slate-800 dark:text-slate-200")}>
        {msg.tool_calls && msg.tool_calls.length > 0 && (
          <div className="mb-1.5 flex items-center gap-1 text-xs opacity-70">
            <Wrench size={12} /> 呼叫工具:{msg.tool_calls.map((t) => t.name).join(", ")}
          </div>
        )}
        {msg.text && (isUser ? <div className="whitespace-pre-wrap">{msg.text}</div> : <Markdown>{msg.text}</Markdown>)}
      </div>
    </div>
  );
}
