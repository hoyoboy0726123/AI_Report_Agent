import clsx from "clsx";
import { Check } from "lucide-react";

export interface Step { key: string; label: string; hint: string; }

export function Stepper({
  steps, current, onJump, maxReached,
}: {
  steps: Step[];
  current: number;
  onJump: (i: number) => void;
  maxReached: number;
}) {
  return (
    <div className="flex items-center">
      {steps.map((s, i) => {
        const done = i < current;
        const active = i === current;
        const reachable = i <= maxReached;
        return (
          <div key={s.key} className="flex flex-1 items-center last:flex-none">
            <button
              disabled={!reachable}
              onClick={() => reachable && onJump(i)}
              className={clsx(
                "group flex items-center gap-3 rounded-lg px-2 py-1 text-left transition",
                reachable ? "cursor-pointer" : "cursor-not-allowed opacity-50"
              )}
            >
              <span
                className={clsx(
                  "flex h-9 w-9 shrink-0 items-center justify-center rounded-full text-sm font-semibold transition-colors",
                  done && "bg-emerald-500 text-white",
                  active && "bg-brand-600 text-white ring-4 ring-brand-100 dark:ring-brand-900/50",
                  !done && !active && "bg-slate-200 text-slate-500 dark:bg-slate-800 dark:text-slate-400"
                )}
              >
                {done ? <Check size={18} /> : i + 1}
              </span>
              <div className="hidden md:block">
                <div className={clsx("text-sm font-semibold", active ? "text-brand-700 dark:text-brand-300" : "text-slate-700 dark:text-slate-300")}>
                  {s.label}
                </div>
                <div className="text-[11px] text-slate-400">{s.hint}</div>
              </div>
            </button>
            {i < steps.length - 1 && (
              <div className={clsx("mx-2 h-0.5 flex-1 rounded", i < current ? "bg-emerald-400" : "bg-slate-200 dark:bg-slate-800")} />
            )}
          </div>
        );
      })}
    </div>
  );
}
