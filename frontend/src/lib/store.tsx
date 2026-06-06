import { createContext, useContext, useEffect, useState, useCallback, type ReactNode } from "react";
import { api, type Settings } from "./api";

interface Store {
  settings: Settings;
  ready: boolean;
  set: (patch: Settings) => Promise<void>;
  reload: () => Promise<void>;
}

const Ctx = createContext<Store>(null as any);

export function SettingsProvider({ children }: { children: ReactNode }) {
  const [settings, setSettings] = useState<Settings>({});
  const [ready, setReady] = useState(false);

  const reload = useCallback(async () => {
    const s = await api.getSettings();
    setSettings(s);
    setReady(true);
  }, []);

  useEffect(() => { reload(); }, [reload]);

  const set = useCallback(async (patch: Settings) => {
    setSettings((prev) => ({ ...prev, ...patch })); // 樂觀更新
    const res = await api.updateSettings(patch);
    if (res?.settings) setSettings(res.settings);
  }, []);

  return <Ctx.Provider value={{ settings, ready, set, reload }}>{children}</Ctx.Provider>;
}

export const useSettings = () => useContext(Ctx);
