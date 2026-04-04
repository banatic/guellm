import { useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { Minus, Square, X } from "lucide-react";
import ChatSidebar from "./components/ChatSidebar";
import ChatInterface from "./components/ChatInterface";
import SettingsModal from "./components/SettingsModal";
import { useAppStore } from "./store/useAppStore";
import type { AppConfig } from "./types";

export default function App() {
  const loadConfig = useAppStore((s) => s.loadConfig);
  const appWindow = useRef(getCurrentWindow());
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);

  useEffect(() => {
    invoke<AppConfig>("get_config")
      .then((cfg) => loadConfig(cfg))
      .catch(console.error);
  }, []);

  const minimize = () => appWindow.current.minimize();
  const toggleMax = () => appWindow.current.toggleMaximize();
  const close = () => appWindow.current.close();

  return (
    <div className="flex flex-col h-screen overflow-hidden bg-bg">
      {/* Title bar */}
      <div className="drag-region flex items-center justify-between px-4 h-11 shrink-0 bg-[#161617] border-b border-white/[0.06]">
        <span className="text-[13px] font-medium text-text-secondary tracking-tight">
          굴림 (guellm)
        </span>
        <div className="no-drag flex items-center">
          <button
            onClick={minimize}
            className="w-8 h-8 flex items-center justify-center rounded-md hover:bg-white/[0.06] text-text-tertiary hover:text-text-secondary transition-colors"
          >
            <Minus size={13} />
          </button>
          <button
            onClick={toggleMax}
            className="w-8 h-8 flex items-center justify-center rounded-md hover:bg-white/[0.06] text-text-tertiary hover:text-text-secondary transition-colors"
          >
            <Square size={11} />
          </button>
          <button
            onClick={close}
            className="w-8 h-8 flex items-center justify-center rounded-md hover:bg-error/20 text-text-tertiary hover:text-error transition-colors"
          >
            <X size={13} />
          </button>
        </div>
      </div>

      {/* Main layout */}
      <div className="flex flex-1 min-h-0 overflow-hidden">
        <ChatSidebar
          collapsed={sidebarCollapsed}
          onToggle={() => setSidebarCollapsed((c) => !c)}
        />
        <div className="flex-1 min-w-0 flex flex-col overflow-hidden bg-[#232325]">
          <ChatInterface />
        </div>
      </div>

      {/* Settings modal */}
      <SettingsModal />
    </div>
  );
}
