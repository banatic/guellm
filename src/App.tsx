import { useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { Minus, Square, X, MessageSquare, FlaskConical } from "lucide-react";
import ChatSidebar from "./components/ChatSidebar";
import ChatInterface from "./components/ChatInterface";
import SettingsModal from "./components/SettingsModal";
import ToolTestPage from "./components/ToolTestPage";
import { useAppStore } from "./store/useAppStore";
import type { AppConfig } from "./types";

type Tab = "chat" | "tools";

export default function App() {
  const loadConfig = useAppStore((s) => s.loadConfig);
  const appWindow = useRef(getCurrentWindow());
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [activeTab, setActiveTab] = useState<Tab>("chat");

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
      <div className="drag-region flex items-center justify-between px-4 h-11 shrink-0 bg-[#141415] border-b border-white/[0.06]">
        <div className="flex items-center gap-5">
          <span className="text-[23px] text-text-secondary" style={{ fontFamily: "'Always Together'", fontWeight: 'normal', transform: 'translateY(3.5px)' }}>
            GEULLM
          </span>
          {/* 탭 전환 */}
          <div className="no-drag flex items-center gap-0.5 bg-white/[0.05] rounded-md p-0.5">
            <button
              onClick={() => setActiveTab("chat")}
              className={`flex items-center gap-1.5 px-2.5 py-1 rounded text-[11px] font-medium transition-all ${
                activeTab === "chat"
                  ? "bg-white/[0.1] text-text-primary"
                  : "text-text-tertiary hover:text-text-secondary"
              }`}
            >
              <MessageSquare size={11} />
              채팅
            </button>
            <button
              onClick={() => setActiveTab("tools")}
              className={`flex items-center gap-1.5 px-2.5 py-1 rounded text-[11px] font-medium transition-all ${
                activeTab === "tools"
                  ? "bg-white/[0.1] text-text-primary"
                  : "text-text-tertiary hover:text-text-secondary"
              }`}
            >
              <FlaskConical size={11} />
              도구 테스트
            </button>
          </div>
        </div>
        <div className="no-drag flex items-center gap-0.5">
          <button
            onClick={minimize}
            className="w-8 h-8 flex items-center justify-center rounded-md hover:bg-white/[0.08] text-text-tertiary hover:text-text-secondary transition-all duration-150"
          >
            <Minus size={13} />
          </button>
          <button
            onClick={toggleMax}
            className="w-8 h-8 flex items-center justify-center rounded-md hover:bg-white/[0.08] text-text-tertiary hover:text-text-secondary transition-all duration-150"
          >
            <Square size={11} />
          </button>
          <button
            onClick={close}
            className="w-8 h-8 flex items-center justify-center rounded-md hover:bg-error/20 text-text-tertiary hover:text-error transition-all duration-150"
          >
            <X size={13} />
          </button>
        </div>
      </div>

      {/* Main layout */}
      <div className="flex flex-1 min-h-0 overflow-hidden">
        {activeTab === "chat" && (
          <ChatSidebar
            collapsed={sidebarCollapsed}
            onToggle={() => setSidebarCollapsed((c) => !c)}
          />
        )}
        <div className="flex-1 min-w-0 flex flex-col overflow-hidden bg-[#212123]">
          {activeTab === "chat" ? <ChatInterface /> : <ToolTestPage />}
        </div>
      </div>

      {/* Settings modal */}
      <SettingsModal />
    </div>
  );
}
