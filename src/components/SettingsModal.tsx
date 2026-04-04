import { useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import {
  X,
  Save,
  Check,
  ChevronDown,
  Plug,
  Loader2,
  Eye,
  EyeOff,
} from "lucide-react";
import { invoke } from "@tauri-apps/api/core";
import { useAppStore } from "../store/useAppStore";
import type { Provider } from "../types";
import { DEFAULT_MODELS, PROVIDER_LABELS } from "../types";

export default function SettingsModal() {
  const {
    settingsOpen,
    setSettingsOpen,
    provider,
    model,
    apiKeys,
    setProvider,
    setModel,
    setApiKey,
    isConnected,
    setConnected,
    selectedFile,
    toConfig,
  } = useAppStore();

  const [saveStatus, setSaveStatus] = useState<"idle" | "saved">("idle");
  const [showKey, setShowKey] = useState(false);
  const [isConnecting, setIsConnecting] = useState(false);

  const apiKey = apiKeys[provider] ?? "";

  async function saveConfig() {
    setApiKey(provider, apiKey);
    const cfg = toConfig();
    try {
      await invoke("update_config", { newConfig: cfg });
      setSaveStatus("saved");
      setTimeout(() => setSaveStatus("idle"), 2000);
    } catch (e) {
      console.error(e);
    }
  }

  async function connectHwp() {
    if (!selectedFile) return;
    setIsConnecting(true);
    try {
      await invoke("connect_hwp", { visible: true });
      await invoke("open_file_in_hwp", { path: selectedFile });
      setConnected(true);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      alert(`연결 실패: ${msg}`);
    } finally {
      setIsConnecting(false);
    }
  }

  if (!settingsOpen) return null;

  return (
    <AnimatePresence>
      <motion.div
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        transition={{ duration: 0.15 }}
        className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm"
        onClick={(e) => {
          if (e.target === e.currentTarget) setSettingsOpen(false);
        }}
      >
        <motion.div
          initial={{ opacity: 0, scale: 0.96, y: 8 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.96, y: 8 }}
          transition={{ duration: 0.2, ease: "easeOut" }}
          className="w-full max-w-md bg-[#2a2a2c] rounded-2xl border border-white/[0.08] shadow-2xl overflow-hidden"
        >
          {/* Header */}
          <div className="flex items-center justify-between px-6 py-4 border-b border-white/[0.06]">
            <h2 className="text-[15px] font-semibold text-text">설정</h2>
            <button
              onClick={() => setSettingsOpen(false)}
              className="w-8 h-8 flex items-center justify-center rounded-lg text-text-tertiary hover:text-text-secondary hover:bg-white/[0.06] transition-colors"
            >
              <X size={16} />
            </button>
          </div>

          {/* Content */}
          <div className="px-6 py-5 space-y-5">
            {/* LLM Provider */}
            <div>
              <label className="block text-[12px] font-medium text-text-secondary mb-2">
                LLM 공급자
              </label>
              <div className="relative">
                <select
                  value={provider}
                  onChange={(e) => setProvider(e.target.value as Provider)}
                  className="w-full bg-white/[0.04] border border-white/[0.08] text-[13px] text-text px-3 py-2.5 rounded-xl outline-none focus:border-accent transition-colors cursor-pointer"
                >
                  {(["openai", "gemini", "anthropic"] as Provider[]).map(
                    (p) => (
                      <option
                        key={p}
                        value={p}
                        className="bg-surface text-text"
                      >
                        {PROVIDER_LABELS[p]}
                      </option>
                    )
                  )}
                </select>
                <ChevronDown
                  size={12}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-text-tertiary pointer-events-none"
                />
              </div>
            </div>

            {/* Model */}
            <div>
              <label className="block text-[12px] font-medium text-text-secondary mb-2">
                모델
              </label>
              <input
                type="text"
                value={model}
                onChange={(e) => setModel(e.target.value)}
                className="w-full bg-white/[0.04] border border-white/[0.08] text-[13px] text-text px-3 py-2.5 rounded-xl outline-none focus:border-accent transition-colors"
                placeholder={DEFAULT_MODELS[provider]}
              />
            </div>

            {/* API Key */}
            <div>
              <label className="block text-[12px] font-medium text-text-secondary mb-2">
                API Key
              </label>
              <div className="relative">
                <input
                  type={showKey ? "text" : "password"}
                  value={apiKey}
                  onChange={(e) => setApiKey(provider, e.target.value)}
                  className="w-full bg-white/[0.04] border border-white/[0.08] text-[13px] text-text px-3 py-2.5 rounded-xl outline-none focus:border-accent transition-colors pr-10"
                  placeholder="sk-..."
                />
                <button
                  onClick={() => setShowKey(!showKey)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-text-tertiary hover:text-text-secondary transition-colors"
                  type="button"
                >
                  {showKey ? <EyeOff size={14} /> : <Eye size={14} />}
                </button>
              </div>
            </div>

            {/* HWP Connection */}
            {selectedFile && (
              <div className="pt-2 border-t border-white/[0.06]">
                <label className="block text-[12px] font-medium text-text-secondary mb-2">
                  HWP 연결
                </label>
                <div className="flex items-center gap-2">
                  <div
                    className={`w-2 h-2 rounded-full ${
                      isConnected ? "bg-success" : "bg-text-tertiary"
                    }`}
                  />
                  <span className="text-[13px] text-text-secondary flex-1">
                    {isConnected ? "연결됨" : "연결 안 됨"}
                  </span>
                  {!isConnected && (
                    <button
                      onClick={connectHwp}
                      disabled={isConnecting}
                      className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[12px] font-medium bg-accent text-white hover:bg-accent-hover transition-colors disabled:opacity-50"
                    >
                      {isConnecting ? (
                        <Loader2 size={12} className="animate-spin" />
                      ) : (
                        <Plug size={12} />
                      )}
                      {isConnecting ? "연결 중..." : "연결"}
                    </button>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Footer */}
          <div className="flex items-center justify-end gap-2 px-6 py-4 border-t border-white/[0.06]">
            <button
              onClick={() => setSettingsOpen(false)}
              className="px-4 py-2 rounded-xl text-[13px] text-text-secondary hover:bg-white/[0.05] transition-colors"
            >
              닫기
            </button>
            <button
              onClick={saveConfig}
              className={`flex items-center gap-1.5 px-4 py-2 rounded-xl text-[13px] font-medium transition-all ${
                saveStatus === "saved"
                  ? "bg-success/10 text-success"
                  : "bg-accent text-white hover:bg-accent-hover"
              }`}
            >
              {saveStatus === "saved" ? (
                <Check size={13} />
              ) : (
                <Save size={13} />
              )}
              {saveStatus === "saved" ? "저장됨" : "저장"}
            </button>
          </div>
        </motion.div>
      </motion.div>
    </AnimatePresence>
  );
}
