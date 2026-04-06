import { useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import {
  X,
  Check,
  Plug,
  Loader2,
  Eye,
  EyeOff,
  ChevronDown,
} from "lucide-react";
import { invoke } from "@tauri-apps/api/core";
import { useAppStore } from "../store/useAppStore";
import type { Provider } from "../types";
import { DEFAULT_MODELS, PROVIDER_LABELS } from "../types";

const ZOOM_STEPS = [0.75, 0.875, 1.0, 1.125, 1.25, 1.5, 1.75, 2.0];

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
    zoom,
    setZoom,
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

  const zoomIndex = ZOOM_STEPS.indexOf(zoom) !== -1
    ? ZOOM_STEPS.indexOf(zoom)
    : ZOOM_STEPS.findIndex((s) => s >= zoom) || 2;

  return (
    <AnimatePresence>
      <motion.div
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        transition={{ duration: 0.18 }}
        className="fixed inset-0 z-50 flex items-center justify-center"
        style={{ backgroundColor: "rgba(0,0,0,0.45)" }}
        onClick={(e) => {
          if (e.target === e.currentTarget) setSettingsOpen(false);
        }}
      >
        <motion.div
          initial={{ opacity: 0, scale: 0.97, y: 8 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.97, y: 8 }}
          transition={{ duration: 0.18, ease: [0.25, 0.46, 0.45, 0.94] }}
          className="w-full max-w-sm overflow-hidden rounded-2xl"
          style={{
            background: "rgba(30, 30, 32, 0.72)",
            backdropFilter: "blur(40px) saturate(1.8)",
            WebkitBackdropFilter: "blur(40px) saturate(1.8)",
            border: "1px solid rgba(255,255,255,0.10)",
            boxShadow: "0 32px 64px rgba(0,0,0,0.6), 0 0 0 0.5px rgba(255,255,255,0.05) inset",
          }}
        >
          {/* Header */}
          <div className="flex items-center justify-between px-5 pt-5 pb-4">
            <span className="text-[15px] font-semibold text-white/90 tracking-tight">설정</span>
            <button
              onClick={() => setSettingsOpen(false)}
              className="w-7 h-7 flex items-center justify-center rounded-full text-white/40 hover:text-white/70 transition-colors duration-150"
              style={{ background: "rgba(255,255,255,0.07)" }}
            >
              <X size={13} />
            </button>
          </div>

          {/* Content */}
          <div className="px-5 pb-5 space-y-3">

            {/* LLM 설정 */}
            <div
              className="rounded-xl overflow-hidden"
              style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.07)" }}
            >
              <div className="px-4 pt-3 pb-1">
                <span className="text-[11px] font-semibold text-white/35 uppercase tracking-widest">
                  LLM
                </span>
              </div>

              {/* Provider */}
              <div className="px-4 py-2.5 flex items-center justify-between border-b border-white/[0.05]">
                <span className="text-[13px] text-white/70">공급자</span>
                <div className="relative">
                  <select
                    value={provider}
                    onChange={(e) => setProvider(e.target.value as Provider)}
                    className="appearance-none pr-6 pl-3 py-1.5 rounded-lg text-[13px] text-white/85 outline-none cursor-pointer transition-colors duration-150"
                    style={{
                      background: "rgba(255,255,255,0.07)",
                      border: "1px solid rgba(255,255,255,0.08)",
                    }}
                  >
                    {(["openai", "gemini", "anthropic"] as Provider[]).map((p) => (
                      <option key={p} value={p} className="bg-[#252527]">
                        {PROVIDER_LABELS[p]}
                      </option>
                    ))}
                  </select>
                  <ChevronDown size={11} className="absolute right-1.5 top-1/2 -translate-y-1/2 text-white/35 pointer-events-none" />
                </div>
              </div>

              {/* Model */}
              <div className="px-4 py-2.5 flex items-center justify-between">
                <span className="text-[13px] text-white/70">모델</span>
                <input
                  type="text"
                  value={model}
                  onChange={(e) => setModel(e.target.value)}
                  placeholder={DEFAULT_MODELS[provider]}
                  className="w-44 px-3 py-1.5 rounded-lg text-[13px] text-white/85 outline-none transition-colors duration-150 placeholder:text-white/25 text-right"
                  style={{
                    background: "rgba(255,255,255,0.07)",
                    border: "1px solid rgba(255,255,255,0.08)",
                  }}
                />
              </div>
            </div>

            {/* API Key */}
            <div
              className="rounded-xl overflow-hidden"
              style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.07)" }}
            >
              <div className="px-4 pt-3 pb-1">
                <span className="text-[11px] font-semibold text-white/35 uppercase tracking-widest">
                  인증
                </span>
              </div>
              <div className="px-4 py-2.5">
                <div className="relative">
                  <input
                    type={showKey ? "text" : "password"}
                    value={apiKey}
                    onChange={(e) => setApiKey(provider, e.target.value)}
                    placeholder="API Key"
                    className="w-full px-3 py-2 pr-9 rounded-lg text-[13px] text-white/85 outline-none font-mono transition-colors duration-150 placeholder:text-white/25"
                    style={{
                      background: "rgba(255,255,255,0.07)",
                      border: "1px solid rgba(255,255,255,0.08)",
                    }}
                  />
                  <button
                    onClick={() => setShowKey(!showKey)}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-white/35 hover:text-white/60 transition-colors duration-150"
                    type="button"
                  >
                    {showKey ? <EyeOff size={14} /> : <Eye size={14} />}
                  </button>
                </div>
              </div>
            </div>

            {/* 화면 배율 */}
            <div
              className="rounded-xl overflow-hidden"
              style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.07)" }}
            >
              <div className="px-4 pt-3 pb-1">
                <span className="text-[11px] font-semibold text-white/35 uppercase tracking-widest">
                  화면 배율
                </span>
              </div>
              <div className="px-4 py-3">
                <div className="flex items-center gap-3">
                  <input
                    type="range"
                    min={0}
                    max={ZOOM_STEPS.length - 1}
                    step={1}
                    value={zoomIndex}
                    onChange={(e) => setZoom(ZOOM_STEPS[Number(e.target.value)])}
                    className="flex-1 accent-accent cursor-pointer"
                  />
                  <span className="text-[13px] text-white/70 font-mono w-10 text-right shrink-0">
                    {Math.round(zoom * 100)}%
                  </span>
                </div>
                <div className="flex justify-between mt-1.5 px-0.5">
                  {ZOOM_STEPS.map((s) => (
                    <span key={s} className="text-[10px] text-white/20">
                      {Math.round(s * 100)}
                    </span>
                  ))}
                </div>
              </div>
            </div>

            {/* HWP 연결 */}
            {selectedFile && (
              <div
                className="rounded-xl overflow-hidden"
                style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.07)" }}
              >
                <div className="px-4 pt-3 pb-1">
                  <span className="text-[11px] font-semibold text-white/35 uppercase tracking-widest">
                    HWP 연결
                  </span>
                </div>
                <div className="px-4 py-2.5 flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <div
                      className={`w-2 h-2 rounded-full ${
                        isConnected ? "bg-green-400" : "bg-white/20"
                      }`}
                    />
                    <span className="text-[13px] text-white/70">
                      {isConnected ? "연결됨" : "연결 안 됨"}
                    </span>
                  </div>
                  {!isConnected && (
                    <button
                      onClick={connectHwp}
                      disabled={isConnecting}
                      className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[12px] font-medium text-white/80 transition-colors duration-150 disabled:opacity-40"
                      style={{ background: "rgba(255,255,255,0.09)", border: "1px solid rgba(255,255,255,0.1)" }}
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
          <div
            className="flex items-center justify-end gap-2 px-5 py-4"
            style={{ borderTop: "1px solid rgba(255,255,255,0.07)" }}
          >
            <button
              onClick={() => setSettingsOpen(false)}
              className="px-4 py-2 rounded-xl text-[13px] text-white/50 hover:text-white/70 transition-colors duration-150"
            >
              닫기
            </button>
            <button
              onClick={saveConfig}
              className="flex items-center gap-1.5 px-4 py-2 rounded-xl text-[13px] font-medium transition-all duration-200"
              style={{
                background: saveStatus === "saved"
                  ? "rgba(52, 199, 89, 0.15)"
                  : "rgba(99, 102, 241, 0.8)",
                color: saveStatus === "saved" ? "rgba(52, 199, 89, 1)" : "white",
                border: saveStatus === "saved"
                  ? "1px solid rgba(52, 199, 89, 0.3)"
                  : "1px solid rgba(255,255,255,0.15)",
              }}
            >
              {saveStatus === "saved" ? <Check size={13} /> : null}
              {saveStatus === "saved" ? "저장됨" : "저장"}
            </button>
          </div>
        </motion.div>
      </motion.div>
    </AnimatePresence>
  );
}
