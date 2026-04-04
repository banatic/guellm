import { useRef, useEffect, useState, useCallback } from "react";
import { AnimatePresence, motion } from "framer-motion";
import {
  ArrowUp,
  Trash2,
  XCircle,
  Undo2,
  Paperclip,
  FileText,
  X,
  Plug,
  Loader2,
  Check,
  Eye,
} from "lucide-react";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { useAppStore } from "../store/useAppStore";
import type { AgentEvent } from "../types";
import MessageBubble from "./MessageBubble";

const PLACEHOLDER = "메시지를 입력하세요... (Ctrl+Enter)";

function formatTokens(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M`;
  if (n >= 1_000) return `${(n / 1_000).toFixed(1)}K`;
  return `${n}`;
}

export default function ChatInterface() {
  const [query, setQuery] = useState("");
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const {
    messages,
    isAgentRunning,
    selectedFile,
    isConnected,
    provider,
    model,
    currentApiKey,
    addUserMessage,
    addAssistantMessage,
    appendToolCall,
    updateToolCall,
    setFinalResponse,
    clearMessages,
    setAgentRunning,
    appendTextToAssistant,
    setPendingConfirm,
    hasBackup,
    setHasBackup,
    tokenUsage,
    setTokenUsage,
    setSelectedFile,
    setConnected,
    activeConversationId,
    createConversation,
    setSettingsOpen,
  } = useAppStore();

  const [isConnecting, setIsConnecting] = useState(false);
  const [previewResult, setPreviewResult] = useState<string | null>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  useEffect(() => {
    const ta = textareaRef.current;
    if (!ta) return;
    ta.style.height = "auto";
    ta.style.height = Math.min(ta.scrollHeight, 160) + "px";
  }, [query]);

  async function handleFileOpen() {
    try {
      const path = await invoke<string | null>("open_file_dialog");
      if (path) {
        setSelectedFile(path);
        setConnected(false);
        // 대화가 없으면 새로 생성
        if (!activeConversationId) {
          createConversation(path);
        }
      }
    } catch (e) {
      console.error(e);
    }
  }

  async function handleConnect() {
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

  async function handlePreview() {
    try {
      const result = await invoke<string>("preview_structure");
      const parsed = JSON.parse(result);
      setPreviewResult(JSON.stringify(parsed, null, 2));
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      setPreviewResult(`오류: ${msg}`);
    }
  }

  const handleSubmit = useCallback(async () => {
    const q = query.trim();
    if (!q || isAgentRunning) return;
    if (!isConnected) {
      alert("먼저 한글 문서를 열어주세요.");
      return;
    }
    const apiKey = currentApiKey();
    if (!apiKey) {
      setSettingsOpen(true);
      return;
    }

    const history: { role: string; content: string }[] = [];
    for (const msg of messages) {
      const textContent = msg.contents
        .filter((c): c is { type: "text"; text: string } => c.type === "text")
        .map((c) => c.text)
        .join("\n");
      if (textContent) {
        history.push({ role: msg.role, content: textContent });
      }
    }

    setQuery("");
    setAgentRunning(true);
    setTokenUsage(null);
    addUserMessage(q);
    const assistantMsgId = addAssistantMessage();

    const unlisten = await listen<AgentEvent>("agent-event", (event) => {
      const payload = event.payload;

      if (payload.type === "toolConfirmRequest") {
        appendToolCall(assistantMsgId, {
          name: payload.name,
          args: payload.args,
          status: "pending",
        });
        setPendingConfirm({ name: payload.name, args: payload.args });
      } else if (payload.type === "toolCall") {
        const store = useAppStore.getState();
        const msg = store.messages.find((m) => m.id === assistantMsgId);
        const hasPending = msg?.contents.some(
          (c) =>
            c.type === "tool_call" &&
            c.tool.name === payload.name &&
            c.tool.status === "pending"
        );
        if (hasPending) {
          updateToolCall(assistantMsgId, payload.name, { status: "running" });
        } else {
          appendToolCall(assistantMsgId, {
            name: payload.name,
            args: payload.args,
            status: "running",
          });
        }
        setPendingConfirm(null);
      } else if (payload.type === "toolResult") {
        updateToolCall(assistantMsgId, payload.name, {
          result: payload.result,
          status: payload.result.startsWith("\u274c") ? "error" : "done",
        });
      } else if (payload.type === "llmThinking") {
        appendTextToAssistant(assistantMsgId, payload.text);
      } else if (payload.type === "tokenUsage") {
        setTokenUsage({
          prompt: payload.prompt_tokens,
          completion: payload.completion_tokens,
          total: payload.prompt_tokens + payload.completion_tokens,
        });
      } else if (payload.type === "finalResponse") {
        setFinalResponse(assistantMsgId, payload.text);
        setAgentRunning(false);
        setHasBackup(true);
        setPendingConfirm(null);
        unlisten();
      } else if (payload.type === "error") {
        appendTextToAssistant(assistantMsgId, `오류: ${payload.message}`);
        setAgentRunning(false);
        setHasBackup(true);
        setPendingConfirm(null);
        unlisten();
      }
    });

    try {
      await invoke("run_agent", {
        params: { provider, api_key: apiKey, model, query: q, history },
      });
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      appendTextToAssistant(assistantMsgId, `실행 오류: ${msg}`);
      setAgentRunning(false);
      setPendingConfirm(null);
      unlisten();
    }
  }, [
    query,
    isAgentRunning,
    isConnected,
    currentApiKey,
    provider,
    model,
    activeConversationId,
    selectedFile,
  ]);

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      handleSubmit();
    }
  };

  const handleCancel = async () => {
    try {
      await invoke("cancel_agent");
    } catch {
      // ignore
    }
  };

  const handleRollback = async () => {
    try {
      await invoke("rollback_agent");
      setHasBackup(false);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      alert(`롤백 실패: ${msg}`);
    }
  };

  const filename = selectedFile
    ? selectedFile.replace(/\\/g, "/").split("/").pop()
    : null;

  const isReady = isConnected && !isAgentRunning && query.trim().length > 0;
  const hasConversation = activeConversationId !== null;

  return (
    <div className="flex flex-col h-full">
      {/* Top bar - file info */}
      {selectedFile && (
        <div className="shrink-0 px-6 py-2 border-b border-white/[0.06] flex items-center gap-3">
          <div className="flex items-center gap-2 flex-1 min-w-0">
            <FileText size={13} className="text-accent shrink-0" />
            <span className="text-[12px] text-text-secondary truncate">
              {filename}
            </span>
            <div
              className={`flex items-center gap-1.5 px-2 py-0.5 rounded-full text-[11px] ${
                isConnected
                  ? "bg-success/10 text-success"
                  : "bg-white/[0.04] text-text-tertiary"
              }`}
            >
              <div
                className={`w-1.5 h-1.5 rounded-full ${
                  isConnected ? "bg-success" : "bg-text-tertiary"
                }`}
              />
              {isConnected ? "연결됨" : "미연결"}
            </div>
          </div>
          <div className="flex items-center gap-1">
            {!isConnected && (
              <button
                onClick={handleConnect}
                disabled={isConnecting}
                className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] font-medium bg-accent text-white hover:bg-accent-hover transition-colors disabled:opacity-50"
              >
                {isConnecting ? (
                  <Loader2 size={11} className="animate-spin" />
                ) : (
                  <Plug size={11} />
                )}
                {isConnecting ? "연결 중" : "열기"}
              </button>
            )}
            {isConnected && (
              <button
                onClick={handlePreview}
                className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] text-text-tertiary hover:text-text-secondary hover:bg-white/[0.05] transition-colors"
              >
                <Eye size={11} />
                구조
              </button>
            )}
          </div>
        </div>
      )}

      {/* Preview overlay */}
      <AnimatePresence>
        {previewResult && (
          <motion.div
            initial={{ opacity: 0, height: 0 }}
            animate={{ opacity: 1, height: "auto" }}
            exit={{ opacity: 0, height: 0 }}
            className="shrink-0 overflow-hidden border-b border-white/[0.06]"
          >
            <div className="relative px-6 py-3 bg-black/20 max-h-48 overflow-y-auto">
              <pre className="font-mono text-[10px] text-text-secondary">
                {previewResult}
              </pre>
              <button
                onClick={() => setPreviewResult(null)}
                className="absolute top-2 right-4 w-6 h-6 flex items-center justify-center rounded text-text-tertiary hover:text-text hover:bg-white/10"
              >
                <X size={12} />
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Messages */}
      <div className="flex-1 overflow-y-auto">
        <div className="max-w-3xl mx-auto px-6 py-5">
          <AnimatePresence initial={false}>
            {messages.length === 0 ? (
              <motion.div
                key="empty"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="flex flex-col items-center justify-center min-h-[60vh] text-center"
              >
                <div className="w-12 h-12 rounded-2xl bg-accent/10 flex items-center justify-center mb-4">
                  <FileText size={22} className="text-accent" />
                </div>
                <h3 className="text-text text-[18px] font-semibold mb-2">
                  HWP AI 에이전트
                </h3>
                {!selectedFile ? (
                  <>
                    <p className="text-text-secondary text-[13px] max-w-sm leading-relaxed mb-5">
                      HWP 파일을 열고 AI에게 편집을 요청하세요.
                    </p>
                    <button
                      onClick={handleFileOpen}
                      className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-[13px] font-medium bg-accent text-white hover:bg-accent-hover transition-colors"
                    >
                      <Paperclip size={14} />
                      파일 열기
                    </button>
                  </>
                ) : !isConnected ? (
                  <>
                    <p className="text-text-secondary text-[13px] max-w-sm leading-relaxed mb-1">
                      <span className="text-accent font-medium">
                        {filename}
                      </span>
                      이(가) 선택되었습니다.
                    </p>
                    <p className="text-text-tertiary text-[12px] mb-5">
                      한글에서 문서를 열어 연결하세요.
                    </p>
                    <button
                      onClick={handleConnect}
                      disabled={isConnecting}
                      className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-[13px] font-medium bg-accent text-white hover:bg-accent-hover transition-colors disabled:opacity-50"
                    >
                      {isConnecting ? (
                        <Loader2 size={14} className="animate-spin" />
                      ) : (
                        <Plug size={14} />
                      )}
                      {isConnecting ? "연결 중..." : "한글에서 열기"}
                    </button>
                  </>
                ) : (
                  <>
                    <p className="text-text-secondary text-[13px] max-w-sm leading-relaxed mb-5">
                      문서가 연결되었습니다. 아래에 편집 요청을 입력하세요.
                    </p>
                    <div className="flex flex-wrap gap-2 justify-center max-w-sm">
                      {[
                        "표 구조 분석해줘",
                        "첫 번째 표 헤더 확인",
                        "모든 필드 목록 보여줘",
                      ].map((hint) => (
                        <button
                          key={hint}
                          onClick={() => setQuery(hint)}
                          className="px-3.5 py-1.5 rounded-full text-[12px] text-text-secondary bg-white/[0.04] border border-white/[0.06] hover:bg-white/[0.07] transition-colors"
                        >
                          {hint}
                        </button>
                      ))}
                    </div>
                  </>
                )}
              </motion.div>
            ) : (
              <div className="space-y-5">
                {messages.map((msg) => (
                  <MessageBubble key={msg.id} message={msg} />
                ))}
              </div>
            )}
          </AnimatePresence>
          <div ref={messagesEndRef} />
        </div>
      </div>

      {/* Input area */}
      <div className="shrink-0 px-4 pb-4 pt-2">
        <div className="max-w-3xl mx-auto">
          {/* File badge */}
          {selectedFile && !isConnected && messages.length > 0 && (
            <div className="flex items-center gap-2 mb-2 px-1">
              <span className="text-[11px] text-warning">
                문서 연결이 필요합니다
              </span>
              <button
                onClick={handleConnect}
                disabled={isConnecting}
                className="text-[11px] text-accent hover:underline"
              >
                연결하기
              </button>
            </div>
          )}

          <div className="bg-white/[0.04] border border-white/[0.06] rounded-2xl focus-within:border-accent/40 transition-colors">
            <textarea
              ref={textareaRef}
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              onKeyDown={handleKeyDown}
              placeholder={PLACEHOLDER}
              disabled={isAgentRunning}
              rows={1}
              className="w-full bg-transparent resize-none border-none outline-none px-4 pt-3 pb-1 text-[14px] text-text placeholder:text-text-tertiary leading-relaxed disabled:opacity-40"
              style={{ minHeight: "40px" }}
            />
            <div className="flex items-center justify-between px-3 pb-2.5">
              <div className="flex items-center gap-2">
                {/* File attach */}
                <button
                  onClick={handleFileOpen}
                  disabled={isAgentRunning}
                  className="w-8 h-8 flex items-center justify-center rounded-lg text-text-tertiary hover:text-text-secondary hover:bg-white/[0.05] transition-colors disabled:opacity-30"
                  title="파일 열기"
                >
                  <Paperclip size={15} />
                </button>

                {isAgentRunning ? (
                  <div className="flex items-center gap-2">
                    <span className="flex items-center gap-1.5 text-[12px] text-text-tertiary">
                      <span className="flex gap-0.5">
                        <span className="typing-dot" />
                        <span className="typing-dot" />
                        <span className="typing-dot" />
                      </span>
                      처리 중
                    </span>
                    <button
                      onClick={handleCancel}
                      className="flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] text-error/80 hover:text-error bg-error/[0.06] hover:bg-error/[0.12] transition-colors"
                      title="에이전트 중단"
                    >
                      <XCircle size={12} />
                      중단
                    </button>
                  </div>
                ) : (
                  <>
                    {tokenUsage && (
                      <span className="text-[10px] text-text-tertiary/60 font-mono">
                        {formatTokens(tokenUsage.total)} tokens
                      </span>
                    )}
                  </>
                )}
              </div>
              <div className="flex items-center gap-1.5">
                {hasBackup && !isAgentRunning && (
                  <button
                    onClick={handleRollback}
                    className="flex items-center gap-1 px-2.5 py-1 rounded-lg text-[11px] text-warning hover:bg-warning/[0.08] transition-colors"
                    title="마지막 에이전트 실행 전으로 되돌리기"
                  >
                    <Undo2 size={12} />
                    되돌리기
                  </button>
                )}
                {messages.length > 0 && !isAgentRunning && (
                  <button
                    onClick={clearMessages}
                    className="w-8 h-8 flex items-center justify-center rounded-lg text-text-tertiary hover:text-error hover:bg-white/[0.05] transition-colors"
                    title="대화 초기화"
                  >
                    <Trash2 size={14} />
                  </button>
                )}
                <button
                  onClick={handleSubmit}
                  disabled={!isReady}
                  className="w-8 h-8 flex items-center justify-center rounded-xl bg-accent text-white hover:bg-accent-hover transition-colors disabled:opacity-25 disabled:cursor-default"
                >
                  <ArrowUp size={16} strokeWidth={2.5} />
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
