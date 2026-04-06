import { create } from "zustand";
import {
  Message,
  Provider,
  ToolCallData,
  DEFAULT_MODELS,
  AppConfig,
  Conversation,
} from "../types";

function generateId() {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function deriveTitle(messages: Message[]): string {
  const first = messages.find((m) => m.role === "user");
  if (!first) return "새 채팅";
  const text = first.contents
    .filter((c): c is { type: "text"; text: string } => c.type === "text")
    .map((c) => c.text)
    .join(" ");
  return text.length > 30 ? text.slice(0, 30) + "..." : text || "새 채팅";
}

interface AppStore {
  // 대화 관리
  conversations: Conversation[];
  activeConversationId: string | null;
  createConversation: (file?: string | null) => string;
  switchConversation: (id: string) => void;
  deleteConversation: (id: string) => void;
  activeConversation: () => Conversation | null;

  // 파일
  selectedFile: string | null;
  setSelectedFile: (path: string | null) => void;

  // 연결 상태
  isConnected: boolean;
  setConnected: (v: boolean) => void;

  // 설정
  provider: Provider;
  model: string;
  apiKeys: Record<string, string>;
  setProvider: (p: Provider) => void;
  setModel: (m: string) => void;
  setApiKey: (provider: string, key: string) => void;
  currentApiKey: () => string;

  // 설정 모달
  settingsOpen: boolean;
  setSettingsOpen: (v: boolean) => void;

  // 화면 배율
  zoom: number;
  setZoom: (v: number) => void;

  // 채팅
  messages: Message[];
  isAgentRunning: boolean;
  addUserMessage: (text: string) => string;
  addAssistantMessage: () => string;
  appendToolCall: (msgId: string, tool: ToolCallData) => string;
  updateToolCall: (
    msgId: string,
    toolId: string,
    updates: Partial<ToolCallData>
  ) => void;
  appendTextToAssistant: (msgId: string, text: string) => void;
  setFinalResponse: (msgId: string, text: string) => void;
  clearMessages: () => void;
  setAgentRunning: (v: boolean) => void;

  // Human-in-the-Loop
  pendingConfirm: { name: string; args: Record<string, unknown> } | null;
  setPendingConfirm: (
    v: { name: string; args: Record<string, unknown> } | null
  ) => void;

  // Undo/Rollback
  hasBackup: boolean;
  setHasBackup: (v: boolean) => void;

  // Token usage
  tokenUsage: { prompt: number; completion: number; total: number } | null;
  setTokenUsage: (
    v: { prompt: number; completion: number; total: number } | null
  ) => void;

  // 설정 로드/저장
  loadConfig: (cfg: AppConfig) => void;
  toConfig: () => AppConfig;
}

// 대화 목록을 messages 배열과 동기화
function syncConversation(state: {
  conversations: Conversation[];
  activeConversationId: string | null;
  messages: Message[];
  selectedFile: string | null;
}): Partial<AppStore> {
  if (!state.activeConversationId) return {};
  return {
    conversations: state.conversations.map((c) =>
      c.id === state.activeConversationId
        ? {
            ...c,
            messages: state.messages,
            file: state.selectedFile,
            title: deriveTitle(state.messages),
            updatedAt: Date.now(),
          }
        : c
    ),
  };
}

export const useAppStore = create<AppStore>((set, get) => ({
  // 대화 관리
  conversations: [],
  activeConversationId: null,

  createConversation: (file = null) => {
    const s = get();
    // 현재 대화 저장
    const updated = s.activeConversationId
      ? s.conversations.map((c) =>
          c.id === s.activeConversationId
            ? {
                ...c,
                messages: s.messages,
                file: s.selectedFile,
                title: deriveTitle(s.messages),
                updatedAt: Date.now(),
              }
            : c
        )
      : s.conversations;

    const id = generateId();
    const conv: Conversation = {
      id,
      title: "새 채팅",
      messages: [],
      file,
      createdAt: Date.now(),
      updatedAt: Date.now(),
    };
    set({
      conversations: [conv, ...updated],
      activeConversationId: id,
      messages: [],
      selectedFile: file,
      isConnected: false,
      tokenUsage: null,
      hasBackup: false,
      pendingConfirm: null,
    });
    return id;
  },

  switchConversation: (id) => {
    const s = get();
    if (s.activeConversationId === id) return;
    // 현재 대화 저장
    const updated = s.activeConversationId
      ? s.conversations.map((c) =>
          c.id === s.activeConversationId
            ? {
                ...c,
                messages: s.messages,
                file: s.selectedFile,
                title: deriveTitle(s.messages),
                updatedAt: Date.now(),
              }
            : c
        )
      : s.conversations;

    const target = updated.find((c) => c.id === id);
    if (!target) return;
    set({
      conversations: updated,
      activeConversationId: id,
      messages: target.messages,
      selectedFile: target.file,
      isConnected: false,
      tokenUsage: null,
      hasBackup: false,
      pendingConfirm: null,
    });
  },

  deleteConversation: (id) => {
    const s = get();
    const filtered = s.conversations.filter((c) => c.id !== id);
    if (s.activeConversationId === id) {
      const next = filtered[0] || null;
      set({
        conversations: filtered,
        activeConversationId: next?.id || null,
        messages: next?.messages || [],
        selectedFile: next?.file || null,
        isConnected: false,
        tokenUsage: null,
        hasBackup: false,
      });
    } else {
      set({ conversations: filtered });
    }
  },

  activeConversation: () => {
    const s = get();
    return s.conversations.find((c) => c.id === s.activeConversationId) || null;
  },

  selectedFile: null,
  setSelectedFile: (path) => set({ selectedFile: path }),

  isConnected: false,
  setConnected: (v) => set({ isConnected: v }),

  provider: "anthropic",
  model: DEFAULT_MODELS["anthropic"],
  apiKeys: {},
  setProvider: (p) =>
    set(() => ({
      provider: p,
      model: DEFAULT_MODELS[p],
    })),
  setModel: (m) => set({ model: m }),
  setApiKey: (provider, key) =>
    set((s) => ({ apiKeys: { ...s.apiKeys, [provider]: key } })),
  currentApiKey: () => {
    const s = get();
    return s.apiKeys[s.provider] ?? "";
  },

  settingsOpen: false,
  setSettingsOpen: (v) => set({ settingsOpen: v }),

  zoom: parseFloat(localStorage.getItem("guellm_zoom") ?? "1"),
  setZoom: (v) => {
    localStorage.setItem("guellm_zoom", String(v));
    set({ zoom: v });
  },

  messages: [],
  isAgentRunning: false,

  addUserMessage: (text) => {
    const id = generateId();
    set((s) => {
      const newMessages = [
        ...s.messages,
        {
          id,
          role: "user" as const,
          contents: [{ type: "text" as const, text }],
          timestamp: Date.now(),
        },
      ];
      const convUpdate = s.activeConversationId
        ? {
            conversations: s.conversations.map((c) =>
              c.id === s.activeConversationId
                ? {
                    ...c,
                    messages: newMessages,
                    title: deriveTitle(newMessages),
                    updatedAt: Date.now(),
                  }
                : c
            ),
          }
        : {};
      return { messages: newMessages, ...convUpdate };
    });
    return id;
  },

  addAssistantMessage: () => {
    const id = generateId();
    set((s) => ({
      messages: [
        ...s.messages,
        {
          id,
          role: "assistant",
          contents: [],
          timestamp: Date.now(),
        },
      ],
    }));
    return id;
  },

  appendToolCall: (msgId, tool) => {
    const toolId = generateId();
    set((s) => ({
      messages: s.messages.map((m) =>
        m.id === msgId
          ? {
              ...m,
              contents: [
                ...m.contents,
                { type: "tool_call" as const, tool: { ...tool, id: toolId } },
              ],
            }
          : m
      ),
    }));
    return toolId;
  },

  updateToolCall: (msgId, toolCallName, updates) => {
    set((s) => ({
      messages: s.messages.map((m) =>
        m.id === msgId
          ? {
              ...m,
              contents: m.contents.map((c) =>
                c.type === "tool_call" && c.tool.name === toolCallName
                  ? { ...c, tool: { ...c.tool, ...updates } }
                  : c
              ),
            }
          : m
      ),
    }));
  },

  appendTextToAssistant: (msgId, text) => {
    set((s) => ({
      messages: s.messages.map((m) => {
        if (m.id !== msgId) return m;
        const last = m.contents[m.contents.length - 1];
        if (last?.type === "text") {
          return {
            ...m,
            contents: [
              ...m.contents.slice(0, -1),
              { type: "text" as const, text: last.text + text },
            ],
          };
        }
        return {
          ...m,
          contents: [...m.contents, { type: "text" as const, text }],
        };
      }),
    }));
  },

  setFinalResponse: (msgId, text) => {
    set((s) => {
      const newMessages = s.messages.map((m) =>
        m.id === msgId
          ? {
              ...m,
              contents: [
                ...m.contents.filter((c) => c.type === "tool_call"),
                { type: "text" as const, text },
              ],
            }
          : m
      );
      const convUpdate = s.activeConversationId
        ? {
            conversations: s.conversations.map((c) =>
              c.id === s.activeConversationId
                ? {
                    ...c,
                    messages: newMessages,
                    title: deriveTitle(newMessages),
                    updatedAt: Date.now(),
                  }
                : c
            ),
          }
        : {};
      return { messages: newMessages, ...convUpdate };
    });
  },

  clearMessages: () => set({ messages: [], tokenUsage: null }),
  setAgentRunning: (v) => set({ isAgentRunning: v }),

  pendingConfirm: null,
  setPendingConfirm: (v) => set({ pendingConfirm: v }),

  hasBackup: false,
  setHasBackup: (v) => set({ hasBackup: v }),

  tokenUsage: null,
  setTokenUsage: (v) => set({ tokenUsage: v }),

  loadConfig: (cfg) => {
    const provider = (cfg.last_provider as Provider) || "anthropic";
    const zoom = cfg.zoom ?? parseFloat(localStorage.getItem("guellm_zoom") ?? "1");
    localStorage.setItem("guellm_zoom", String(zoom));
    set({
      provider,
      model: cfg.last_model || DEFAULT_MODELS[provider],
      apiKeys: cfg.api_keys || {},
      selectedFile: cfg.last_file || null,
      zoom,
    });
  },

  toConfig: (): AppConfig => {
    const s = get();
    return {
      api_keys: s.apiKeys,
      last_provider: s.provider,
      last_model: s.model,
      last_file: s.selectedFile,
      zoom: s.zoom,
    };
  },
}));
