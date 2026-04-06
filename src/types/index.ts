export type Provider = "openai" | "gemini" | "anthropic";

export interface AppConfig {
  api_keys: Record<string, string>;
  last_provider: string;
  last_model: string;
  last_file: string | null;
  zoom?: number;
}

export type MessageRole = "user" | "assistant" | "system";

export interface ToolCallData {
  name: string;
  args: Record<string, unknown>;
  result?: string;
  status: "pending" | "running" | "done" | "error";
}

export type MessageContent =
  | { type: "text"; text: string }
  | { type: "tool_call"; tool: ToolCallData }
  | { type: "thinking"; text: string };

export interface Message {
  id: string;
  role: MessageRole;
  contents: MessageContent[];
  timestamp: number;
}

export interface Conversation {
  id: string;
  title: string;
  messages: Message[];
  file: string | null;
  createdAt: number;
  updatedAt: number;
}

// Tauri 이벤트 페이로드
export type AgentEvent =
  | { type: "toolCall"; name: string; args: Record<string, unknown> }
  | { type: "toolResult"; name: string; result: string }
  | { type: "toolConfirmRequest"; name: string; args: Record<string, unknown> }
  | { type: "llmThinking"; text: string }
  | { type: "finalResponse"; text: string }
  | { type: "error"; message: string }
  | {
      type: "tokenUsage";
      prompt_tokens: number;
      completion_tokens: number;
      total_tokens: number;
    };

export const DEFAULT_MODELS: Record<Provider, string> = {
  openai: "gpt-4o",
  gemini: "gemini-2.0-flash",
  anthropic: "claude-sonnet-4-6",
};

export const PROVIDER_LABELS: Record<Provider, string> = {
  openai: "OpenAI",
  gemini: "Google Gemini",
  anthropic: "Anthropic",
};
