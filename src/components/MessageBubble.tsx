import { motion } from "framer-motion";
import type { Message } from "../types";
import ToolCallCard from "./ToolCallCard";

interface Props {
  message: Message;
}

function TypingIndicator() {
  return (
    <div className="flex items-center gap-1.5 py-1">
      <span className="typing-dot" />
      <span className="typing-dot" />
      <span className="typing-dot" />
    </div>
  );
}

function TextContent({ text }: { text: string }) {
  const paragraphs = text.split("\n\n").filter(Boolean);
  return (
    <div>
      {paragraphs.map((para, i) => {
        if (para.startsWith("```")) {
          const lines = para.split("\n");
          const code = lines.slice(1, -1).join("\n");
          return (
            <pre key={i} className="bg-black/30 rounded-lg p-3 my-2 overflow-x-auto">
              <code className="font-mono text-[12px] text-text-secondary">{code}</code>
            </pre>
          );
        }
        if (para.startsWith("## ")) {
          return (
            <p key={i} className="font-semibold text-text text-[14px] mt-3 mb-1">
              {para.slice(3)}
            </p>
          );
        }
        if (para.startsWith("# ")) {
          return (
            <p key={i} className="font-semibold text-text text-[15px] mt-3 mb-1">
              {para.slice(2)}
            </p>
          );
        }
        return (
          <p key={i} className="text-[14px] leading-[1.6] text-text mb-1.5 last:mb-0">
            {renderInline(para)}
          </p>
        );
      })}
    </div>
  );
}

function renderInline(text: string): React.ReactNode {
  const parts = text.split(/(`[^`]+`)/g);
  return parts.map((part, i) =>
    part.startsWith("`") && part.endsWith("`") ? (
      <code key={i} className="font-mono text-[0.85em] bg-white/[0.06] rounded px-1 py-0.5 text-accent">
        {part.slice(1, -1)}
      </code>
    ) : (
      <span key={i}>{part}</span>
    )
  );
}

export default function MessageBubble({ message }: Props) {
  const isUser = message.role === "user";
  const isEmpty = message.contents.length === 0;

  return (
    <motion.div
      initial={{ opacity: 0, y: 6 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.2, ease: "easeOut" }}
      className={`flex gap-3 ${isUser ? "justify-end" : "justify-start"}`}
    >
      <div className={`max-w-[80%] ${isUser ? "items-end" : "items-start"} flex flex-col gap-0.5`}>
        <div
          className={`px-4 py-2.5 ${
            isUser
              ? "bg-accent text-white rounded-[18px] rounded-br-md"
              : "bg-white/[0.04] border border-white/[0.06] text-text rounded-[18px] rounded-bl-md"
          }`}
        >
          {isEmpty && !isUser ? (
            <TypingIndicator />
          ) : (
            message.contents.map((content, i) => {
              if (content.type === "text") {
                return <TextContent key={i} text={content.text} />;
              }
              if (content.type === "tool_call") {
                return <ToolCallCard key={i} tool={content.tool} />;
              }
              if (content.type === "thinking") {
                return (
                  <div key={i} className="text-text-tertiary text-[12px] italic flex items-center gap-2">
                    <span className="flex gap-0.5">
                      <span className="typing-dot" />
                      <span className="typing-dot" />
                      <span className="typing-dot" />
                    </span>
                    {content.text}
                  </div>
                );
              }
              return null;
            })
          )}
        </div>

        <span className={`text-[10px] text-text-tertiary px-2 ${isUser ? "text-right" : "text-left"}`}>
          {new Date(message.timestamp).toLocaleTimeString("ko-KR", {
            hour: "2-digit",
            minute: "2-digit",
          })}
        </span>
      </div>
    </motion.div>
  );
}
