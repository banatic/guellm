import { motion } from "framer-motion";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { Bot } from "lucide-react";
import type { Message } from "../types";
import ToolCallCard from "./ToolCallCard";

interface Props {
  message: Message;
}

function TypingIndicator() {
  return (
    <div className="flex items-center gap-2 py-1.5">
      <span className="flex gap-1">
        <span className="typing-dot" />
        <span className="typing-dot" />
        <span className="typing-dot" />
      </span>
      <span className="text-[12px] text-text-tertiary">생각하는 중...</span>
    </div>
  );
}

function TextContent({ text }: { text: string }) {
  return (
    <div className="prose-message">
      <ReactMarkdown remarkPlugins={[remarkGfm]}>{text}</ReactMarkdown>
    </div>
  );
}

export default function MessageBubble({ message }: Props) {
  const isUser = message.role === "user";
  const isEmpty = message.contents.length === 0;

  return (
    <motion.div
      initial={{ opacity: 0, y: 8 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.25, ease: "easeOut" }}
      className={`flex gap-3 ${isUser ? "justify-end" : "justify-start"}`}
    >
      {/* Assistant avatar */}
      {!isUser && (
        <div className="w-7 h-7 rounded-lg bg-gradient-to-br from-accent/20 to-accent/5 border border-accent/10 flex items-center justify-center shrink-0 mt-0.5">
          <Bot size={14} className="text-accent" />
        </div>
      )}

      <div
        className={`max-w-[80%] ${isUser ? "items-end" : "items-start"} flex flex-col gap-1`}
      >
        <div
          className={`px-4 py-3 ${
            isUser
              ? "bg-gradient-to-br from-accent to-blue-600 text-white rounded-2xl rounded-br-md shadow-glow"
              : "bg-white/[0.06] border border-white/[0.07] text-text rounded-2xl rounded-bl-md shadow-card"
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
                  <div
                    key={i}
                    className="text-text-tertiary text-[12px] italic flex items-center gap-2"
                  >
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

        <span
          className={`text-[10px] text-text-tertiary/60 px-2 ${isUser ? "text-right" : "text-left"}`}
        >
          {new Date(message.timestamp).toLocaleTimeString("ko-KR", {
            hour: "2-digit",
            minute: "2-digit",
          })}
        </span>
      </div>
    </motion.div>
  );
}
