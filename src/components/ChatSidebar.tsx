import { useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import {
  Plus,
  MessageSquare,
  Settings,
  Trash2,
  PanelLeftClose,
  PanelLeft,
  FileText,
} from "lucide-react";
import { useAppStore } from "../store/useAppStore";

interface Props {
  collapsed: boolean;
  onToggle: () => void;
}

export default function ChatSidebar({ collapsed, onToggle }: Props) {
  const {
    conversations,
    activeConversationId,
    createConversation,
    switchConversation,
    deleteConversation,
    isAgentRunning,
    setSettingsOpen,
  } = useAppStore();

  const [hoveredId, setHoveredId] = useState<string | null>(null);

  const handleNew = () => {
    if (isAgentRunning) return;
    createConversation();
  };

  const handleDelete = (e: React.MouseEvent, id: string) => {
    e.stopPropagation();
    if (isAgentRunning) return;
    deleteConversation(id);
  };

  if (collapsed) {
    return (
      <div className="flex flex-col items-center py-3 gap-1 w-12 shrink-0 bg-[#181819] border-r border-white/[0.06]">
        <button
          onClick={onToggle}
          className="w-9 h-9 flex items-center justify-center rounded-lg text-text-tertiary hover:text-text-secondary hover:bg-white/[0.07] transition-all duration-150"
          title="사이드바 열기"
        >
          <PanelLeft size={18} />
        </button>
        <button
          onClick={handleNew}
          disabled={isAgentRunning}
          className="w-9 h-9 flex items-center justify-center rounded-lg text-text-tertiary hover:text-text-secondary hover:bg-white/[0.07] transition-all duration-150 disabled:opacity-30 mt-1"
          title="새 채팅"
        >
          <Plus size={18} />
        </button>
        <div className="flex-1" />
        <button
          onClick={() => setSettingsOpen(true)}
          className="w-9 h-9 flex items-center justify-center rounded-lg text-text-tertiary hover:text-text-secondary hover:bg-white/[0.07] transition-all duration-150"
          title="설정"
        >
          <Settings size={16} />
        </button>
      </div>
    );
  }

  return (
    <div className="flex flex-col w-[260px] shrink-0 bg-[#181819] border-r border-white/[0.06]">
      {/* Header */}
      <div className="flex items-center justify-between px-3 py-3">
        <button
          onClick={onToggle}
          className="w-8 h-8 flex items-center justify-center rounded-lg text-text-tertiary hover:text-text-secondary hover:bg-white/[0.07] transition-all duration-150"
          title="사이드바 닫기"
        >
          <PanelLeftClose size={18} />
        </button>
        <button
          onClick={handleNew}
          disabled={isAgentRunning}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-[13px] text-text-secondary hover:text-text hover:bg-white/[0.07] transition-all duration-150 disabled:opacity-30"
        >
          <Plus size={15} />
          새 채팅
        </button>
      </div>

      {/* Conversation list */}
      <div className="flex-1 overflow-y-auto px-2 pb-2">
        <AnimatePresence initial={false}>
          {conversations.length === 0 ? (
            <div className="px-3 py-8 text-center">
              <p className="text-[12px] text-text-tertiary">
                대화가 없습니다
              </p>
            </div>
          ) : (
            conversations.map((conv) => {
              const isActive = conv.id === activeConversationId;
              const isHovered = conv.id === hoveredId;
              const filename = conv.file
                ?.replace(/\\/g, "/")
                .split("/")
                .pop();

              return (
                <motion.button
                  key={conv.id}
                  initial={{ opacity: 0, x: -8 }}
                  animate={{ opacity: 1, x: 0 }}
                  exit={{ opacity: 0, x: -8 }}
                  transition={{ duration: 0.15 }}
                  onClick={() => switchConversation(conv.id)}
                  onMouseEnter={() => setHoveredId(conv.id)}
                  onMouseLeave={() => setHoveredId(null)}
                  className={`w-full flex items-start gap-2.5 px-3 py-2.5 rounded-xl text-left transition-all duration-150 mb-0.5 group relative ${
                    isActive
                      ? "bg-white/[0.08] text-text"
                      : "text-text-secondary hover:bg-white/[0.05]"
                  }`}
                >
                  {/* Active indicator bar */}
                  {isActive && (
                    <motion.div
                      layoutId="sidebar-active"
                      className="absolute left-0 top-1/2 -translate-y-1/2 w-[3px] h-4 rounded-r-full bg-accent"
                      transition={{ type: "spring", stiffness: 400, damping: 30 }}
                    />
                  )}
                  <MessageSquare
                    size={14}
                    className={`mt-0.5 shrink-0 ${isActive ? "opacity-70" : "opacity-40"}`}
                  />
                  <div className="flex-1 min-w-0">
                    <div className="text-[13px] truncate leading-tight">
                      {conv.title}
                    </div>
                    {filename && (
                      <div className="flex items-center gap-1 mt-1 text-[11px] text-text-tertiary truncate">
                        <FileText size={10} className="shrink-0" />
                        {filename}
                      </div>
                    )}
                  </div>
                  {(isActive || isHovered) && !isAgentRunning && (
                    <div
                      role="button"
                      tabIndex={0}
                      onClick={(e) => handleDelete(e, conv.id)}
                      onKeyDown={(e) => e.key === "Enter" && handleDelete(e as unknown as React.MouseEvent, conv.id)}
                      className="shrink-0 mt-0.5 w-6 h-6 flex items-center justify-center rounded-md text-text-tertiary hover:text-error hover:bg-error/[0.1] transition-all duration-150 cursor-pointer"
                    >
                      <Trash2 size={12} />
                    </div>
                  )}
                </motion.button>
              );
            })
          )}
        </AnimatePresence>
      </div>

      {/* Footer */}
      <div className="px-2 pb-3 pt-1 border-t border-white/[0.06]">
        <button
          onClick={() => setSettingsOpen(true)}
          className="w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-[13px] text-text-tertiary hover:text-text-secondary hover:bg-white/[0.07] transition-all duration-150"
        >
          <Settings size={14} />
          설정
        </button>
      </div>
    </div>
  );
}
