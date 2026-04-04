import { useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import {
  ChevronDown, ChevronRight, Check, XCircle, Loader2,
  Search, Table2, FileText, Type, Image, FileOutput, Zap, Settings,
  ShieldCheck, ShieldX,
} from "lucide-react";
import { invoke } from "@tauri-apps/api/core";
import type { ToolCallData } from "../types";

interface Props {
  tool: ToolCallData;
}

const TOOL_ICONS: Record<string, React.ReactNode> = {
  analyze_document_structure: <FileText size={12} />,
  get_field_info: <Type size={12} />,
  get_all_tables_overview: <Table2 size={12} />,
  get_table_schema: <Table2 size={12} />,
  find_text_anchor: <Search size={12} />,
  fill_field_data: <Type size={12} />,
  replace_text_patterns: <Search size={12} />,
  set_checkbox_state: <Check size={12} />,
  insert_image_box: <Image size={12} />,
  sync_table_rows: <Table2 size={12} />,
  fill_table_data_matrix: <Table2 size={12} />,
  format_table_cells: <Settings size={12} />,
  set_font_style: <Type size={12} />,
  auto_fit_paragraph: <Type size={12} />,
  append_page_from_template: <FileText size={12} />,
  manage_page_visibility: <FileText size={12} />,
  export_to_pdf: <FileOutput size={12} />,
  execute_raw_action: <Zap size={12} />,
};

function formatArgs(args: Record<string, unknown>): string {
  if (!args || Object.keys(args).length === 0) return "";
  const parts = Object.entries(args)
    .slice(0, 2)
    .map(([k, v]) => {
      const sv = JSON.stringify(v);
      return `${k}: ${sv.length > 30 ? sv.slice(0, 30) + "..." : sv}`;
    });
  return parts.join(", ");
}

export default function ToolCallCard({ tool }: Props) {
  const [expanded, setExpanded] = useState(tool.status === "pending");
  const icon = TOOL_ICONS[tool.name] ?? <Zap size={12} />;

  const handleConfirm = async (approved: boolean) => {
    try {
      await invoke("confirm_tool", { approved });
    } catch {
      // ignore
    }
  };

  const statusIcon =
    tool.status === "pending" ? (
      <ShieldCheck size={12} className="text-warning animate-pulse" />
    ) : tool.status === "running" ? (
      <Loader2 size={12} className="text-accent animate-spin" />
    ) : tool.status === "done" ? (
      <Check size={12} className="text-success" />
    ) : tool.status === "error" ? (
      <XCircle size={12} className="text-error" />
    ) : null;

  return (
    <motion.div
      initial={{ opacity: 0, y: 4 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.15 }}
      className={`my-1.5 rounded-lg border transition-colors ${
        tool.status === "pending"
          ? "border-warning/30 bg-warning/[0.06]"
          : tool.status === "running"
          ? "border-accent/20 bg-accent/[0.04]"
          : tool.status === "error"
          ? "border-error/20 bg-error/[0.04]"
          : tool.status === "done"
          ? "border-success/10 bg-white/[0.02]"
          : "border-white/[0.06] bg-white/[0.02]"
      }`}
    >
      <button
        onClick={() => setExpanded((e) => !e)}
        className="w-full flex items-center gap-2 px-3 py-2 text-left"
      >
        <span className="text-text-tertiary">{icon}</span>
        <div className="flex-1 min-w-0">
          <span className="text-[12px] font-medium text-text font-mono">
            {tool.name}
          </span>
          {!expanded && formatArgs(tool.args) && (
            <span className="text-[11px] text-text-tertiary font-mono ml-1.5 truncate">
              {formatArgs(tool.args)}
            </span>
          )}
        </div>
        <div className="flex items-center gap-1.5 shrink-0">
          {statusIcon}
          {expanded ? (
            <ChevronDown size={11} className="text-text-tertiary" />
          ) : (
            <ChevronRight size={11} className="text-text-tertiary" />
          )}
        </div>
      </button>

      <AnimatePresence>
        {expanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: "auto", opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.15 }}
            className="overflow-hidden"
          >
            <div className="px-3 pb-2.5 space-y-2 border-t border-white/[0.04] pt-2">
              {Object.keys(tool.args).length > 0 && (
                <div>
                  <div className="text-[10px] text-text-tertiary mb-1 font-medium">입력</div>
                  <pre className="font-mono text-[11px] text-text-secondary bg-black/20 rounded-md p-2 overflow-x-auto whitespace-pre-wrap break-all max-h-32">
                    {JSON.stringify(tool.args, null, 2)}
                  </pre>
                </div>
              )}

              {/* Human-in-the-Loop: 승인/거부 버튼 */}
              {tool.status === "pending" && (
                <div className="flex items-center gap-2 py-1">
                  <span className="text-[11px] text-warning font-medium flex-1">
                    이 도구를 실행할까요?
                  </span>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleConfirm(true);
                    }}
                    className="flex items-center gap-1 px-3 py-1.5 rounded-md text-[11px] font-medium text-white bg-success/80 hover:bg-success transition-colors"
                  >
                    <ShieldCheck size={12} />
                    승인
                  </button>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      handleConfirm(false);
                    }}
                    className="flex items-center gap-1 px-3 py-1.5 rounded-md text-[11px] font-medium text-white bg-error/80 hover:bg-error transition-colors"
                  >
                    <ShieldX size={12} />
                    거부
                  </button>
                </div>
              )}

              {tool.result && (
                <div>
                  <div className="text-[10px] text-text-tertiary mb-1 font-medium">결과</div>
                  <pre
                    className={`font-mono text-[11px] rounded-md p-2 overflow-x-auto whitespace-pre-wrap break-all max-h-48 ${
                      tool.status === "error"
                        ? "bg-error/[0.06] text-error"
                        : "bg-black/20 text-success"
                    }`}
                  >
                    {tool.result}
                  </pre>
                </div>
              )}
              {tool.status === "running" && !tool.result && (
                <div className="flex items-center gap-2 py-1">
                  <span className="flex gap-0.5">
                    <span className="typing-dot" />
                    <span className="typing-dot" />
                    <span className="typing-dot" />
                  </span>
                  <span className="text-[11px] text-text-tertiary">실행 중...</span>
                </div>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </motion.div>
  );
}
