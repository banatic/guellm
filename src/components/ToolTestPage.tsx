import { useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  Play, ChevronDown, ChevronRight, Check, XCircle,
  Loader2, Search, Table2, FileText, Type, Image,
  FileOutput, Zap, Settings, FlaskConical,
} from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";

// ── 도구 정의 (tools.rs와 동기화) ──────────────────────────

interface ToolParam {
  type: string;
  description?: string;
  enum?: string[];
  items?: unknown;
  default?: unknown;
}

interface ToolDef {
  name: string;
  description: string;
  category: "read" | "write" | "diag";
  params: Record<string, ToolParam>;
  required: string[];
}

const TOOLS: ToolDef[] = [
  // ── Read ──
  {
    name: "analyze_document_structure",
    description: "문서 전체 구조(페이지 수, 표 개수, 필드 목록)를 분석합니다.",
    category: "read",
    params: {},
    required: [],
  },
  {
    name: "get_document_text",
    description: "문서의 전체 텍스트 내용을 반환합니다. 요약/번역/분석 시 사용.",
    category: "read",
    params: {},
    required: [],
  },
  {
    name: "get_field_info",
    description: "문서 내 모든 누름틀(필드)의 이름과 현재 값을 반환합니다.",
    category: "read",
    params: {},
    required: [],
  },
  {
    name: "get_document_fields",
    description: "HTML FieldStart:/FieldEnd: 파싱으로 문서 내 모든 누름틀을 감지합니다. 이름 없는 필드도 포함. placeholder 여부(미입력 = italic+red)를 함께 반환.",
    category: "read",
    params: {},
    required: [],
  },
  {
    name: "get_field_list",
    description: "GetFieldList COM 진단용 — 이름 있는 누름틀만 반환됩니다. 이름 없으면 빈 결과.",
    category: "read",
    params: {},
    required: [],
  },
  {
    name: "get_field_values",
    description: "이름 있는 누름틀의 현재 값을 반환합니다. 이름 없는 필드는 get_document_fields 사용.",
    category: "read",
    params: {},
    required: [],
  },
  {
    name: "get_all_tables_overview",
    description: "문서 내 모든 표의 인덱스, 행/열 수, 헤더를 한 번에 반환합니다.",
    category: "read",
    params: {},
    required: [],
  },
  {
    name: "get_table_schema",
    description: "특정 표의 행/열 수, 헤더, 각 셀의 현재 텍스트를 반환합니다.",
    category: "read",
    params: {
      table_index: { type: "integer", description: "0부터 시작하는 표 인덱스" },
    },
    required: ["table_index"],
  },
  {
    name: "get_cell_text",
    description: "특정 셀의 전체 텍스트를 반환합니다 (truncate 없음).",
    category: "read",
    params: {
      table_index: { type: "integer", description: "0부터 시작하는 표 인덱스" },
      row: { type: "integer", description: "0부터 시작하는 행 인덱스" },
      col: { type: "integer", description: "0부터 시작하는 열 인덱스" },
    },
    required: ["table_index", "row", "col"],
  },
  {
    name: "find_text_anchor",
    description: "특정 텍스트가 문서에 존재하는지 확인합니다.",
    category: "read",
    params: {
      keyword: { type: "string", description: "찾을 텍스트" },
    },
    required: ["keyword"],
  },
  // ── Write ──
  {
    name: "fill_field_data",
    description: "누름틀(필드)에 데이터를 채웁니다.",
    category: "write",
    params: {
      data_map: { type: "object", description: '{"필드이름": "값"} 형태의 딕셔너리' },
    },
    required: ["data_map"],
  },
  {
    name: "replace_text_patterns",
    description: "{{이름}}, [날짜] 등 텍스트 패턴을 실제 데이터로 일괄 치환합니다.",
    category: "write",
    params: {
      mapping: { type: "object", description: '{"찾을텍스트": "바꿀텍스트"} 딕셔너리' },
    },
    required: ["mapping"],
  },
  {
    name: "set_checkbox_state",
    description: "체크박스 상태를 변경합니다.",
    category: "write",
    params: {
      label: { type: "string", description: "체크박스 레이블" },
      is_checked: { type: "boolean", description: "true/false" },
    },
    required: ["label", "is_checked"],
  },
  {
    name: "fill_table_data_matrix",
    description: "2차원 배열 데이터를 표의 특정 행부터 채웁니다.",
    category: "write",
    params: {
      table_index: { type: "integer", description: "0부터 시작하는 표 인덱스" },
      start_row: { type: "integer", description: "시작 행 인덱스 (0부터)" },
      start_col: { type: "integer", description: "시작 열 인덱스 (기본 0). 라벨 열 건너뛰려면 1", default: 0 },
      matrix: { type: "array", description: '2차원 문자열 배열. 예: [["a","b"],["c","d"]]' },
      cell_delay: { type: "number", description: "셀 입력 지연(초), 기본 0.3", default: 0.3 },
    },
    required: ["table_index", "start_row", "matrix"],
  },
  {
    name: "sync_table_rows",
    description: "입력할 데이터 개수에 맞춰 표의 데이터 행을 추가/삭제합니다.",
    category: "write",
    params: {
      table_index: { type: "integer", description: "표 인덱스" },
      data_count: { type: "integer", description: "필요한 데이터 행 수" },
      header_rows: { type: "integer", description: "헤더 행 수 (기본 1)", default: 1 },
    },
    required: ["table_index", "data_count"],
  },
  {
    name: "format_table_cells",
    description: "셀 배경색이나 테두리를 변경합니다.",
    category: "write",
    params: {
      table_index: { type: "integer", description: "표 인덱스" },
      format_dict: { type: "object", description: "서식 딕셔너리" },
    },
    required: ["table_index", "format_dict"],
  },
  {
    name: "set_font_style",
    description: "현재 선택 영역의 글자 서식을 설정합니다.",
    category: "write",
    params: {
      font_name: { type: "string", description: "폰트 이름" },
      size_pt: { type: "number", description: "포인트 크기" },
      bold: { type: "boolean", description: "굵게" },
      color_rgb: { type: "integer", description: "RGB 색상 정수" },
    },
    required: [],
  },
  {
    name: "auto_fit_paragraph",
    description: "자간을 좁혀 텍스트를 한 줄에 맞춥니다.",
    category: "write",
    params: {
      decrease_count: { type: "integer", description: "감소 횟수 (기본 3)", default: 3 },
    },
    required: [],
  },
  {
    name: "append_page_from_template",
    description: "다른 HWP 파일을 현재 문서 끝에 이어 붙입니다.",
    category: "write",
    params: {
      file_path: { type: "string", description: "추가할 HWP 파일 경로" },
    },
    required: ["file_path"],
  },
  {
    name: "manage_page_visibility",
    description: "특정 페이지를 감추거나 표시합니다.",
    category: "write",
    params: {
      page_number: { type: "integer", description: "페이지 번호" },
      action: { type: "string", description: '"hide" 또는 "show"', enum: ["hide", "show"] },
    },
    required: ["page_number", "action"],
  },
  {
    name: "fill_field",
    description: "누름틀에 값을 설정합니다. 이름 없는 필드는 placeholder 텍스트로 찾아 ForwardFind+SetFieldText(\"\") 방식으로 채웁니다.",
    category: "write",
    params: {
      placeholder: { type: "string", description: "채울 필드의 현재 placeholder 텍스트 (예: '학년', '과목명')" },
      value: { type: "string", description: "설정할 값" },
    },
    required: ["placeholder", "value"],
  },
  {
    name: "set_field",
    description: "이름 있는 누름틀에 값을 설정합니다. SetFieldText(name, value) 직접 호출.",
    category: "write",
    params: {
      name:  { type: "string", description: "필드 이름 (get_document_fields로 확인)" },
      value: { type: "string", description: "설정할 텍스트 값" },
    },
    required: ["name", "value"],
  },
  {
    name: "export_to_pdf",
    description: "현재 문서를 PDF로 저장합니다.",
    category: "write",
    params: {
      target_path: { type: "string", description: "저장 경로 (예: C:\\output.pdf)" },
    },
    required: ["target_path"],
  },
  {
    name: "execute_raw_action",
    description: "HWP COM Action을 직접 실행합니다.",
    category: "write",
    params: {
      action_id: { type: "string", description: "액션 ID (예: FileNew, Undo)" },
      params: { type: "object", description: "액션 파라미터 (선택)" },
    },
    required: ["action_id"],
  },

  // ── 진단 (Diag) ──
  {
    name: "diag_raw_html",
    description:
      "GetTextFile(\"HTML\") 원시 결과를 반환합니다. 태그 통계 + 미리보기. HWP가 실제로 어떤 HTML을 생성하는지 확인.",
    category: "diag",
    params: {
      limit: { type: "integer", description: "미리보기 최대 문자 수 (기본 3000)", default: 3000 },
    },
    required: [],
  },
  {
    name: "diag_text_file_txt",
    description:
      "GetTextFile(\"TEXT\") 결과를 반환합니다. HTML 파싱 없는 대안 — 탭(→)으로 셀 구분. 이스케이프 시각화 포함.",
    category: "diag",
    params: {
      limit: { type: "integer", description: "미리보기 최대 문자 수 (기본 3000)", default: 3000 },
    },
    required: [],
  },
  {
    name: "diag_normalize_keyword",
    description:
      "html_contains_text 내부 공백정규화를 시뮬레이션합니다. false_positive_risk=true → ForwardFind 실패/wrap 다이얼로그 위험.",
    category: "diag",
    params: {
      keyword: { type: "string", description: "확인할 키워드 (예: '탐구방법 및 실행계획')" },
    },
    required: ["keyword"],
  },
  {
    name: "diag_cell_raw",
    description:
      "특정 셀의 원시 텍스트를 이스케이프 표기(\\n \\r \\t)로 반환합니다. HTML 파서가 셀 내 단락을 어떻게 처리하는지 확인.",
    category: "diag",
    params: {
      table_index: { type: "integer", description: "표 인덱스 (0부터)", default: 0 },
      row: { type: "integer", description: "행 인덱스 (0부터)", default: 0 },
      col: { type: "integer", description: "열 인덱스 (0부터)", default: 0 },
    },
    required: ["table_index", "row", "col"],
  },
  {
    name: "diag_html_table_extract",
    description:
      "특정 표의 원시 HTML 청크를 추출합니다. merged cell(colspan/rowspan), 다단락 셀, 중첩 태그 구조 확인.",
    category: "diag",
    params: {
      table_index: { type: "integer", description: "표 인덱스 (0부터)", default: 0 },
      limit: { type: "integer", description: "미리보기 최대 문자 수 (기본 2000)", default: 2000 },
    },
    required: ["table_index"],
  },
  {
    name: "probe_scan",
    description:
      "InitScan/GetText 실측 — 각 이벤트의 state 값, text, CurFieldName(누름틀 감지), CurCtrl.CtrlID를 기록합니다. HTML 파싱 대체 구현 전 API 동작 확인용.",
    category: "diag",
    params: {
      max_events: { type: "integer", description: "수집할 최대 이벤트 수 (기본 300)", default: 300 },
    },
    required: [],
  },
  {
    name: "diag_initscan_gettext",
    description:
      "InitScan → GetText → ReleaseScan COM API를 테스트합니다. VT_BYREF output param 지원 여부 확인 — 지원되면 HTML 의존도 제거 가능.",
    category: "diag",
    params: {},
    required: [],
  },
  {
    name: "diag_get_pos",
    description:
      "hwp.GetPos() VT_BYREF 테스트. 현재 커서 좌표(list, para, pos)를 output 파라미터로 수신. 비영값이면 VT_BYREF 정상 작동.",
    category: "diag",
    params: {},
    required: [],
  },
  {
    name: "diag_key_indicator",
    description:
      "hwp.KeyIndicator() VT_BYREF 테스트. 커서가 표 안이면 ctrl_name='표' 반환. VT_BYREF 동작 여부 확인.",
    category: "diag",
    params: {},
    required: [],
  },
  {
    name: "diag_set_pos",
    description:
      "hwp.SetPos(list, para, pos) 테스트. diag_get_pos로 얻은 좌표를 입력해 커서 이동 확인. before/after 좌표 비교.",
    category: "diag",
    params: {
      list: { type: "integer", description: "list 좌표 (GetPos로 얻은 값)", default: 0 },
      para: { type: "integer", description: "para 좌표 (GetPos로 얻은 값)", default: 0 },
      pos:  { type: "integer", description: "pos 좌표 (GetPos로 얻은 값)", default: 0 },
    },
    required: ["list", "para", "pos"],
  },
  {
    name: "diag_initscan_1param",
    description:
      "InitScan 파라미터 개수 실험: 1개(0x37) vs 0개. 어떤 호출이 성공하는지 확인해 올바른 시그니처 파악.",
    category: "diag",
    params: {},
    required: [],
  },
  {
    name: "diag_scan_table_positions",
    description:
      "InitScan → GetText+GetPos 루프로 각 단락/셀의 위치 좌표(list,para,pos)를 수집. 최대 30개 항목 미리보기.",
    category: "diag",
    params: {},
    required: [],
  },
  {
    name: "diag_phys_structure",
    description:
      "parse_physical_tables 결과와 physical_cell_offset 계산값을 노출. navigation 계산이 맞는지 확인.",
    category: "diag",
    params: {
      table_index: { type: "integer", description: "표 인덱스", default: 0 },
      target_row:  { type: "integer", description: "확인할 행 (visual)", default: 4 },
      target_col:  { type: "integer", description: "확인할 열 (visual)", default: 2 },
    },
    required: [],
  },
  {
    name: "diag_table_cell_walk",
    description:
      "TableRightCell + GetPos로 표의 모든 물리 셀 좌표를 수집. SetPos 기반 navigation의 핵심 데이터. InitScan 불필요.",
    category: "diag",
    params: {
      table_index: { type: "integer", description: "0부터 시작하는 표 인덱스", default: 0 },
    },
    required: [],
  },
  {
    name: "diag_nav",
    description:
      "MoveDocBegin 후 MoveNextChar를 steps회 반복하며 GetPos + CurFieldName을 기록. 커서 이동 여부, 누름틀 감지 여부 확인.",
    category: "diag",
    params: {
      steps: { type: "integer", description: "이동 횟수 (기본 20)", default: 20 },
    },
    required: [],
  },
];

const TOOL_ICONS: Record<string, React.ReactNode> = {
  analyze_document_structure: <FileText size={13} />,
  get_document_text: <FileText size={13} />,
  get_field_info: <Type size={13} />,
  get_document_fields: <Type size={13} />,
  get_field_list: <Type size={13} />,
  get_field_values: <Type size={13} />,
  fill_field: <Type size={13} />,
  set_field: <Type size={13} />,
  get_all_tables_overview: <Table2 size={13} />,
  get_table_schema: <Table2 size={13} />,
  get_cell_text: <Table2 size={13} />,
  find_text_anchor: <Search size={13} />,
  fill_field_data: <Type size={13} />,
  replace_text_patterns: <Search size={13} />,
  set_checkbox_state: <Check size={13} />,
  insert_image_box: <Image size={13} />,
  sync_table_rows: <Table2 size={13} />,
  fill_table_data_matrix: <Table2 size={13} />,
  format_table_cells: <Settings size={13} />,
  set_font_style: <Type size={13} />,
  auto_fit_paragraph: <Type size={13} />,
  append_page_from_template: <FileText size={13} />,
  manage_page_visibility: <FileText size={13} />,
  export_to_pdf: <FileOutput size={13} />,
  execute_raw_action: <Zap size={13} />,
  // diag
  probe_scan: <FlaskConical size={13} />,
  diag_raw_html: <FlaskConical size={13} />,
  diag_text_file_txt: <FlaskConical size={13} />,
  diag_normalize_keyword: <FlaskConical size={13} />,
  diag_cell_raw: <FlaskConical size={13} />,
  diag_html_table_extract: <FlaskConical size={13} />,
  diag_initscan_gettext: <FlaskConical size={13} />,
  diag_get_pos: <FlaskConical size={13} />,
  diag_key_indicator: <FlaskConical size={13} />,
  diag_set_pos: <FlaskConical size={13} />,
  diag_initscan_1param: <FlaskConical size={13} />,
  diag_scan_table_positions: <FlaskConical size={13} />,
  diag_phys_structure: <FlaskConical size={13} />,
  diag_table_cell_walk: <FlaskConical size={13} />,
  diag_nav: <FlaskConical size={13} />,
};

function categoryColor(cat: ToolDef["category"]) {
  if (cat === "write") return { text: "text-warning", badge: "bg-warning/15 text-warning", icon: "text-warning/60" };
  if (cat === "diag")  return { text: "text-purple-400", badge: "bg-purple-500/15 text-purple-400", icon: "text-purple-400/60" };
  return { text: "text-accent", badge: "bg-accent/15 text-accent", icon: "text-text-tertiary" };
}

function categoryLabel(cat: ToolDef["category"]) {
  if (cat === "write") return "쓰기";
  if (cat === "diag")  return "진단";
  return "읽기";
}

// ── 기본 args JSON 생성 ────────────────────────────────────

function buildDefaultArgs(tool: ToolDef): string {
  if (Object.keys(tool.params).length === 0) return "{}";
  const obj: Record<string, unknown> = {};
  for (const [key, param] of Object.entries(tool.params)) {
    if (param.default !== undefined) {
      obj[key] = param.default;
    } else if (param.type === "integer" || param.type === "number") {
      obj[key] = 0;
    } else if (param.type === "boolean") {
      obj[key] = false;
    } else if (param.type === "array") {
      obj[key] = [];
    } else if (param.type === "object") {
      obj[key] = {};
    } else if (param.enum) {
      obj[key] = param.enum[0];
    } else {
      obj[key] = "";
    }
  }
  return JSON.stringify(obj, null, 2);
}

// ── 컴포넌트 ──────────────────────────────────────────────

export default function ToolTestPage() {
  const [selected, setSelected] = useState<ToolDef | null>(null);
  const [argsJson, setArgsJson] = useState("{}");
  const [argsError, setArgsError] = useState<string | null>(null);
  const [running, setRunning] = useState(false);
  const [result, setResult] = useState<{ ok: boolean; text: string } | null>(null);
  const [filter, setFilter] = useState<"all" | "read" | "write" | "diag">("all");
  const [expandedTool, setExpandedTool] = useState<string | null>(null);

  const visibleTools = TOOLS.filter(
    (t) => filter === "all" || t.category === filter
  );

  function selectTool(tool: ToolDef) {
    setSelected(tool);
    setArgsJson(buildDefaultArgs(tool));
    setArgsError(null);
    setResult(null);
  }

  function validateArgs(): Record<string, unknown> | null {
    try {
      const parsed = JSON.parse(argsJson);
      setArgsError(null);
      return parsed;
    } catch (e) {
      setArgsError(`JSON 파싱 오류: ${(e as Error).message}`);
      return null;
    }
  }

  async function runTool() {
    if (!selected) return;
    const args = validateArgs();
    if (!args) return;

    setRunning(true);
    setResult(null);
    try {
      const res = await invoke<string>("call_tool", { name: selected.name, args });
      setResult({ ok: true, text: res });
    } catch (e) {
      setResult({ ok: false, text: String(e) });
    } finally {
      setRunning(false);
    }
  }

  return (
    <div className="flex h-full min-h-0 overflow-hidden">
      {/* ── 왼쪽: 도구 목록 ── */}
      <div className="w-64 shrink-0 flex flex-col border-r border-white/[0.06] bg-[#1a1a1c] overflow-hidden">
        {/* 필터 탭 */}
        <div className="flex gap-1 p-2 border-b border-white/[0.06]">
          {(["all", "read", "write", "diag"] as const).map((f) => (
            <button
              key={f}
              onClick={() => setFilter(f)}
              className={`flex-1 text-[11px] py-1 rounded-md font-medium transition-all ${
                filter === f
                  ? f === "diag"
                    ? "bg-purple-500/20 text-purple-400"
                    : "bg-accent/20 text-accent"
                  : "text-text-tertiary hover:text-text-secondary hover:bg-white/[0.05]"
              }`}
            >
              {f === "all" ? "전체" : f === "read" ? "읽기" : f === "write" ? "쓰기" : "진단"}
            </button>
          ))}
        </div>

        {/* 도구 목록 */}
        <div className="flex-1 overflow-y-auto p-1.5 space-y-0.5">
          {visibleTools.map((tool) => (
            <button
              key={tool.name}
              onClick={() => selectTool(tool)}
              className={`w-full text-left px-2.5 py-2 rounded-lg flex items-start gap-2 transition-all group ${
                selected?.name === tool.name
                  ? tool.category === "diag"
                    ? "bg-purple-500/10 text-text-primary"
                    : "bg-accent/15 text-text-primary"
                  : "text-text-secondary hover:bg-white/[0.05] hover:text-text-primary"
              }`}
            >
              <span
                className={`mt-0.5 shrink-0 ${
                  selected?.name === tool.name
                    ? categoryColor(tool.category).text
                    : categoryColor(tool.category).icon
                }`}
              >
                {TOOL_ICONS[tool.name] ?? <Zap size={13} />}
              </span>
              <div className="min-w-0">
                <div className="text-[12px] font-medium truncate">{tool.name}</div>
                {tool.category !== "read" && (
                  <span className={`text-[10px] font-medium ${categoryColor(tool.category).text} opacity-70`}>
                    {categoryLabel(tool.category)}
                  </span>
                )}
              </div>
            </button>
          ))}
        </div>
      </div>

      {/* ── 오른쪽: 상세 / 실행 ── */}
      <div className="flex-1 min-w-0 flex flex-col overflow-hidden">
        {!selected ? (
          <div className="flex-1 flex items-center justify-center text-text-tertiary text-[13px]">
            왼쪽에서 테스트할 도구를 선택하세요
          </div>
        ) : (
          <div className="flex-1 flex flex-col overflow-y-auto p-4 gap-4">
            {/* 도구 헤더 */}
            <div className="flex items-start gap-3">
              <span className={`mt-0.5 ${categoryColor(selected.category).text}`}>
                {TOOL_ICONS[selected.name] ?? <Zap size={16} />}
              </span>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-[15px] font-semibold text-text-primary font-mono">
                    {selected.name}
                  </span>
                  <span className={`text-[10px] px-1.5 py-0.5 rounded-md font-medium ${categoryColor(selected.category).badge}`}>
                    {categoryLabel(selected.category)}
                  </span>
                </div>
                <p className="text-[12px] text-text-secondary mt-0.5">{selected.description}</p>
              </div>
            </div>

            {/* 파라미터 설명 */}
            {Object.keys(selected.params).length > 0 && (
              <div className="rounded-lg border border-white/[0.07] bg-white/[0.02] overflow-hidden">
                <button
                  onClick={() =>
                    setExpandedTool(expandedTool === selected.name ? null : selected.name)
                  }
                  className="w-full flex items-center justify-between px-3 py-2 text-[12px] text-text-secondary hover:text-text-primary transition-colors"
                >
                  <span className="font-medium">파라미터 설명</span>
                  {expandedTool === selected.name ? (
                    <ChevronDown size={13} />
                  ) : (
                    <ChevronRight size={13} />
                  )}
                </button>
                <AnimatePresence>
                  {expandedTool === selected.name && (
                    <motion.div
                      initial={{ height: 0 }}
                      animate={{ height: "auto" }}
                      exit={{ height: 0 }}
                      className="overflow-hidden"
                    >
                      <div className="px-3 pb-3 space-y-1.5 border-t border-white/[0.07]">
                        {Object.entries(selected.params).map(([key, param]) => (
                          <div key={key} className="flex gap-2 text-[11.5px]">
                            <span className="font-mono text-accent shrink-0">{key}</span>
                            <span className="text-text-tertiary">
                              <span className="text-text-secondary/60">{param.type}</span>
                              {selected.required.includes(key) && (
                                <span className="text-error ml-1">*</span>
                              )}
                              {param.description && (
                                <span className="ml-1">— {param.description}</span>
                              )}
                            </span>
                          </div>
                        ))}
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            )}

            {/* Args JSON 에디터 */}
            <div className="flex flex-col gap-1.5">
              <div className="flex items-center justify-between">
                <label className="text-[12px] font-medium text-text-secondary">
                  인수 (JSON)
                </label>
                {argsError && (
                  <span className="text-[11px] text-error flex items-center gap-1">
                    <XCircle size={11} />
                    {argsError}
                  </span>
                )}
              </div>
              <textarea
                value={argsJson}
                onChange={(e) => {
                  setArgsJson(e.target.value);
                  setArgsError(null);
                }}
                spellCheck={false}
                rows={Object.keys(selected.params).length > 0 ? 8 : 3}
                className={`w-full font-mono text-[12px] bg-[#161617] border rounded-lg px-3 py-2.5 text-text-primary resize-none outline-none transition-colors ${
                  argsError
                    ? "border-error/50 focus:border-error"
                    : "border-white/[0.08] focus:border-accent/50"
                }`}
              />
            </div>

            {/* 실행 버튼 */}
            <button
              onClick={runTool}
              disabled={running}
              className={`flex items-center justify-center gap-2 py-2.5 px-4 rounded-lg text-[13px] font-medium transition-all disabled:opacity-50 ${
                selected.category === "write"
                  ? "bg-warning/20 hover:bg-warning/30 text-warning"
                  : selected.category === "diag"
                  ? "bg-purple-500/20 hover:bg-purple-500/30 text-purple-400"
                  : "bg-accent/20 hover:bg-accent/30 text-accent"
              }`}
            >
              {running ? (
                <>
                  <Loader2 size={14} className="animate-spin" />
                  실행 중...
                </>
              ) : (
                <>
                  <Play size={14} />
                  {selected.name}() 실행
                </>
              )}
            </button>

            {/* 결과 */}
            <AnimatePresence>
              {result && (
                <motion.div
                  initial={{ opacity: 0, y: 6 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0 }}
                  className="flex flex-col gap-1.5"
                >
                  <div
                    className={`flex items-center gap-1.5 text-[12px] font-medium ${
                      result.ok ? "text-success" : "text-error"
                    }`}
                  >
                    {result.ok ? <Check size={13} /> : <XCircle size={13} />}
                    {result.ok ? "성공" : "오류"}
                  </div>
                  <pre className="bg-[#161617] border border-white/[0.07] rounded-lg px-3 py-2.5 text-[11.5px] text-text-secondary font-mono overflow-auto max-h-96 whitespace-pre-wrap break-words">
                    {(() => {
                      try {
                        return JSON.stringify(JSON.parse(result.text), null, 2);
                      } catch {
                        return result.text;
                      }
                    })()}
                  </pre>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        )}
      </div>
    </div>
  );
}
