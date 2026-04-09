/// tools.rs
/// HWP LLM 도구 정의 — OpenAI/Anthropic/Gemini 공통 스키마 + 공급자별 변환

use serde_json::{json, Value};

pub const SYSTEM_PROMPT: &str = r#"당신은 한글(HWP) 문서 편집 에이전트입니다. 반드시 도구를 호출해서 실행하세요.

## 도구 사용 순서
1. 텍스트 읽기: get_document_text() — 문서 전체 텍스트(표 포함). 특정 표 셀만 읽으려면 get_table_schema() 사용
2. 표 작업: get_all_tables_overview() → get_table_schema(index) → fill_table_data_matrix()
3. 수정: 누름틀=fill_field_data, 표 밖 텍스트=replace_text_patterns, 표 안=fill_table_data_matrix

## 특정 셀 내용 읽기/요약
- "탐구방법 및 실행계획 요약해줘" 같이 특정 셀 내용이 필요한 경우:
  1. get_all_tables_overview() → 표 인덱스 확인
  2. get_table_schema(index) → 해당 셀의 row/col 번호 확인
  3. get_cell_text(table_index, row, col) → 전체 내용 읽기(truncate 없음) → 요약/분석

## replace_text_patterns 사용 금지 조건 (위반 시 자동 거부)
- 표 셀 내용 수정 → fill_table_data_matrix 사용
- 80자 초과 검색어
- \n 포함 검색어 (HWP 단락 경계 검색 불가)

## 표 셀 수정 절차
1. get_all_tables_overview() → 표 인덱스·역할(role) 확인
2. get_table_schema(index) → 행 번호(row) 확인
3. fill_table_data_matrix(index, start_row, [["값"]], start_col=0 또는 1)
   - start_col=1: 라벨 열 건너뜀 (내용 열만 수정)

## 주의
- 이전 대화 맥락("다음 표", "같은 방식으로")을 반드시 참고
- 도구 호출 없이 "완료" 선언 금지
- 답으로 줄바꿈을 할 때에너는 /r/n 을 사용하세요."#;

/// 사용자 확인이 필요한 write(파괴적) 도구 목록
pub const WRITE_TOOLS: &[&str] = &[
    "fill_field_data",
    "replace_text_patterns",
    "set_checkbox_state",
    "fill_table_data_matrix",
    "sync_table_rows",
    "format_table_cells",
    "set_font_style",
    "insert_image_box",
    "append_page_from_template",
    "manage_page_visibility",
    "execute_raw_action",
];

/// 공통 스키마 도구 목록
pub fn hwp_tools() -> Vec<Value> {
    vec![
        json!({
            "name": "analyze_document_structure",
            "description": "문서 전체 구조(페이지 수, 표 개수, 필드 목록)를 분석합니다. 모든 작업 시작 시 먼저 호출하세요.",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }),
        json!({
            "name": "get_document_text",
            "description": "문서의 전체 텍스트 내용을 반환합니다. 표 내부 텍스트도 포함됩니다. 내용 요약, 번역, 분석 등 텍스트를 읽어야 할 때 사용하세요.",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }),
        json!({
            "name": "get_field_info",
            "description": "문서 내 모든 누름틀(필드)의 이름과 현재 값을 반환합니다.",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }),
        json!({
            "name": "get_all_tables_overview",
            "description": "문서 내 모든 표의 인덱스, 행/열 수, 헤더(첫 행)를 한 번에 반환합니다. get_table_schema 반복 호출 대신 이 도구를 먼저 사용하세요.",
            "parameters": {"type": "object", "properties": {}, "required": []}
        }),
        json!({
            "name": "get_table_schema",
            "description": "특정 표의 행/열 수, 헤더, 각 셀의 현재 텍스트를 반환합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "table_index": {"type": "integer", "description": "0부터 시작하는 표 인덱스"}
                },
                "required": ["table_index"]
            }
        }),
        json!({
            "name": "find_text_anchor",
            "description": "특정 텍스트가 문서에 존재하는지 확인합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "keyword": {"type": "string", "description": "찾을 텍스트"}
                },
                "required": ["keyword"]
            }
        }),
        json!({
            "name": "get_cell_text",
            "description": "특정 셀의 전체 텍스트를 반환합니다 (truncate 없음). get_table_schema는 60자로 잘리므로, 셀 내용을 읽거나 요약/번역할 때는 이 도구를 사용하세요. row/col 번호는 get_table_schema로 먼저 확인하세요.",
            "parameters": {
                "type": "object",
                "properties": {
                    "table_index": {"type": "integer"},
                    "row": {"type": "integer", "description": "0부터 시작하는 행 인덱스"},
                    "col": {"type": "integer", "description": "0부터 시작하는 열 인덱스"}
                },
                "required": ["table_index", "row", "col"]
            }
        }),
        json!({
            "name": "fill_field_data",
            "description": "누름틀(필드)에 데이터를 채웁니다. get_field_info()로 필드 이름을 먼저 확인하세요.",
            "parameters": {
                "type": "object",
                "properties": {
                    "data_map": {"type": "object", "description": "{'필드이름': '값'} 형태의 딕셔너리"}
                },
                "required": ["data_map"]
            }
        }),
        json!({
            "name": "replace_text_patterns",
            "description": "{{이름}}, [날짜] 등 텍스트 패턴을 실제 데이터로 일괄 치환합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mapping": {"type": "object", "description": "{'찾을텍스트': '바꿀텍스트'} 딕셔너리"}
                },
                "required": ["mapping"]
            }
        }),
        json!({
            "name": "set_checkbox_state",
            "description": "'label [ ]' 또는 '[ ] label' 형태의 체크박스 상태를 변경합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "label": {"type": "string"},
                    "is_checked": {"type": "boolean"}
                },
                "required": ["label", "is_checked"]
            }
        }),
        json!({
            "name": "insert_image_box",
            "description": "특정 텍스트 위치에 이미지를 삽입합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "anchor_text": {"type": "string"},
                    "image_path": {"type": "string"},
                    "size_mode": {"type": "string", "enum": ["fit", "original"]}
                },
                "required": ["image_path"]
            }
        }),
        json!({
            "name": "sync_table_rows",
            "description": "입력할 데이터 개수에 맞춰 표의 데이터 행을 추가/삭제합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "table_index": {"type": "integer"},
                    "data_count": {"type": "integer"},
                    "header_rows": {"type": "integer", "default": 1}
                },
                "required": ["table_index", "data_count"]
            }
        }),
        json!({
            "name": "fill_table_data_matrix",
            "description": "2차원 배열 데이터를 표의 특정 행/열부터 채웁니다. ⚠️ matrix 내부 배열의 각 원소 = 셀 1개. 한 셀에 여러 줄 입력 시 \\n 사용 (예: [[\"1줄\\n2줄\"]] = 셀 1개에 2단락). start_col로 라벨 열을 건너뛸 수 있습니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "table_index": {"type": "integer"},
                    "start_row": {"type": "integer", "description": "0부터 시작하는 행 인덱스"},
                    "start_col": {"type": "integer", "description": "0부터 시작하는 열 인덱스 (기본 0). 라벨 열을 건너뛰려면 1 이상"},
                    "matrix": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "string"}},
                        "description": "행 배열. 각 내부 배열 = 1행의 셀들. 내부 원소 하나 = 셀 하나. 셀 내 줄바꿈은 \\n."
                    },
                    "cell_delay": {"type": "number", "default": 0.3}
                },
                "required": ["table_index", "start_row", "matrix"]
            }
        }),
        json!({
            "name": "format_table_cells",
            "description": "선택된 셀의 배경색이나 테두리를 변경합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "table_index": {"type": "integer"},
                    "format_dict": {"type": "object"}
                },
                "required": ["table_index", "format_dict"]
            }
        }),
        json!({
            "name": "set_font_style",
            "description": "현재 선택 영역의 글자 서식을 설정합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "font_name": {"type": "string"},
                    "size_pt": {"type": "number"},
                    "bold": {"type": "boolean"},
                    "color_rgb": {"type": "integer"}
                },
                "required": []
            }
        }),
        json!({
            "name": "auto_fit_paragraph",
            "description": "자간을 좁혀 텍스트를 한 줄에 맞춥니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "decrease_count": {"type": "integer", "default": 3}
                },
                "required": []
            }
        }),
        json!({
            "name": "append_page_from_template",
            "description": "다른 HWP 파일을 현재 문서 끝에 이어 붙입니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string"}
                },
                "required": ["file_path"]
            }
        }),
        json!({
            "name": "manage_page_visibility",
            "description": "특정 페이지를 감추거나 표시합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "page_number": {"type": "integer"},
                    "action": {"type": "string", "enum": ["hide", "show"]}
                },
                "required": ["page_number", "action"]
            }
        }),
        json!({
            "name": "export_to_pdf",
            "description": "현재 문서를 PDF로 저장합니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "target_path": {"type": "string"}
                },
                "required": ["target_path"]
            }
        }),
        json!({
            "name": "execute_raw_action",
            "description": "ActionTable.md의 모든 HWP COM Action을 직접 실행합니다. 다른 도구로 처리 불가능한 특수 케이스에 사용하세요.",
            "parameters": {
                "type": "object",
                "properties": {
                    "action_id": {"type": "string"},
                    "params": {"type": "object"}
                },
                "required": ["action_id"]
            }
        }),
    ]
}

/// OpenAI function calling 형식으로 변환
pub fn to_openai_tools(tools: &[Value]) -> Value {
    let converted: Vec<Value> = tools
        .iter()
        .map(|t| {
            json!({
                "type": "function",
                "function": {
                    "name": t["name"],
                    "description": t["description"],
                    "parameters": t["parameters"]
                }
            })
        })
        .collect();
    json!(converted)
}

/// Anthropic tools 형식으로 변환
pub fn to_anthropic_tools(tools: &[Value]) -> Value {
    let converted: Vec<Value> = tools
        .iter()
        .map(|t| {
            json!({
                "name": t["name"],
                "description": t["description"],
                "input_schema": t["parameters"]
            })
        })
        .collect();
    json!(converted)
}

/// Gemini function declarations 형식으로 변환
pub fn to_gemini_tools(tools: &[Value]) -> Value {
    let declarations: Vec<Value> = tools
        .iter()
        .map(|t| {
            let params = &t["parameters"];
            // Gemini는 빈 properties 객체를 좋아하지 않으므로 처리
            let has_props = params["properties"]
                .as_object()
                .map(|p| !p.is_empty())
                .unwrap_or(false);
            if has_props {
                json!({
                    "name": t["name"],
                    "description": t["description"],
                    "parameters": params
                })
            } else {
                json!({
                    "name": t["name"],
                    "description": t["description"]
                })
            }
        })
        .collect();
    json!([{"functionDeclarations": declarations}])
}
