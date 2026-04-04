/// tools.rs
/// HWP LLM 도구 정의 — OpenAI/Anthropic/Gemini 공통 스키마 + 공급자별 변환

use serde_json::{json, Value};

pub const SYSTEM_PROMPT: &str = r#"당신은 한글(HWP) 문서 편집 에이전트입니다.
사용자의 요청을 반드시 도구(tool)를 호출하여 실행하세요.

## 절대 규칙
- 문서를 수정하라는 요청에는 반드시 도구를 호출하여 실제로 수정하세요. 텍스트로 "수정했습니다"라고만 답하면 안 됩니다.
- 어떤 내용을 쓸지 설명만 하지 말고, 반드시 fill_table_data_matrix / fill_field_data / replace_text_patterns 등의 도구를 호출해서 실제 데이터를 문서에 기록하세요.
- 도구 호출 없이 작업 완료를 선언하는 것은 금지입니다.

## 효율적 탐색 원칙
1. 작업 시작: get_all_tables_overview() → 전체 표 목록 파악 (analyze_document_structure 대신 이것을 우선 사용)
2. 대상 표 확정 후에만 get_table_schema(table_index)로 세부 구조 확인
3. 수정: 누름틀은 fill_field_data(), 일반 텍스트는 replace_text_patterns(), 표 데이터는 fill_table_data_matrix() 사용
4. 모든 도구 호출이 완료된 후, 마지막에 간단히 결과를 요약

## 핵심 규칙
- get_all_tables_overview() 한 번으로 표 목록 파악 → 필요한 표만 get_table_schema() 호출
- 문서에서 실제 확인한 텍스트만 사용하세요. 추측 금지
- 대화 히스토리가 제공되면, 이전 대화 맥락("다음 표", "같은 방식으로" 등)을 반드시 참고하세요"#;

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
            "description": "2차원 배열 데이터를 표의 특정 행부터 채웁니다.",
            "parameters": {
                "type": "object",
                "properties": {
                    "table_index": {"type": "integer"},
                    "start_row": {"type": "integer"},
                    "matrix": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "string"}}
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
