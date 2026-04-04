"""
llm_client.py
LLM API 클라이언트 (BYOK) - OpenAI, Google Gemini, Anthropic Claude 지원

【에이전틱 Tool Use 방식】
LLM이 HWP 도구를 직접 호출하며 문서를 탐색·수정합니다.
단방향 JSON 생성 대신 다중 턴 도구 호출 루프를 사용합니다.
"""
import json
import re
from typing import Callable, Literal

Provider = Literal["openai", "gemini", "anthropic"]

DEFAULT_MODELS: dict[str, str] = {
    "openai": "gpt-4o",
    "gemini": "gemini-2.0-flash",
    "anthropic": "claude-sonnet-4-6",
}

SYSTEM_PROMPT = """당신은 한글(HWP) 문서 편집 전문가입니다.
사용자의 요청에 따라 제공된 도구(tool)를 사용해 문서를 분석하고 수정하세요.

## 효율적 탐색 원칙 (턴 절약 필수)
1. 작업 시작: analyze_document_structure() → 전체 구조 파악
2. 표 탐색: get_table_schema를 반복 호출하지 마세요. 반드시 get_all_tables_overview()를 먼저 호출해 어떤 표가 대상인지 한 번에 파악하세요.
3. 대상 표 확정 후에만 get_table_schema(table_index)로 세부 구조 확인.
4. 수정: 누름틀은 fill_field_data(), 일반 텍스트는 replace_text_patterns(), 표 데이터는 fill_table_data_matrix() 사용.
5. 모든 수정이 완료되면 자연어로 결과를 요약하여 응답하세요.

## 핵심 규칙
- get_all_tables_overview() 한 번으로 표 목록 파악 → 필요한 표만 get_table_schema() 호출.
- 문서에서 실제 확인한 텍스트만 사용하세요. 추측 금지.
- 빈 셀에 fill_table_data_matrix를 사용할 때는 플레이스홀더 텍스트가 먼저 있어야 합니다.
- 모든 작업 완료 후 자연어로 결과를 요약하여 응답하세요.
"""

# ──────────────────────────────────────────────────────────────
# Tool (Function) 정의 — OpenAI/Gemini/Anthropic 공통 스키마
# ──────────────────────────────────────────────────────────────

HWP_TOOLS: list[dict] = [
    {
        "name": "analyze_document_structure",
        "description": "문서 전체 구조(페이지 수, 표 개수, 필드 목록)를 분석합니다. 모든 작업 시작 시 먼저 호출하세요.",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_field_info",
        "description": "문서 내 모든 누름틀(필드)의 이름과 현재 값을 반환합니다.",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_all_tables_overview",
        "description": (
            "문서 내 모든 표의 인덱스, 행/열 수, 헤더(첫 행)를 한 번에 반환합니다. "
            "어떤 표를 수정해야 할지 파악할 때 get_table_schema를 반복 호출하는 대신 이 도구를 먼저 사용하세요."
        ),
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "name": "get_table_schema",
        "description": "특정 표의 행/열 수, 헤더, 각 셀의 현재 텍스트를 반환합니다. get_all_tables_overview로 대상 표를 특정한 후 호출하세요.",
        "parameters": {
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "0부터 시작하는 표 인덱스 (첫 번째 표=0)",
                }
            },
            "required": ["table_index"],
        },
    },
    {
        "name": "find_text_anchor",
        "description": "특정 텍스트가 문서에 존재하는지 확인합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "keyword": {"type": "string", "description": "찾을 텍스트"}
            },
            "required": ["keyword"],
        },
    },
    {
        "name": "fill_field_data",
        "description": "누름틀(필드)에 데이터를 채웁니다. get_field_info()로 필드 이름을 먼저 확인하세요.",
        "parameters": {
            "type": "object",
            "properties": {
                "data_map": {
                    "type": "object",
                    "description": "{'필드이름': '값'} 형태의 딕셔너리",
                }
            },
            "required": ["data_map"],
        },
    },
    {
        "name": "replace_text_patterns",
        "description": "{{이름}}, [날짜] 등 텍스트 패턴을 실제 데이터로 일괄 치환합니다. 일반 텍스트 수정의 기본 도구입니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "mapping": {
                    "type": "object",
                    "description": "{'찾을텍스트': '바꿀텍스트'} 딕셔너리",
                }
            },
            "required": ["mapping"],
        },
    },
    {
        "name": "set_checkbox_state",
        "description": "'label [ ]' 또는 '[ ] label' 형태의 체크박스 상태를 변경합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "label": {"type": "string", "description": "체크박스 옆 텍스트"},
                "is_checked": {"type": "boolean", "description": "체크 여부"},
            },
            "required": ["label", "is_checked"],
        },
    },
    {
        "name": "insert_image_box",
        "description": "특정 텍스트 위치에 이미지를 삽입합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "anchor_text": {"type": "string", "description": "이미지를 삽입할 위치의 텍스트"},
                "image_path": {"type": "string", "description": "이미지 파일 절대 경로"},
                "size_mode": {
                    "type": "string",
                    "enum": ["fit", "original"],
                    "description": "fit=셀에 맞춤, original=원본 크기",
                },
            },
            "required": ["image_path"],
        },
    },
    {
        "name": "sync_table_rows",
        "description": "입력할 데이터 개수에 맞춰 표의 데이터 행을 추가/삭제합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "table_index": {"type": "integer"},
                "data_count": {
                    "type": "integer",
                    "description": "채울 데이터의 행 수 (헤더 제외)",
                },
                "header_rows": {
                    "type": "integer",
                    "description": "헤더 행 수 (기본값: 1)",
                    "default": 1,
                },
            },
            "required": ["table_index", "data_count"],
        },
    },
    {
        "name": "fill_table_data_matrix",
        "description": "2차원 배열 데이터를 표의 특정 행부터 채웁니다. 기존 셀 텍스트를 교체합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "table_index": {"type": "integer"},
                "start_row": {
                    "type": "integer",
                    "description": "시작 행 인덱스 (0=첫 행, 헤더 다음 행이면 1)",
                },
                "matrix": {
                    "type": "array",
                    "items": {"type": "array", "items": {"type": "string"}},
                    "description": "2차원 문자열 배열 [[행1열1, 행1열2], [행2열1, ...]]",
                },
                "cell_delay": {
                    "type": "number",
                    "description": "셀 교체 간 딜레이(초). 기본값 0.3. 0이면 딜레이 없음.",
                },
            },
            "required": ["table_index", "start_row", "matrix"],
        },
    },
    {
        "name": "format_table_cells",
        "description": "현재 선택된 셀의 배경색이나 테두리를 변경합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "table_index": {"type": "integer"},
                "format_dict": {
                    "type": "object",
                    "description": "{'fill_color': RGB정수, 'border_width': 너비} 형태",
                },
            },
            "required": ["table_index", "format_dict"],
        },
    },
    {
        "name": "set_font_style",
        "description": "현재 선택 영역의 글자 서식을 설정합니다. (글꼴, 크기, 굵기, 색상)",
        "parameters": {
            "type": "object",
            "properties": {
                "font_name": {"type": "string", "description": "글꼴 이름 (예: '맑은 고딕')"},
                "size_pt": {"type": "number", "description": "글자 크기 (포인트)"},
                "bold": {"type": "boolean"},
                "color_rgb": {"type": "integer", "description": "글자색 RGB 정수 (예: 0xFF0000=빨강)"},
            },
            "required": [],
        },
    },
    {
        "name": "auto_fit_paragraph",
        "description": "자간을 좁혀 텍스트를 한 줄에 맞춥니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "decrease_count": {
                    "type": "integer",
                    "description": "자간 축소 단계 수 (기본값: 3)",
                    "default": 3,
                }
            },
            "required": [],
        },
    },
    {
        "name": "append_page_from_template",
        "description": "다른 HWP 파일을 현재 문서 끝에 이어 붙입니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "병합할 HWP 파일 경로"}
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "manage_page_visibility",
        "description": "특정 페이지를 감추거나 표시합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "page_number": {"type": "integer"},
                "action": {
                    "type": "string",
                    "enum": ["hide", "show"],
                    "description": "hide=감추기, show=표시",
                },
            },
            "required": ["page_number", "action"],
        },
    },
    {
        "name": "export_to_pdf",
        "description": "현재 문서를 PDF로 저장합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "target_path": {"type": "string", "description": "저장할 PDF 파일 경로"}
            },
            "required": ["target_path"],
        },
    },
    {
        "name": "execute_raw_action",
        "description": (
            "ActionTable.md의 모든 HWP COM Action을 직접 실행합니다. "
            "다른 도구로 처리 불가능한 특수 케이스에 사용하세요. "
            "ParameterSet '-' 표시 Action은 params 없이 호출하면 hwp.Run()으로 실행됩니다. "
            "'*' 표시 Action은 반드시 params를 포함해야 합니다."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "action_id": {
                    "type": "string",
                    "description": "ActionTable.md의 Action ID (예: 'TableInsertLowerRow')",
                },
                "params": {
                    "type": "object",
                    "description": "ParameterSet Item 딕셔너리 (없으면 생략)",
                },
            },
            "required": ["action_id"],
        },
    },
]


class LlmClient:
    """다중 공급자 LLM 클라이언트 — Tool Use(Function Calling) 지원"""

    def __init__(self, provider: Provider, api_key: str, model: str = None):
        self.provider = provider
        cleaned = "".join(c for c in api_key.strip() if ord(c) < 128)
        if not cleaned:
            raise ValueError("API Key가 비어 있거나 유효하지 않습니다.")
        self.api_key = cleaned
        self.model = model or DEFAULT_MODELS.get(provider, "")
        if self.provider not in DEFAULT_MODELS:
            raise ValueError(f"지원하지 않는 공급자: {self.provider}")

    # ──────────────────────────────────────────────
    # 에이전틱 Tool Use 루프 (메인 인터페이스)
    # ──────────────────────────────────────────────

    def call_agentic(
        self,
        doc_summary: str,
        user_query: str,
        tool_executor: Callable[[str, dict], str],
        log_cb: Callable[[str, str], None] = None,
        max_turns: int = 12,
    ) -> str:
        """
        에이전틱 Tool Use 루프.
        LLM이 HWP 도구를 직접 호출하며 문서를 탐색·수정합니다.

        Args:
            doc_summary: 문서 초기 요약 (analyze_document_structure 결과)
            user_query: 사용자 수정 요청
            tool_executor: (tool_name, tool_args) → result_str 콜백 (메인 스레드에서 실행)
            log_cb: (message, tag) 로그 콜백
            max_turns: 최대 LLM 호출 횟수
        Returns:
            LLM의 최종 응답 텍스트
        """
        def _log(msg: str, tag: str = "dim"):
            if log_cb:
                log_cb(msg, tag)

        initial_message = (
            f"[문서 초기 정보]\n{doc_summary}\n\n"
            f"[사용자 요청]\n{user_query}"
        )

        if self.provider == "openai":
            return self._agentic_openai(initial_message, tool_executor, _log, max_turns)
        elif self.provider == "anthropic":
            return self._agentic_anthropic(initial_message, tool_executor, _log, max_turns)
        elif self.provider == "gemini":
            return self._agentic_gemini(initial_message, tool_executor, _log, max_turns)
        else:
            raise ValueError(f"알 수 없는 공급자: {self.provider}")

    # ──────────────────────────────────────────────
    # OpenAI Function Calling
    # ──────────────────────────────────────────────

    def _agentic_openai(self, initial_message: str, tool_executor, log_cb, max_turns: int) -> str:
        from openai import OpenAI
        client = OpenAI(api_key=self.api_key)

        openai_tools = [
            {"type": "function", "function": {
                "name": t["name"],
                "description": t["description"],
                "parameters": t["parameters"],
            }}
            for t in HWP_TOOLS
        ]

        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": initial_message},
        ]

        for turn in range(max_turns):
            response = client.chat.completions.create(
                model=self.model,
                messages=messages,
                tools=openai_tools,
                tool_choice="auto",
                temperature=0.1,
            )
            choice = response.choices[0]
            messages.append(choice.message.model_dump())

            if choice.finish_reason == "stop":
                return choice.message.content or ""

            if choice.finish_reason == "tool_calls":
                tool_results = []
                for tc in (choice.message.tool_calls or []):
                    args = json.loads(tc.function.arguments or "{}")
                    log_cb(f"  🔧 {tc.function.name}({_fmt_args(args)})", "info")
                    result = tool_executor(tc.function.name, args)
                    log_cb(f"     → {str(result)[:120]}", "dim")
                    tool_results.append({
                        "role": "tool",
                        "tool_call_id": tc.id,
                        "content": str(result),
                    })
                messages.extend(tool_results)

        return "⚠️ 최대 턴 수 초과"

    # ──────────────────────────────────────────────
    # Anthropic Tool Use
    # ──────────────────────────────────────────────

    def _agentic_anthropic(self, initial_message: str, tool_executor, log_cb, max_turns: int) -> str:
        import anthropic
        client = anthropic.Anthropic(api_key=self.api_key)

        anthropic_tools = [
            {
                "name": t["name"],
                "description": t["description"],
                "input_schema": t["parameters"],
            }
            for t in HWP_TOOLS
        ]

        messages = [{"role": "user", "content": initial_message}]

        for turn in range(max_turns):
            response = client.messages.create(
                model=self.model,
                max_tokens=4096,
                system=SYSTEM_PROMPT,
                tools=anthropic_tools,
                messages=messages,
            )

            # assistant 메시지 추가
            messages.append({"role": "assistant", "content": response.content})

            if response.stop_reason == "end_turn":
                # 최종 텍스트 추출
                for block in response.content:
                    if hasattr(block, "text"):
                        return block.text
                return ""

            if response.stop_reason == "tool_use":
                tool_results = []
                for block in response.content:
                    if block.type == "tool_use":
                        log_cb(f"  🔧 {block.name}({_fmt_args(block.input)})", "info")
                        result = tool_executor(block.name, block.input)
                        log_cb(f"     → {str(result)[:120]}", "dim")
                        tool_results.append({
                            "type": "tool_result",
                            "tool_use_id": block.id,
                            "content": str(result),
                        })
                messages.append({"role": "user", "content": tool_results})

        return "⚠️ 최대 턴 수 초과"

    # ──────────────────────────────────────────────
    # Google Gemini Function Calling
    # ──────────────────────────────────────────────

    def _agentic_gemini(self, initial_message: str, tool_executor, log_cb, max_turns: int) -> str:
        import google.generativeai as genai
        from google.generativeai import types as gtypes

        genai.configure(api_key=self.api_key)

        func_declarations = []
        for t in HWP_TOOLS:
            func_declarations.append(
                gtypes.FunctionDeclaration(
                    name=t["name"],
                    description=t["description"],
                    parameters=t["parameters"] if t["parameters"].get("properties") else None,
                )
            )
        gemini_tools = [gtypes.Tool(function_declarations=func_declarations)]

        model = genai.GenerativeModel(
            model_name=self.model,
            system_instruction=SYSTEM_PROMPT,
            generation_config=genai.GenerationConfig(temperature=0.1),
            tools=gemini_tools,
        )

        history = []
        chat = model.start_chat(history=history)

        response = chat.send_message(initial_message)

        for turn in range(max_turns):
            # 함수 호출 확인
            fn_calls = [
                part.function_call
                for part in response.parts
                if hasattr(part, "function_call") and part.function_call.name
            ]

            if not fn_calls:
                # 최종 텍스트 응답
                return response.text

            # 도구 실행 후 결과 반환
            fn_responses = []
            for fc in fn_calls:
                args = dict(fc.args)
                log_cb(f"  🔧 {fc.name}({_fmt_args(args)})", "info")
                result = tool_executor(fc.name, args)
                log_cb(f"     → {str(result)[:120]}", "dim")
                fn_responses.append(
                    gtypes.Part.from_function_response(
                        name=fc.name,
                        response={"result": str(result)},
                    )
                )

            response = chat.send_message(fn_responses)

        return "⚠️ 최대 턴 수 초과"

    # ──────────────────────────────────────────────
    # 레거시 호환 (단방향 JSON 방식)
    # ──────────────────────────────────────────────

    def build_prompt(self, doc_summary: str, user_query: str) -> str:
        return (
            f"{doc_summary}\n\n"
            f"## 사용자 요청\n{user_query}\n\n"
            f'위 요청에 따라 문서를 수정하는 명령을 {{"commands": [...]}} 형태의 JSON으로 반환하세요.'
        )

    def call(self, doc_summary: str, user_query: str) -> tuple[list[dict], str]:
        """레거시 단방향 JSON 방식. 신규 코드는 call_agentic()을 사용하세요."""
        prompt = self.build_prompt(doc_summary, user_query)

        legacy_system = """당신은 HWP 문서 편집 전문가입니다.
반드시 {"commands": [...]} 형태의 JSON만 반환하세요.
지원 action: find_and_replace, replace_paragraph, set_field_value"""

        if self.provider == "openai":
            raw = self._call_openai_legacy(prompt, legacy_system)
        elif self.provider == "gemini":
            raw = self._call_gemini_legacy(prompt, legacy_system)
        elif self.provider == "anthropic":
            raw = self._call_anthropic_legacy(prompt, legacy_system)
        else:
            raise ValueError(f"알 수 없는 공급자: {self.provider}")

        commands = self._parse_response(raw)
        return commands, raw

    def _call_openai_legacy(self, prompt: str, system: str) -> str:
        from openai import OpenAI
        client = OpenAI(api_key=self.api_key)
        response = client.chat.completions.create(
            model=self.model,
            messages=[{"role": "system", "content": system},
                      {"role": "user", "content": prompt}],
            temperature=0.1,
            response_format={"type": "json_object"},
        )
        return response.choices[0].message.content

    def _call_gemini_legacy(self, prompt: str, system: str) -> str:
        import google.generativeai as genai
        genai.configure(api_key=self.api_key)
        model = genai.GenerativeModel(
            model_name=self.model,
            system_instruction=system,
            generation_config=genai.GenerationConfig(
                temperature=0.1,
                response_mime_type="application/json",
            ),
        )
        return model.generate_content(prompt).text

    def _call_anthropic_legacy(self, prompt: str, system: str) -> str:
        import anthropic
        client = anthropic.Anthropic(api_key=self.api_key)
        message = client.messages.create(
            model=self.model,
            max_tokens=4096,
            system=system,
            messages=[{"role": "user", "content": prompt}],
        )
        return message.content[0].text

    def _parse_response(self, raw: str) -> list[dict]:
        text = re.sub(r"```(?:json)?\s*", "", raw.strip()).strip().rstrip("`").strip()
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            match = re.search(r"\[.*\]", text, re.DOTALL)
            if match:
                parsed = json.loads(match.group())
            else:
                raise ValueError(f"JSON 파싱 실패:\n{text[:500]}")

        if isinstance(parsed, dict):
            for key in ("commands", "edits", "actions", "result", "modifications", "changes"):
                if key in parsed and isinstance(parsed[key], list):
                    parsed = parsed[key]
                    break
            else:
                if "action" in parsed:
                    parsed = [parsed]
                else:
                    for v in parsed.values():
                        if isinstance(v, list):
                            parsed = v
                            break
                    else:
                        parsed = []

        if not isinstance(parsed, list):
            raise ValueError(f"예상치 못한 응답 형식: {type(parsed)}")

        return [item for item in parsed if isinstance(item, dict) and "action" in item]

    def test_connection(self) -> bool:
        try:
            self.call(
                doc_summary="[테스트]\n내용: 테스트",
                user_query="아무것도 변경하지 마세요."
            )
            return True
        except Exception:
            return False


def _fmt_args(args: dict) -> str:
    """로그용 args 간략 표현"""
    if not args:
        return ""
    parts = []
    for k, v in args.items():
        sv = str(v)
        if len(sv) > 40:
            sv = sv[:40] + "..."
        parts.append(f"{k}={sv}")
    return ", ".join(parts[:3])
