# HWP-LLM Wrapper

한글(HWP/HWPX) 문서를 LLM으로 편집하는 Python PoC 도구입니다.  
COM 자동화를 통해 한글 프로그램을 제어하며, BYOK(Bring Your Own Key) 방식으로 OpenAI, Google Gemini, Anthropic Claude를 지원합니다.

## 요구사항

- **Windows** + **한글(한/글) 2020 이상** 설치 필요
- Python 3.10+

## 설치

```bash
pip install -r requirements.txt
```

## 실행

```bash
python main.py
```

## 사용 방법

1. **파일 선택** — HWP 파일 경로를 선택합니다.
2. **한글에서 열기** — 한글 프로그램을 실행하고 파일을 엽니다.
3. **LLM 설정** — 공급자(OpenAI/Gemini/Anthropic), 모델, API Key를 입력합니다.
4. **수정 요청 입력** — 자연어로 원하는 수정 내용을 입력합니다.
   - 예: `"제목을 '2026년 보고서'로 바꿔줘"`
   - 예: `"본문에서 '홍길동'을 '김철수'로 모두 바꿔줘"`
5. **AI로 수정 실행** — LLM이 명령을 생성하고 자동으로 문서를 수정합니다.
6. `_edited` 접미사가 붙은 사본으로 저장됩니다 (원본 유지).

## 프로젝트 구조

```
hwp-llm-wrapper/
├── main.py              # Tkinter GUI 진입점
├── hwp_controller.py    # HWP COM 자동화 레이어
├── llm_client.py        # LLM API 클라이언트 (BYOK)
└── requirements.txt     # 의존성
```

## 지원 수정 명령

LLM이 JSON 형태로 명령을 생성합니다:

| action | 설명 |
|---|---|
| `find_and_replace` | 특정 텍스트를 찾아 교체 |
| `replace_paragraph` | 특정 단락 전체 교체 |
| `set_field_value` | 누름틀(필드) 값 설정 |
| `insert_text` | 특정 단락 뒤에 텍스트 삽입 |

## 향후 계획

- [ ] Tauri 기반 크로스플랫폼 GUI 전환
- [ ] 문서 diff 미리보기 (수정 전/후 비교)
- [ ] 표/이미지 편집 지원
- [ ] OS 키체인 연동 (API Key 보안 강화)
