# HWP 에이전트 Function Chain 설계

> 작성일: 2026-04-06
> 목적: 사용자 명령 유형별 최적 도구 호출 순서(chain) 정의 및 도구 공백(gap) 식별

---

## 1. 현재 도구 목록

| 도구 | 분류 | 비고 |
|------|------|------|
| `analyze_document_structure` | 읽기 | 페이지 수, 표 개수, 필드 목록 |
| `get_document_text` | 읽기 | 전체 텍스트 (8000자 제한) |
| `get_field_info` | 읽기 | 누름틀 목록 |
| `get_all_tables_overview` | 읽기 | 모든 표 개요 (헤더만) |
| `get_table_schema` | 읽기 | 특정 표 셀 내용 (60자 truncate) |
| `get_cell_text` | 읽기 | 특정 셀 전체 텍스트 (truncate 없음) |
| `find_text_anchor` | 읽기 | 텍스트 존재 여부 (HTML-only) |
| `fill_field_data` | 쓰기 | 누름틀 채우기 |
| `replace_text_patterns` | 쓰기 | 텍스트 패턴 치환 (단락 내) |
| `set_checkbox_state` | 쓰기 | 체크박스 on/off |
| `fill_table_data_matrix` | 쓰기 | 표 셀 일괄 입력 |
| `sync_table_rows` | 쓰기 | 표 행 수 조정 |
| `format_table_cells` | 서식 | 셀 배경/테두리 |
| `set_font_style` | 서식 | 폰트/크기/볼드/색상 |
| `auto_fit_paragraph` | 서식 | 자간 축소 (한 줄 맞춤) |
| `insert_image_box` | 삽입 | 이미지 삽입 |
| `append_page_from_template` | 삽입 | HWP 파일 이어붙이기 |
| `manage_page_visibility` | 구조 | 페이지 숨기기/표시 |
| `export_to_pdf` | 출력 | PDF 저장 |
| `execute_raw_action` | 탈출구 | COM Action 직접 실행 |

---

## 2. 명령 유형별 Function Chain

### 2-1. 읽기/조회 (Read)

#### 패턴 A: "~~ 표에 뭐가 있어?" / "~~ 표 구조 알려줘"

```
get_all_tables_overview()
  → (index 확인)
  → get_table_schema(index)
```

**제약**: `get_table_schema` 셀 내용 60자 truncate.
**판단 기준**: 60자 이내로 답할 수 있으면 충분. 아니면 Pattern B로 진입.

---

#### 패턴 B: "~~ 셀 내용 요약/번역/분석해줘"

```
get_all_tables_overview()
  → get_table_schema(index)          ← row/col 번호 확인
  → get_cell_text(index, row, col)   ← 전체 텍스트 읽기
  → [LLM 추론: 요약/번역/분석]
```

**주의**: `get_document_text`는 8000자 제한이 있어 긴 문서에서 특정 셀만 필요할 때 낭비적.
→ **항상 `get_cell_text`를 우선 사용**.

---

#### 패턴 C: "몇 페이지야?" / "필드 목록 알려줘" / "전체 구조 파악"

```
analyze_document_structure()
```

또는 누름틀이 필요한 경우:
```
get_field_info()
```

---

#### 패턴 D: "~~ 단어/텍스트 있어?"

```
find_text_anchor(keyword)
```

**제약**: 단락 경계(줄바꿈)를 포함하는 키워드는 검색 불가 → `⛔ found: false` 반환.

---

### 2-2. 수정/교체 (Edit)

#### 패턴 E: "~~ 표의 ~~ 셀을 ~~로 바꿔줘" (직접 입력)

```
get_all_tables_overview()
  → get_table_schema(index)
  → fill_table_data_matrix(index, start_row, [[값]], start_col)
```

**결정 규칙**:
- `start_col=0`: 라벨 열 포함해서 덮어쓸 때
- `start_col=1`: 라벨 열(col 0) 보존, 내용 열(col 1+)만 수정

---

#### 패턴 F: "~~ 셀 내용을 LLM이 재작성해줘" (읽기 → 재작성 → 쓰기)

```
get_all_tables_overview()
  → get_table_schema(index)
  → get_cell_text(index, row, col)    ← 원본 전체 읽기
  → [LLM 추론: 재작성]
  → fill_table_data_matrix(index, row, [[재작성된 내용]])
```

---

#### 패턴 G: "~~를 ~~로 바꿔줘" (표 밖 텍스트)

```
find_text_anchor(keyword)             ← 존재 여부 확인 (⛔ 방지)
  → [found=true] replace_text_patterns({keyword: 새값})
  → [found=false] 사용자에게 보고
```

**제약**:
- `keyword`에 `\n` 포함 불가 (단락 경계 검색 불가)
- `keyword` 80자 초과 불가

---

#### 패턴 H: "누름틀 ~~에 ~~를 입력해줘"

```
get_field_info()
  → fill_field_data({필드명: 값})
```

---

#### 패턴 I: "체크박스 ~~ 체크해줘 / 해제해줘"

```
set_checkbox_state(label, is_checked)
```

---

#### 패턴 J: "데이터 개수에 맞게 표 행 늘리고 채워줘"

```
get_all_tables_overview()
  → sync_table_rows(index, data_count, header_rows)
  → fill_table_data_matrix(index, start_row, matrix)
```

---

### 2-3. 삽입/추가 (Insert)

#### 패턴 K: "~~위치에 이미지 삽입해줘"

```
find_text_anchor(anchor_text)         ← 삽입 위치 확인
  → insert_image_box(anchor_text, image_path, size_mode)
```

---

#### 패턴 L: "~~파일을 뒤에 이어붙여줘"

```
append_page_from_template(file_path)
```

---

#### 패턴 M: "표 추가" / "섹션 추가" ← **⚠️ 현재 도구 없음**

```
execute_raw_action("TableCreate", {rows, cols, ...})   ← COM 직접
```

→ **도구 공백**: `create_table` 전용 도구 필요 (아래 섹션 참조)

---

### 2-4. 삭제 (Delete)

#### 패턴 N: "~~ 텍스트 지워줘"

```
find_text_anchor(keyword)
  → replace_text_patterns({keyword: ""})
```

---

#### 패턴 O: "표에서 ~~ 행 지워줘 / 행 수 줄여줘"

```
get_all_tables_overview()
  → sync_table_rows(index, new_count, header_rows)
```

---

#### 패턴 P: "빈 단락 / 특정 패턴 모두 제거" ← **⚠️ 현재 도구 미흡**

```
execute_raw_action("AllReplace", {find: "  +", replace: " ", use_regex: true})
```

→ **도구 공백**: 정규식 기반 replace 도구 필요.

---

### 2-5. 구조 변환 (Transform)

#### 패턴 Q: "이 표를 bullet list로" / "내용을 표로 만들어줘" ← **⚠️ 현재 도구 없음**

현재 직접 지원 불가. 복합 체인 필요:

```
get_table_schema(index)               ← 원본 읽기
  → [LLM 추론: 변환 데이터 생성]
  → execute_raw_action("TableCreate", ...)   또는
  → replace_text_patterns(...)              (텍스트로 변환 시)
```

→ **도구 공백**: `create_table`, `delete_table`, `insert_paragraph` 필요.

---

### 2-6. 서식 조작 (Style)

#### 패턴 R: "~~텍스트 폰트 키워줘 / 볼드 적용"

```
find_text_anchor(keyword)             ← 위치 확인
  → set_font_style(font_name, size_pt, bold, color_rgb)
```

**한계**: `set_font_style`은 현재 커서 선택 영역에만 작용 → 커서 이동 선행 필요.
→ **보완 필요**: `select_text(keyword)` 도구로 선택 영역 지정 후 서식 적용.

---

#### 패턴 S: "~~ 셀 배경색 바꿔줘"

```
get_all_tables_overview()
  → format_table_cells(index, {row, col, background_color})
```

---

#### 패턴 T: "텍스트가 한 줄에 안 들어가, 자간 줄여줘"

```
auto_fit_paragraph(decrease_count)
```

---

### 2-7. 출력 (Export)

#### 패턴 U: "PDF로 저장해줘"

```
export_to_pdf(target_path)
```

---

## 3. 도구 공백(Gap) 분석

| 사용자 명령 예시 | 현재 대응 | 필요한 신규 도구 | 우선순위 |
|----------------|----------|----------------|---------|
| "표 하나 추가해줘" | `execute_raw_action` 우회 | `create_table(rows, cols, position)` | 중 |
| "이 표 삭제해줘" | `execute_raw_action` 우회 | `delete_table(index)` | 중 |
| "정규식으로 전체 치환" | 불가 | `replace_regex_patterns(pattern, replacement)` | 낮 |
| "~~ 단어 선택 후 서식 적용" | 커서 위치 불확실 | `select_text(keyword)` | 높 |
| "특정 셀 텍스트 지우기만" | `fill_table_data_matrix([""])` | (현재 가능, 명시화 필요) | 낮 |
| "단락 삽입 (표 밖)" | `execute_raw_action` 우회 | `insert_paragraph(anchor, text)` | 중 |
| "특정 셀 몇 개 읽기 (배치)" | `get_cell_text` 반복 | `get_cells_batch([{index,row,col}])` | 중 |

---

## 4. 시스템 프롬프트 개선 방향

### 4-1. 현재 문제점

- LLM이 `get_all_tables_overview` 없이 바로 `get_table_schema(0)` 호출하는 경우 있음
- `fill_table_data_matrix`에서 `matrix=[["줄1","줄2"]]` (2셀) vs `[["줄1\n줄2"]]` (1셀 2단락) 혼동 여전히 가능
- 셀 내용 읽기가 필요한 경우 `get_document_text` 호출 후 "너무 길어서 못 하겠다"고 포기

### 4-2. 시스템 프롬프트 개선안

```
## 명령 유형별 필수 체인

### 표 셀 읽기/요약/번역
1. get_all_tables_overview()     → 표 번호 확인
2. get_table_schema(index)       → row/col 번호 확인
3. get_cell_text(index, row, col) → 전체 텍스트 (잘리지 않음)
⚠️ get_document_text는 표 셀 읽기에 쓰지 마세요.

### 표 셀 수정
1. get_all_tables_overview()
2. get_table_schema(index)       → 현재 내용 확인 후
3. fill_table_data_matrix(index, start_row, [["값"]], start_col=1)
⚠️ matrix=[[줄1, 줄2]] ← 이것은 2개 셀. 한 셀 2단락은 [["줄1\n줄2"]]

### 표 밖 텍스트 치환
1. find_text_anchor(keyword)     → found 확인 후에만 치환
2. replace_text_patterns(...)
⚠️ \n 포함 keyword 사용 금지
```

---

## 5. 도구 추가 구현 우선순위 로드맵

| 순위 | 도구 | 이유 |
|------|------|------|
| 1 | `select_text(keyword)` | `set_font_style` 연계 필수 |
| 2 | `get_cells_batch([{index,row,col}])` | 여러 셀 한 번에 읽기 (토큰 절약) |
| 3 | `create_table(rows, cols, anchor_text)` | 삽입 명령 지원 |
| 4 | `delete_table(index)` | 삭제 명령 지원 |
| 5 | `insert_paragraph(anchor_text, text, position)` | 단락 삽입 지원 |

---

## 6. 체인 결정 트리 (LLM 판단 보조)

```
사용자 명령 수신
├── "읽기/조회"?
│   ├── 특정 셀 내용 필요 → get_all_tables_overview → get_table_schema → get_cell_text
│   ├── 표 구조만 필요  → get_all_tables_overview (→ get_table_schema 선택)
│   ├── 문서 전체 통계  → analyze_document_structure
│   └── 텍스트 존재 확인 → find_text_anchor
│
├── "수정/교체"?
│   ├── 표 안 셀       → get_all_tables_overview → get_table_schema → fill_table_data_matrix
│   ├── 누름틀         → get_field_info → fill_field_data
│   ├── 표 밖 텍스트   → find_text_anchor → replace_text_patterns
│   └── 체크박스       → set_checkbox_state
│
├── "삽입"?
│   ├── 이미지         → find_text_anchor → insert_image_box
│   └── 페이지/파일    → append_page_from_template
│
├── "삭제"?
│   ├── 텍스트 삭제    → find_text_anchor → replace_text_patterns(→"")
│   └── 행 삭제        → sync_table_rows
│
└── "서식"?
    ├── 셀 서식        → get_all_tables_overview → format_table_cells
    ├── 폰트 서식      → find_text_anchor → set_font_style   ⚠️ select_text 미구현
    └── 자간 조절      → auto_fit_paragraph
```
