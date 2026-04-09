/// hwp_controller.rs
/// HWP COM 자동화 컨트롤러
///
/// 설계: COM STA 스레드에서만 실행. spawn_hwp_thread()로 전용 스레드 생성 후
/// mpsc 채널로 명령을 전달하면 결과를 oneshot 채널로 받습니다.
///
/// 표 읽기는 InitScan/GetText + 테이블 네비게이션 기반.
/// HTML 파싱에 의존하지 않습니다.

use crate::com_dispatch::{ComObject, Variant};
use serde_json::{json, Value};
use std::time::Duration;

// ──────────────────────────────────────────────────────────────
// 명령 열거형 (채널 통신용)
// ──────────────────────────────────────────────────────────────

#[derive(Debug)]
pub enum HwpCommand {
    Connect { visible: bool },
    OpenFile { path: String },
    Save { save_path: Option<String> },
    Close,
    DispatchTool { name: String, args: Value },
    IsConnected,
    Snapshot { backup_path: String },
    Rollback { backup_path: String },
}

pub struct HwpRequest {
    pub command: HwpCommand,
    pub reply: tokio::sync::oneshot::Sender<anyhow::Result<String>>,
}

// ──────────────────────────────────────────────────────────────
// HWP 스레드 엔트리포인트
// ──────────────────────────────────────────────────────────────

pub fn spawn_hwp_thread(
    mut rx: tokio::sync::mpsc::Receiver<HwpRequest>,
) -> std::thread::JoinHandle<()> {
    std::thread::spawn(move || {
        #[cfg(windows)]
        {
            if let Err(e) = crate::com_dispatch::com_initialize() {
                eprintln!("COM 초기화 실패: {e}");
                return;
            }
        }

        let mut ctrl = HwpController::new();

        while let Some(req) = rx.blocking_recv() {
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                ctrl.handle_command(req.command)
            }));
            let reply_result = match result {
                Ok(r) => r,
                Err(_) => Err(anyhow::anyhow!("COM 스레드 패닉 발생 — 복구됨")),
            };
            let _ = req.reply.send(reply_result);
        }

        #[cfg(windows)]
        crate::com_dispatch::com_uninitialize();
    })
}

// ──────────────────────────────────────────────────────────────
// HwpController
// ──────────────────────────────────────────────────────────────

pub struct HwpController {
    hwp: Option<ComObject>,
    connected: bool,
}

impl HwpController {
    pub fn new() -> Self {
        HwpController {
            hwp: None,
            connected: false,
        }
    }

    fn handle_command(&mut self, cmd: HwpCommand) -> anyhow::Result<String> {
        match cmd {
            HwpCommand::Connect { visible } => {
                self.connect(visible)?;
                Ok("connected".to_string())
            }
            HwpCommand::OpenFile { path } => {
                self.open_file(&path)?;
                Ok("opened".to_string())
            }
            HwpCommand::Save { save_path } => {
                let p = self.save(save_path.as_deref())?;
                Ok(p)
            }
            HwpCommand::Close => {
                self.close()?;
                Ok("closed".to_string())
            }
            HwpCommand::IsConnected => Ok(self.connected.to_string()),
            HwpCommand::DispatchTool { name, args } => self.dispatch_tool(&name, &args),
            HwpCommand::Snapshot { backup_path } => {
                self.save(Some(&backup_path))?;
                Ok(backup_path)
            }
            HwpCommand::Rollback { backup_path } => {
                self.open_file(&backup_path)?;
                Ok("롤백 완료".to_string())
            }
        }
    }

    fn hwp(&self) -> anyhow::Result<&ComObject> {
        self.hwp
            .as_ref()
            .ok_or_else(|| anyhow::anyhow!("HWP에 연결되지 않았습니다. 먼저 connect()를 호출하세요."))
    }

    // ─────────────── 연결 / 파일 관리 ───────────────────────

    fn connect(&mut self, _visible: bool) -> anyhow::Result<()> {
        let hwp = ComObject::from_prog_id("HWPFrame.HwpObject")?;
        eprintln!("[HWP] CoCreateInstance 성공");

        match hwp.get("XHwpWindows") {
            Ok(windows) => {
                if let Some(wins) = windows.as_object() {
                    match wins.call("Item", vec![Variant::I32(0)]) {
                        Ok(item) => {
                            if let Some(win) = item.as_object() {
                                match win.put("Visible", Variant::Bool(true)) {
                                    Ok(_) => eprintln!("[HWP] Visible=true 설정 성공"),
                                    Err(e) => eprintln!("[HWP] Visible 설정 실패: {e}"),
                                }
                            } else {
                                eprintln!("[HWP] Item(0) 이 Object가 아님");
                            }
                        }
                        Err(e) => eprintln!("[HWP] XHwpWindows.Item(0) 실패: {e}"),
                    }
                } else {
                    eprintln!("[HWP] XHwpWindows가 Object가 아님");
                }
            }
            Err(e) => eprintln!("[HWP] XHwpWindows 접근 실패: {e}"),
        }

        let _ = hwp.call(
            "RegisterModule",
            vec![
                Variant::String("FilePathCheckDLL".to_string()),
                Variant::String("FilePathCheckerModule".to_string()),
            ],
        );

        self.hwp = Some(hwp);
        self.connected = true;
        Ok(())
    }

    fn open_file(&mut self, path: &str) -> anyhow::Result<()> {
        let hwp = self.hwp()?;
        let abs = std::path::Path::new(path)
            .canonicalize()
            .unwrap_or_else(|_| std::path::PathBuf::from(path));
        let abs_str = abs.to_string_lossy().to_string();

        if !abs.exists() {
            anyhow::bail!("파일 없음: {abs_str}");
        }
        let ext = abs
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();
        if ext != "hwp" && ext != "hwpx" {
            anyhow::bail!("HWP/HWPX 파일만 지원합니다.");
        }
        let fmt = if ext == "hwpx" { "HWPX" } else { "HWP" };
        let result = hwp.call(
            "Open",
            vec![
                Variant::String(abs_str.clone()),
                Variant::String(fmt.to_string()),
                Variant::String("forceopen:true".to_string()),
            ],
        )?;
        if result.as_bool() == Some(false) {
            let _ = hwp.call(
                "Open",
                vec![
                    Variant::String(abs_str),
                    Variant::String(String::new()),
                    Variant::String("forceopen:true".to_string()),
                ],
            );
        }
        Ok(())
    }

    fn save(&self, save_path: Option<&str>) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        if let Some(path) = save_path {
            // SaveAs COM 직접 호출 — FileSaveAs_S Action과 달리 UI 다이얼로그가 절대 뜨지 않음
            hwp.call("SaveAs", vec![
                Variant::String(path.to_string()),
                Variant::String("HWP".to_string()),
                Variant::String(String::new()),
            ])?;
            Ok(path.to_string())
        } else {
            let act = hwp
                .call("CreateAction", vec![Variant::String("FileSave_S".to_string())])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CreateAction 실패"))?;
            let ps = act.call("CreateSet", vec![])?;
            act.call("Execute", vec![ps])?;
            let path_var = hwp.get("Path")?;
            Ok(path_var.to_string_repr())
        }
    }

    fn close(&mut self) -> anyhow::Result<()> {
        if let Some(hwp) = &self.hwp {
            let _ = hwp.call("Quit", vec![]);
        }
        self.hwp = None;
        self.connected = false;
        Ok(())
    }

    // ─────────────── HTML 기반 문서 읽기 ────────────────────
    // Python PoC와 동일하게 GetTextFile("HTML") → HTML 파싱 방식을 사용합니다.
    // GetPos/GetText/KeyIndicator의 VT_BYREF output 파라미터는 신뢰할 수 없어서
    // 모든 READ 경로는 이 HTML 기반 방식으로 대체합니다.

    fn get_html(&self) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let result = hwp.call(
            "GetTextFile",
            vec![
                Variant::String("HTML".to_string()),
                Variant::String(String::new()),
            ],
        )?;
        Ok(result.to_string_repr())
    }

    /// HTML에서 누름틀(field) 목록을 파싱합니다.
    ///
    /// HWP HTML 출력 포맷:
    ///   <a name="FieldStart:"></a>     ← 이름 없는 누름틀 시작
    ///   <a name="FieldStart:이름"></a>  ← 이름 있는 누름틀 시작
    ///   <span ...>내용</span>
    ///   <a name="FieldEnd:"></a>       ← 누름틀 끝
    ///
    /// placeholder 판별: span에 font-style:italic + color:#ff0000 → 미입력 상태
    fn parse_fields_from_html(html: &str) -> Vec<HtmlField> {
        let mut fields = Vec::new();
        let mut search = html;
        let mut idx = 0usize;

        while let Some(start_pos) = search.find("<a name=\"FieldStart:") {
            let after_tag = &search[start_pos..];

            // FieldStart 태그에서 field name 추출
            let name_start = after_tag.find('"').map(|p| p + 1).unwrap_or(0);
            let colon_offset = after_tag[name_start..].find(':').unwrap_or(0);
            let name_end = after_tag[name_start..].find('"').unwrap_or(0);
            let field_name = if colon_offset + 1 < name_end {
                after_tag[name_start + colon_offset + 1..name_start + name_end].to_string()
            } else {
                String::new()
            };

            // FieldEnd 위치 찾기
            let end_marker = "FieldEnd:";
            let Some(end_pos) = after_tag.find(end_marker) else { break; };

            // FieldStart ~ FieldEnd 사이의 HTML 청크
            let chunk = &after_tag[..end_pos];
            let content = extract_plain_text_from_html_chunk(chunk);
            let is_placeholder = chunk.contains("font-style:italic")
                && chunk.to_lowercase().contains("color:#ff0000");

            // FieldEnd 이후 텍스트: 다음 FieldStart 또는 단락 끝까지
            // find_replace 시 고유 문맥(suffix)을 구성하는 데 사용
            let after_end_html = {
                let consumed_so_far = end_pos + end_marker.len();
                // FieldEnd: 뒤에 오는 닫는 '>' 건너뜀
                let rest = &after_tag[consumed_so_far..];
                let past_gt = rest.find('>').map(|p| &rest[p + 1..]).unwrap_or(rest);
                // 다음 FieldStart 또는 </p> 까지의 HTML 청크
                let next_field = past_gt.find("<a name=\"FieldStart:").unwrap_or(past_gt.len());
                let next_para  = past_gt.find("</p>").unwrap_or(past_gt.len());
                &past_gt[..next_field.min(next_para)]
            };
            // HTML 태그를 제거하되 <a> 같은 인라인 태그는 공백 삽입 없이 처리
            let text_after = extract_plain_text_from_html_chunk(after_end_html);
            let text_after = decode_html_entities(text_after.trim());

            fields.push(HtmlField {
                index: idx,
                name: field_name,
                content: content.trim().to_string(),
                text_after,
                is_placeholder,
            });
            idx += 1;

            // 다음 FieldStart 검색 위치: consumed = 현재 FieldStart 이후의 FieldEnd 위치 + end_marker 길이
            let consumed = start_pos + end_pos + end_marker.len();
            search = &search[consumed..];
        }

        fields
    }

    /// 문서에서 텍스트가 존재하는지 HTML 문자열에서 확인합니다.
    /// Python PoC의 _text_exists()와 동일 — ForwardFind 호출 전 사전 검증용.
    /// ForwardFind는 텍스트가 없으면 HWP 다이얼로그를 블로킹하므로 이 검사가 필수입니다.
    fn text_exists(&self, text: &str) -> bool {
        let html = match self.get_html() {
            Ok(h) => h,
            Err(_) => return false,
        };
        html_contains_text(&html, text, false)
    }

    /// 단락 수를 HTML에서 카운트합니다 (InitScan/GetText VT_BYREF 대체)
    fn count_paragraphs_from_html(&self) -> usize {
        let html = self.get_html().unwrap_or_default();
        let lower = html.to_lowercase();
        lower.matches("<p").count()
    }

    fn get_field_names(&self) -> Vec<String> {
        let hwp = match self.hwp() {
            Ok(h) => h,
            Err(_) => return vec![],
        };
        let result = hwp
            .call(
                "GetFieldList",
                vec![Variant::I32(0), Variant::I32(1)],
            )
            .ok()
            .map(|v| v.to_string_repr())
            .unwrap_or_default();
        result
            .split('\x02')
            .filter(|s| !s.is_empty())
            .map(str::to_string)
            .collect()
    }

    fn find_replace(&self, find: &str, replace: &str, case_sensitive: bool) -> anyhow::Result<()> {
        let hwp = self.hwp()?;
        // \n 포함 패턴은 공백으로 정규화 (HWP FindReplace는 개행 리터럴 미지원)
        let find_norm: String = find.split_whitespace().collect::<Vec<_>>().join(" ");
        let find = find_norm.as_str();
        if !self.text_exists(find) {
            return Ok(());
        }

        // 커서를 문서 처음으로 이동 — AllReplace가 문서 끝에서 시작할 경우
        // "문서 끝까지 찾을까요?" 다이얼로그가 떠서 COM이 블로킹되는 것을 방지
        hwp.call("Run", vec![Variant::String("MoveDocBegin".to_string())])?;

        let act = hwp
            .call("CreateAction", vec![Variant::String("AllReplace".to_string())])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("AllReplace 액션 생성 실패"))?;
        let ps = act
            .call("CreateSet", vec![])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
        ps.call("SetItem", vec![
            Variant::String("FindString".to_string()),
            Variant::String(find.to_string()),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("ReplaceString".to_string()),
            Variant::String(replace.to_string()),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("IgnoreCase".to_string()),
            Variant::Bool(!case_sensitive),
        ])?;
        // AllWordReplace=false: 부분 문자열도 치환
        ps.call("SetItem", vec![
            Variant::String("AllWordReplace".to_string()),
            Variant::Bool(false),
        ])?;
        // Direction=0: 커서(문서 처음) 기준 앞방향 전체 검색 — wrap 없으므로 다이얼로그 없음
        ps.call("SetItem", vec![
            Variant::String("Direction".to_string()),
            Variant::I32(0),
        ])?;
        // FindRegExp=false: 정규식 아닌 리터럴 문자열 매칭
        ps.call("SetItem", vec![
            Variant::String("FindRegExp".to_string()),
            Variant::Bool(false),
        ])?;
        // SearchTbl=true: 표 안 셀에서도 검색
        ps.call("SetItem", vec![
            Variant::String("SearchTbl".to_string()),
            Variant::Bool(true),
        ])?;
        act.call("Execute", vec![Variant::Object(ps)])?;
        Ok(())
    }

    // ─────────────── 검사 도구 ───────────────────────────────

    fn analyze_document_structure(&self) -> anyhow::Result<Value> {
        let hwp = self.hwp()?;
        let mut result = json!({
            "pages": 0,
            "table_count": 0,
            "fields": [],
            "paragraph_count": 0
        });

        if let Ok(pages) = hwp.get("PageCount") {
            result["pages"] = json!(pages.as_i32().unwrap_or(0));
        }

        // HTML 기반으로 표/단락 수 계산 (Python PoC와 동일)
        let html = self.get_html().unwrap_or_default();
        let lower = html.to_lowercase();
        result["table_count"] = json!(lower.matches("<table").count());
        result["paragraph_count"] = json!(lower.matches("<p").count());

        let fields = self.get_field_names();
        result["fields"] = json!(fields);

        Ok(result)
    }

    fn get_document_text(&self) -> anyhow::Result<Value> {
        // GetTextFile("TEXT") 직접 사용 — HTML 파싱 불필요, 14배 경량
        let hwp = self.hwp()?;
        let result = hwp.call(
            "GetTextFile",
            vec![
                Variant::String("TEXT".to_string()),
                Variant::String(String::new()),
            ],
        )?;
        let raw = result.to_string_repr();
        // \r\n → \n 정규화, 연속 빈 줄 제거
        let cleaned: String = raw
            .lines()
            .map(|l| l.trim())
            .filter(|l| !l.is_empty())
            .collect::<Vec<_>>()
            .join("\n");
        // 토큰 절감: 최대 8000자
        let truncated = truncate_chars(&cleaned, 8000);
        let note = if cleaned.chars().count() > 8000 {
            format!("\n…(이하 {}자 생략. 특정 표 셀 내용은 get_cell_text 사용 권장)", cleaned.chars().count() - 8000)
        } else {
            String::new()
        };
        Ok(json!({ "text": format!("{truncated}{note}") }))
    }

    fn get_field_info(&self) -> anyhow::Result<Value> {
        let hwp = self.hwp()?;
        let names = self.get_field_names();
        let mut fields = vec![];
        for name in names {
            let value = hwp
                .call("GetFieldText", vec![Variant::String(name.clone())])
                .ok()
                .map(|v| v.to_string_repr())
                .unwrap_or_default();
            fields.push(json!({"name": name, "value": value}));
        }
        Ok(json!(fields))
    }

    // ─────────────── 표 스캔 (HTML 기반) ────────────────────
    // Python PoC의 _parse_all_tables()와 동일한 방식입니다.
    // GetPos/KeyIndicator VT_BYREF 방식 대신 GetTextFile("HTML") 파싱을 사용합니다.

    fn scan_tables(&self) -> anyhow::Result<Vec<TableData>> {
        let html = self.get_html()?;
        Ok(parse_tables_from_html(&html))
    }


    fn get_all_tables_overview(&self) -> anyhow::Result<Value> {
        let tables = self.scan_tables()?;
        if tables.is_empty() {
            return Ok(json!([]));
        }

        let overview: Vec<Value> = tables
            .iter()
            .enumerate()
            .map(|(idx, table)| {
                let headers: Vec<String> = table
                    .cells
                    .first()
                    .cloned()
                    .unwrap_or_default()
                    .iter()
                    .map(|h| truncate_chars(h, 30))
                    .collect();

                // 제목표 판별:
                // (1) 행/열이 작고 (2) 셀 내용이 짧으며 (3) 다음 표가 더 크면 제목표로 간주
                let next_is_bigger = idx + 1 < tables.len()
                    && (tables[idx + 1].rows > table.rows || tables[idx + 1].cols > table.cols);
                let is_title = table.rows <= 2
                    && table.cols <= 3
                    && table.cells.iter().all(|r| r.iter().all(|c| c.len() < 100))
                    && next_is_bigger;

                let role = if is_title { "title_table" } else { "data_table" };

                let mut entry = json!({
                    "table_index": idx,
                    "role": role,
                    "rows": table.rows,
                    "cols": table.cols,
                    "headers": headers,
                });

                // 제목표인 경우 연결된 데이터표 정보 추가
                if is_title && idx + 1 < tables.len() {
                    let next = &tables[idx + 1];
                    let next_headers: Vec<String> = next
                        .cells
                        .first()
                        .cloned()
                        .unwrap_or_default()
                        .iter()
                        .take(4)
                        .map(|h| truncate_chars(h, 20))
                        .collect();
                    entry["linked_data_table"] = json!({
                        "table_index": idx + 1,
                        "rows": next.rows,
                        "cols": next.cols,
                        "headers": next_headers,
                    });
                }

                entry
            })
            .collect();
        Ok(json!(overview))
    }

    fn get_table_schema(&self, table_index: usize) -> anyhow::Result<Value> {
        let tables = self.scan_tables()?;
        if table_index >= tables.len() {
            return Ok(json!({
                "rows": 0, "cols": 0, "headers": [], "cells": [],
                "error": format!("표 {}번 없음 (총 {}개)", table_index, tables.len())
            }));
        }
        let table = &tables[table_index];
        let headers: Vec<String> = table
            .cells
            .first()
            .cloned()
            .unwrap_or_default();

        // cells를 [{row: N, cells: [...]}] 형식으로 반환 — LLM이 start_row를 정확히 파악
        // 셀 내용은 60자로 잘라 토큰 절감
        let cells: Vec<Value> = table
            .cells
            .iter()
            .enumerate()
            .map(|(i, row)| {
                let trimmed: Vec<String> = row.iter().map(|c| truncate_chars(c, 60)).collect();
                json!({"row": i, "cells": trimmed})
            })
            .collect();

        Ok(json!({
            "table_index": table_index,
            "rows": table.rows,
            "cols": table.cols,
            "headers": headers,
            "cells": cells,
        }))
    }

    fn get_cell_text(&self, table_index: usize, row: usize, col: usize) -> anyhow::Result<Value> {
        let tables = self.scan_tables()?;
        if table_index >= tables.len() {
            anyhow::bail!("표 {table_index}번 없음 (총 {}개)", tables.len());
        }
        let table = &tables[table_index];
        let cell_text = table
            .cells
            .get(row)
            .and_then(|r| r.get(col))
            .map(|s| s.as_str())
            .unwrap_or("");

        // 원본 그대로 반환 (truncate 없음) — 요약/번역 등 전체 내용이 필요한 경우 사용
        Ok(json!({
            "table_index": table_index,
            "row": row,
            "col": col,
            "text": cell_text,
        }))
    }

    fn find_text_anchor(&self, keyword: &str) -> anyhow::Result<Value> {
        // ForwardFind를 사용하지 않고 HTML 기반으로만 존재 여부 확인.
        // ForwardFind는 단락 경계를 넘는 텍스트를 못 찾으면 "문서 마지막까지 찾으시겠습니까?"
        // 다이얼로그를 띄워 COM을 블로킹하므로 이 함수에서는 사용하지 않음.
        let found = self.text_exists(keyword);
        Ok(json!({"found": found, "keyword": keyword}))
    }

    // ─────────────── 데이터 입력 ────────────────────────────

    fn fill_field_data(&self, data_map: &Value) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let map = data_map.as_object().ok_or_else(|| anyhow::anyhow!("data_map이 object여야 합니다"))?;
        let mut results = vec![];
        for (field_name, value) in map {
            let val_str = value.as_str().map(str::to_string).unwrap_or_else(|| value.to_string());
            match hwp.call("SetFieldText", vec![
                Variant::String(field_name.clone()),
                Variant::String(val_str.clone()),
            ]) {
                Ok(_) => results.push(format!("✅ '{field_name}' = '{val_str}'")),
                Err(e) => results.push(format!("❌ '{field_name}': {e}")),
            }
        }
        Ok(results.join("\n"))
    }

    fn replace_text_patterns(&self, mapping: &Value) -> anyhow::Result<String> {
        let mut parsed_fallback: Value;
        let map = if let Some(m) = mapping.as_object() {
            m
        } else if let Some(s) = mapping.as_str() {
            parsed_fallback = serde_json::from_str(s)
                .map_err(|_| anyhow::anyhow!("mapping 문자열 JSON 파싱 실패: {s}"))?;
            parsed_fallback.as_object()
                .ok_or_else(|| anyhow::anyhow!("mapping이 object여야 합니다 (파싱 결과: {parsed_fallback})"))?
        } else {
            anyhow::bail!("mapping이 object여야 합니다 (받은 값: {mapping})");
        };
        let html = self.get_html().unwrap_or_default();
        let mut results = vec![];
        for (pattern, value) in map {
            let val_str = value.as_str().map(str::to_string).unwrap_or_else(|| value.to_string());

            // \n 포함 패턴: HWP AllReplace는 단락 경계를 넘어 검색 불가
            // 표 셀의 여러 단락에 걸친 내용은 fill_table_data_matrix를 사용해야 함
            if pattern.contains('\n') {
                results.push(format!(
                    "⛔ '{}' — 검색어에 줄바꿈(\\n)이 포함되어 있습니다. \
                    HWP는 단락 경계를 넘어 검색할 수 없습니다. \
                    표 셀 전체 내용을 바꾸려면 fill_table_data_matrix를 사용하세요. \
                    (get_all_tables_overview → 표·행 인덱스 확인 → fill_table_data_matrix)",
                    truncate_chars(pattern, 40)
                ));
                continue;
            }

            // 긴 패턴 경고: 표 셀 내용 전체를 검색어로 쓰면 HWP FindReplace가 단락 구분자 때문에 실패함
            let pattern_norm_chars: usize = pattern.split_whitespace().collect::<Vec<_>>().join(" ").chars().count();
            if pattern_norm_chars > 80 {
                results.push(format!(
                    "⛔ '{}'... — 검색어가 너무 깁니다({pattern_norm_chars}자). \
                    표 셀 내용 전체를 바꾸려면 fill_table_data_matrix를 사용하세요. \
                    (get_all_tables_overview → 표 인덱스 확인 → fill_table_data_matrix)",
                    truncate_chars(pattern, 30)
                ));
                continue;
            }

            // text_exists 사전 확인 후 skip 여부를 명시적으로 보고
            if !html_contains_text(&html, pattern, true) {
                results.push(format!("⚠️ '{}' 문서에서 찾을 수 없어 건너뜀", truncate_chars(pattern, 50)));
                continue;
            }
            match self.find_replace(pattern, &val_str, true) {
                Ok(_) => results.push(format!("✅ '{}' → '{}'", truncate_chars(pattern, 40), truncate_chars(&val_str, 40))),
                Err(e) => results.push(format!("❌ '{}': {e}", truncate_chars(pattern, 40))),
            }
        }
        Ok(results.join("\n"))
    }

    fn set_checkbox_state(&self, label: &str, is_checked: bool) -> anyhow::Result<String> {
        let check_char = if is_checked { "V" } else { " " };
        let undo_char = if is_checked { " " } else { "V" };
        for sep in &["", " "] {
            for box_str in &["[ ]", &format!("[{undo_char}]")] {
                let old = format!("{label}{sep}{box_str}");
                let new = format!("{label}{sep}[{check_char}]");
                let _ = self.find_replace(&old, &new, true);
                let old = format!("{box_str}{sep}{label}");
                let new = format!("[{check_char}]{sep}{label}");
                let _ = self.find_replace(&old, &new, true);
            }
        }
        let state = if is_checked { "체크" } else { "해제" };
        Ok(format!("체크박스 '{label}' {state} 처리 완료"))
    }

    fn insert_image_box(&self, anchor_text: &str, image_path: &str, size_mode: &str) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let abs = std::path::Path::new(image_path)
            .canonicalize()
            .map_err(|_| anyhow::anyhow!("이미지 파일 없음: {image_path}"))?;
        let abs_str = abs.to_string_lossy().to_string();

        if !anchor_text.is_empty() {
            let act = hwp
                .call("CreateAction", vec![Variant::String("ForwardFind".to_string())])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("ForwardFind 생성 실패"))?;
            let ps = act
                .call("CreateSet", vec![])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
            ps.call("SetItem", vec![
                Variant::String("FindString".to_string()),
                Variant::String(anchor_text.to_string()),
            ])?;
            ps.call("SetItem", vec![
                Variant::String("IgnoreCase".to_string()),
                Variant::Bool(true),
            ])?;
            act.call("Execute", vec![Variant::Object(ps)])?;
        }
        let sizetype = if size_mode == "fit" { 1i32 } else { 0i32 };
        hwp.call(
            "InsertPicture",
            vec![
                Variant::String(abs_str.clone()),
                Variant::Bool(true),
                Variant::I32(sizetype),
                Variant::Bool(false),
                Variant::Bool(false),
            ],
        )?;
        let fname = abs.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or(abs_str);
        Ok(format!("✅ 이미지 삽입: {fname}"))
    }

    // ─────────────── 표 동적 제어 ───────────────────────────

    fn sync_table_rows(&self, table_index: usize, data_count: usize, header_rows: usize) -> anyhow::Result<String> {
        let schema = self.get_table_schema(table_index)?;
        let current_rows = schema["rows"].as_u64().unwrap_or(0) as usize;
        if current_rows == 0 {
            return Ok(format!("❌ 표 {table_index}번을 찾을 수 없습니다."));
        }
        let current_data_rows = current_rows.saturating_sub(header_rows);
        if current_data_rows == data_count {
            return Ok(format!("변경 없음 (현재 {current_data_rows}행)"));
        }

        let hwp = self.hwp()?;
        // 표의 마지막 행으로 이동: 해당 표로 네비게이트
        self.navigate_to_table(table_index)?;
        // 표 끝으로 이동
        hwp.call("Run", vec![Variant::String("MoveTopLevelEnd".to_string())])?;

        if data_count > current_data_rows {
            let diff = data_count - current_data_rows;
            for _ in 0..diff {
                hwp.call("Run", vec![Variant::String("TableInsertLowerRow".to_string())])?;
            }
            Ok(format!("✅ 행 {diff}개 추가 (총 {data_count}행)"))
        } else {
            let diff = current_data_rows - data_count;
            for _ in 0..diff {
                if let Ok(act) = hwp.call("CreateAction", vec![Variant::String("TableDeleteRow".to_string())]) {
                    if let Some(act_obj) = act.as_object() {
                        if let Ok(ps) = act_obj.call("CreateSet", vec![]) {
                            let _ = act_obj.call("Execute", vec![ps]);
                        }
                    }
                } else {
                    let _ = hwp.call("Run", vec![Variant::String("TableDeleteRow".to_string())]);
                }
            }
            Ok(format!("✅ 행 {diff}개 삭제 (총 {data_count}행)"))
        }
    }

    /// 특정 인덱스의 표로 커서를 이동합니다.
    /// **반드시 col 0**의 셀 텍스트를 anchor로 사용하여 ForwardFind합니다.
    /// 반환값: anchor가 위치한 행 인덱스 (fill_table_data_matrix에서 정확한 TableDownCell 수 계산에 사용)
    fn navigate_to_table(&self, table_index: usize) -> anyhow::Result<usize> {
        let hwp = self.hwp()?;

        // HTML에서 표 목록을 가져와 anchor 텍스트 결정
        let html = self.get_html()?;
        let tables = parse_tables_from_html(&html);

        if table_index >= tables.len() {
            anyhow::bail!("표 {table_index}번을 찾을 수 없습니다 (총 {}개)", tables.len());
        }

        let table = &tables[table_index];

        // **col 0**에서만 anchor 선택 — TableDownCell이 col 0에서 출발해야 start_row 계산이 정확함
        // col 0이 빈 행은 건너뛰고, 2자 이상인 첫 번째 행의 col 0 셀을 선택
        let (anchor_row, anchor) = table
            .cells
            .iter()
            .enumerate()
            .find_map(|(row_idx, row)| {
                let cell = row.get(0)?;
                let text = cell.trim();
                if text.len() >= 2 { Some((row_idx, text.to_string())) } else { None }
            })
            .ok_or_else(|| anyhow::anyhow!("표 {table_index}번 col 0에서 anchor를 찾을 수 없습니다"))?;

        // HTML에서 텍스트 존재 여부 먼저 확인 (ForwardFind 블로킹 방지)
        if !html_contains_text(&html, &anchor, false) {
            anyhow::bail!("anchor 텍스트를 문서에서 찾을 수 없습니다: {anchor:?}");
        }

        hwp.call("Run", vec![Variant::String("MoveDocBegin".to_string())])?;

        let act = hwp
            .call("CreateAction", vec![Variant::String("ForwardFind".to_string())])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("ForwardFind 생성 실패"))?;
        let ps = act
            .call("CreateSet", vec![])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
        ps.call("SetItem", vec![
            Variant::String("FindString".to_string()),
            Variant::String(anchor.clone()),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("IgnoreCase".to_string()),
            Variant::Bool(false),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("Direction".to_string()),
            Variant::I32(0),  // 0 = 문서 시작점에서 앞방향 (MoveDocBegin 이후)
        ])?;
        ps.call("SetItem", vec![
            Variant::String("FindReplace".to_string()),
            Variant::I32(0),
        ])?;

        let found = act
            .call("Execute", vec![Variant::Object(ps)])?
            .as_bool()
            .unwrap_or(false);

        if !found {
            anyhow::bail!("ForwardFind로 anchor를 찾지 못했습니다: {anchor:?}");
        }

        // 커서가 anchor 텍스트 안에 있음 → anchor_row 반환
        Ok(anchor_row)
    }

    /// (start_row, start_col)에 커서를 정확히 위치시킵니다.
    ///
    /// ## 전략: 물리 셀 오프셋 계산 + TableRightCell 걷기
    ///
    /// 1. HTML에서 표의 물리 셀 구조 (colspan/rowspan) 파싱
    /// 2. navigate_to_table → anchor 행(col 0)에 커서 위치
    /// 3. anchor 물리 offset과 target 물리 offset의 delta만큼 TableRightCell/TableLeftCell
    ///
    /// ForwardFind 기반 이전 방식 폐기: rowspan 표에서 오작동 (텍스트 중복, 빈 셀 등)
    fn navigate_to_row_in_table(
        &self,
        table_index: usize,
        start_row: usize,
        _schema: &Value,   // 하위 호환 — 더 이상 사용 안 함
        start_col: usize,
    ) -> anyhow::Result<()> {
        let html = self.get_html()?;
        let raw_tables = parse_physical_tables(&html);
        let raw = raw_tables
            .get(table_index)
            .ok_or_else(|| anyhow::anyhow!("표 {table_index}번 없음 (HTML에서 {}개 발견)", raw_tables.len()))?;

        let target_offset = physical_cell_offset(raw, start_row, start_col)?;

        // navigate_to_table은 anchor 행의 col 0에 커서를 위치시키고 anchor_row를 반환
        let anchor_row = self.navigate_to_table(table_index)?;
        let anchor_offset = physical_cell_offset(raw, anchor_row, 0)?;

        let hwp = self.hwp()?;
        let delta = target_offset as isize - anchor_offset as isize;

        eprintln!("[NAV] table={table_index} anchor_row={anchor_row}(offset={anchor_offset}) → target({start_row},{start_col})(offset={target_offset}) delta={delta}");

        if delta > 0 {
            for _ in 0..delta {
                hwp.call("Run", vec![Variant::String("TableRightCell".to_string())])?;
            }
        } else if delta < 0 {
            for _ in 0..(-delta) {
                hwp.call("Run", vec![Variant::String("TableLeftCell".to_string())])?;
            }
        }

        Ok(())
    }

    fn fill_table_data_matrix(
        &self,
        table_index: usize,
        start_row: usize,
        start_col: usize,
        matrix: &Value,
        cell_delay_secs: f64,
    ) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let schema = self.get_table_schema(table_index)?;
        let total_rows = schema["rows"].as_u64().unwrap_or(0) as usize;
        if total_rows == 0 {
            return Ok(format!("❌ 표 {table_index}번을 찾을 수 없습니다."));
        }

        let data_rows = matrix
            .as_array()
            .ok_or_else(|| anyhow::anyhow!("matrix가 배열이어야 합니다"))?;

        // 물리 셀 구조: colspan 연속 셀 skip 판단에 사용
        let html = self.get_html()?;
        let raw_tables = parse_physical_tables(&html);

        let mut filled = 0usize;

        for (r_offset, row_data) in data_rows.iter().enumerate() {
            let row_arr = match row_data.as_array() {
                Some(a) => a,
                None => continue,
            };
            let current_row = start_row + r_offset;

            // 각 행마다 독립적으로 navigate_to_row_in_table 호출.
            // TableDownCell + TableLeftCell × N 방식은 colspan 표에서 오작동하므로 폐기.
            self.navigate_to_row_in_table(table_index, current_row, &schema, start_col)?;

            for (c_idx, cell_val) in row_arr.iter().enumerate() {
                let vcol = start_col + c_idx;
                let cell_type = vcol_cell_type(&raw_tables, table_index, current_row, vcol);

                // 다음 HWP 스텝이 필요한 셀이 남아 있는지 확인 (ColspanCont는 HWP 스텝 아님)
                let has_more_hwp = ((c_idx + 1)..row_arr.len()).any(|ci| {
                    vcol_cell_type(&raw_tables, table_index, current_row, start_col + ci)
                        != VcolType::ColspanCont
                });

                match cell_type {
                    VcolType::ColspanCont => {
                        // 같은 물리 셀의 시각적 확장 — HWP 스텝 없으므로 쓰기/이동 모두 불필요
                        continue;
                    }
                    VcolType::RowspanCont => {
                        // 이전 행 rowspan의 재방문 셀 — HWP는 커서를 세우지만 쓰기는 건너뜀
                        // 다음 셀로 가기 위해 TableRightCell만 수행
                        if has_more_hwp {
                            hwp.call("Run", vec![Variant::String("TableRightCell".to_string())])?;
                        }
                        continue;
                    }
                    VcolType::Physical => {
                        // 실제 쓰기 대상 셀
                        let new_text = cell_val
                            .as_str()
                            .map(str::to_string)
                            .unwrap_or_else(|| cell_val.to_string());

                        hwp.call("Run", vec![Variant::String("SelectAll".to_string())])?;

                        // \n 포함 시 단락 나누기를 삽입하며 여러 줄로 입력
                        let lines: Vec<&str> = new_text.split('\n').collect();
                        for (li, line) in lines.iter().enumerate() {
                            let line_clean = line.trim_end_matches('\r').to_string();
                            let ins = hwp
                                .call("CreateAction", vec![Variant::String("InsertText".to_string())])?
                                .as_object()
                                .ok_or_else(|| anyhow::anyhow!("InsertText 생성 실패"))?;
                            let ips = ins
                                .call("CreateSet", vec![])?
                                .as_object()
                                .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
                            ips.call("SetItem", vec![
                                Variant::String("Text".to_string()),
                                Variant::String(line_clean),
                            ])?;
                            ins.call("Execute", vec![Variant::Object(ips)])?;
                            if li < lines.len() - 1 {
                                hwp.call("Run", vec![Variant::String("InsertReturn".to_string())])?;
                            }
                        }
                        filled += 1;

                        if cell_delay_secs > 0.0 {
                            std::thread::sleep(Duration::from_secs_f64(cell_delay_secs));
                        }

                        if has_more_hwp {
                            hwp.call("Run", vec![Variant::String("TableRightCell".to_string())])?;
                        }
                    }
                }
            }
        }

        Ok(format!("✅ {filled}개 셀 채움"))
    }

    fn format_table_cells(&self, _table_index: usize, format_dict: &Value) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        if let Some(fill_color) = format_dict.get("fill_color").and_then(|v| v.as_i64()) {
            let act = hwp
                .call("CreateAction", vec![Variant::String("CellFill".to_string())])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CellFill 생성 실패"))?;
            let ps = act
                .call("CreateSet", vec![])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
            ps.call("SetItem", vec![
                Variant::String("FillColor".to_string()),
                Variant::I32(fill_color as i32),
            ])?;
            act.call("Execute", vec![Variant::Object(ps)])?;
        }
        if let Some(border_width) = format_dict.get("border_width").and_then(|v| v.as_i64()) {
            let act = hwp
                .call("CreateAction", vec![Variant::String("CellBorderFill".to_string())])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CellBorderFill 생성 실패"))?;
            let ps = act
                .call("CreateSet", vec![])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
            ps.call("SetItem", vec![
                Variant::String("BorderWidth".to_string()),
                Variant::I32(border_width as i32),
            ])?;
            act.call("Execute", vec![Variant::Object(ps)])?;
        }
        Ok("✅ 셀 서식 적용".to_string())
    }

    // ─────────────── 서식 ───────────────────────────────────

    fn set_font_style(
        &self,
        font_name: Option<&str>,
        size_pt: Option<f64>,
        bold: Option<bool>,
        color_rgb: Option<i64>,
    ) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let act = hwp
            .call("CreateAction", vec![Variant::String("CharShape".to_string())])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("CharShape 생성 실패"))?;
        let ps = act
            .call("CreateSet", vec![])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
        if let Some(pt) = size_pt {
            ps.call("SetItem", vec![
                Variant::String("Height".to_string()),
                Variant::I32((pt * 100.0) as i32),
            ])?;
        }
        if let Some(name) = font_name {
            for key in &["FaceName", "FaceNameHangul", "FaceNameLatin", "FaceNameHanja"] {
                ps.call("SetItem", vec![
                    Variant::String(key.to_string()),
                    Variant::String(name.to_string()),
                ])?;
            }
        }
        if let Some(b) = bold {
            ps.call("SetItem", vec![
                Variant::String("Bold".to_string()),
                Variant::Bool(b),
            ])?;
        }
        if let Some(color) = color_rgb {
            ps.call("SetItem", vec![
                Variant::String("TextColor".to_string()),
                Variant::I32(color as i32),
            ])?;
        }
        act.call("Execute", vec![Variant::Object(ps)])?;
        Ok("✅ 글자 서식 적용".to_string())
    }

    fn auto_fit_paragraph(&self, decrease_count: usize) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        for _ in 0..decrease_count {
            hwp.call("Run", vec![Variant::String("CharShapeSpacingDecrease".to_string())])?;
        }
        Ok(format!("✅ 자간 {decrease_count}단계 축소"))
    }

    // ─────────────── 구조 편집 ──────────────────────────────

    fn append_page_from_template(&self, file_path: &str) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let abs = std::path::Path::new(file_path)
            .canonicalize()
            .map_err(|_| anyhow::anyhow!("파일 없음: {file_path}"))?;
        let abs_str = abs.to_string_lossy().to_string();
        hwp.call("Run", vec![Variant::String("MoveDocEnd".to_string())])?;
        let act = hwp
            .call("CreateAction", vec![Variant::String("InsertFile".to_string())])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("InsertFile 생성 실패"))?;
        let ps = act
            .call("CreateSet", vec![])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
        ps.call("SetItem", vec![
            Variant::String("filename".to_string()),
            Variant::String(abs_str),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("KeepSection".to_string()),
            Variant::Bool(false),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("KeepCharshape".to_string()),
            Variant::Bool(true),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("KeepParashape".to_string()),
            Variant::Bool(true),
        ])?;
        act.call("Execute", vec![Variant::Object(ps)])?;
        let fname = abs
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_default();
        Ok(format!("✅ 파일 병합: {fname}"))
    }

    fn manage_page_visibility(&self, page_number: i32, action: &str) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let act = hwp
            .call("CreateAction", vec![Variant::String("PageHiding".to_string())])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("PageHiding 생성 실패"))?;
        let ps = act
            .call("CreateSet", vec![])?
            .as_object()
            .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
        ps.call("SetItem", vec![
            Variant::String("PageNum".to_string()),
            Variant::I32(page_number),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("Hide".to_string()),
            Variant::Bool(action == "hide"),
        ])?;
        act.call("Execute", vec![Variant::Object(ps)])?;
        let state = if action == "hide" { "감춤" } else { "표시" };
        Ok(format!("✅ 페이지 {page_number} {state}"))
    }

    fn export_to_pdf(&self, target_path: &str) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let abs = std::path::Path::new(target_path)
            .to_str()
            .unwrap_or(target_path)
            .to_string();
        match (|| -> anyhow::Result<()> {
            let act = hwp
                .call("CreateAction", vec![Variant::String("PrintToPDF".to_string())])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("PrintToPDF 생성 실패"))?;
            let ps = act
                .call("CreateSet", vec![])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
            ps.call("SetItem", vec![
                Variant::String("filename".to_string()),
                Variant::String(abs.clone()),
            ])?;
            ps.call("SetItem", vec![
                Variant::String("Range".to_string()),
                Variant::I32(3),
            ])?;
            ps.call("SetItem", vec![
                Variant::String("Copies".to_string()),
                Variant::I32(1),
            ])?;
            act.call("Execute", vec![Variant::Object(ps)])?;
            Ok(())
        })() {
            Ok(_) => Ok(format!("✅ PDF 저장: {abs}")),
            Err(_) => {
                hwp.call("SaveAs", vec![
                    Variant::String(abs.clone()),
                    Variant::String("PDF".to_string()),
                    Variant::String(String::new()),
                ])?;
                Ok(format!("✅ PDF 저장: {abs}"))
            }
        }
    }

    fn execute_raw_action(&self, action_id: &str, params: Option<&Value>) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        if let Some(p) = params {
            if let Some(map) = p.as_object() {
                let act = hwp
                    .call("CreateAction", vec![Variant::String(action_id.to_string())])?
                    .as_object()
                    .ok_or_else(|| anyhow::anyhow!("CreateAction 실패"))?;
                let ps = act
                    .call("CreateSet", vec![])?
                    .as_object()
                    .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
                for (k, v) in map {
                    let val = match v {
                        Value::Bool(b) => Variant::Bool(*b),
                        Value::Number(n) => {
                            if let Some(i) = n.as_i64() {
                                Variant::I32(i as i32)
                            } else {
                                Variant::F64(n.as_f64().unwrap_or(0.0))
                            }
                        }
                        Value::String(s) => Variant::String(s.clone()),
                        _ => Variant::String(v.to_string()),
                    };
                    ps.call("SetItem", vec![Variant::String(k.clone()), val])?;
                }
                let result = act
                    .call("Execute", vec![Variant::Object(ps)])?
                    .as_bool()
                    .unwrap_or(false);
                return Ok(format!(
                    "✅ {action_id}: {}",
                    if result { "성공" } else { "실패(False 반환)" }
                ));
            }
        }
        let result = hwp
            .call("Run", vec![Variant::String(action_id.to_string())])?
            .as_bool()
            .unwrap_or(false);
        Ok(format!(
            "✅ {action_id}: {}",
            if result { "성공" } else { "실패(False 반환)" }
        ))
    }

    // ─────────────── 문서 요약 텍스트 ───────────────────────

    #[allow(dead_code)]
    pub fn get_document_summary(&self) -> String {
        let mut lines = vec![];
        let hwp_opt = self.hwp();

        if let Ok(hwp) = &hwp_opt {
            if let Ok(pages) = hwp.get("PageCount") {
                lines.push(format!("페이지 수: {}", pages.to_string_repr()));
            }
        }

        let tables = self.scan_tables().unwrap_or_default();
        lines.push(format!("표 개수: {}", tables.len()));

        let fields = self.get_field_names();
        if !fields.is_empty() {
            lines.push(format!("누름틀 필드: {}", fields.join(", ")));
        }

        for (i, table) in tables.iter().enumerate().take(5) {
            if let Some(header) = table.cells.first() {
                let h: Vec<String> = header.iter().take(5).cloned().collect();
                lines.push(format!("표[{i}] 헤더: {}", h.join(" | ")));
            }
        }

        if let Ok(html) = self.get_html() {
            let plain = html_to_plain_text(&html);
            if !plain.is_empty() {
                let truncated = if plain.chars().count() > 2000 {
                    format!("{}...", plain.chars().take(2000).collect::<String>())
                } else {
                    plain
                };
                lines.push(String::new());
                lines.push("=== 본문 텍스트 ===".to_string());
                lines.push(truncated);
            }
        }

        lines.join("\n")
    }

    // ─────────────── Tool Dispatcher ────────────────────────

    pub fn dispatch_tool(&mut self, name: &str, args: &Value) -> anyhow::Result<String> {
        match name {
            "analyze_document_structure" => {
                Ok(self.analyze_document_structure()?.to_string())
            }
            "get_document_text" => Ok(self.get_document_text()?.to_string()),
            "get_field_info" => Ok(self.get_field_info()?.to_string()),
            "get_all_tables_overview" => Ok(self.get_all_tables_overview()?.to_string()),
            "get_table_schema" => {
                let idx = args["table_index"].as_u64().unwrap_or(0) as usize;
                Ok(self.get_table_schema(idx)?.to_string())
            }
            "get_cell_text" => {
                let idx = args["table_index"].as_u64().unwrap_or(0) as usize;
                let row = args["row"].as_u64().unwrap_or(0) as usize;
                let col = args["col"].as_u64().unwrap_or(0) as usize;
                Ok(self.get_cell_text(idx, row, col)?.to_string())
            }
            "find_text_anchor" => {
                let kw = args["keyword"].as_str().unwrap_or("");
                Ok(self.find_text_anchor(kw)?.to_string())
            }
            "fill_field_data" => {
                let data_map = if args.get("data_map").is_some() && !args["data_map"].is_null() {
                    &args["data_map"]
                } else {
                    args
                };
                self.fill_field_data(data_map)
            }
            "replace_text_patterns" => {
                // "mapping" / "replacements" / "patterns" 키 모두 수용
                let mapping = ["mapping", "replacements", "patterns"]
                    .iter()
                    .find_map(|k| {
                        let v = args.get(*k)?;
                        if v.is_null() { None } else { Some(v) }
                    })
                    .unwrap_or(args);
                self.replace_text_patterns(mapping)
            }
            "set_checkbox_state" => {
                let label = args["label"].as_str().unwrap_or("");
                let checked = args["is_checked"].as_bool().unwrap_or(true);
                self.set_checkbox_state(label, checked)
            }
            "insert_image_box" => {
                let anchor = args["anchor_text"].as_str().unwrap_or("");
                let path = args["image_path"].as_str().unwrap_or("");
                let size = args["size_mode"].as_str().unwrap_or("original");
                self.insert_image_box(anchor, path, size)
            }
            "sync_table_rows" => {
                let ti = args["table_index"].as_u64().unwrap_or(0) as usize;
                let dc = args["data_count"].as_u64().unwrap_or(0) as usize;
                let hr = args["header_rows"].as_u64().unwrap_or(1) as usize;
                self.sync_table_rows(ti, dc, hr)
            }
            "fill_table_data_matrix" => {
                let ti = args["table_index"].as_u64().unwrap_or(0) as usize;
                let sr = args["start_row"].as_u64().unwrap_or(1) as usize;
                let sc = args["start_col"].as_u64().unwrap_or(0) as usize;
                let mtx = &args["matrix"];
                let delay = args["cell_delay"].as_f64().unwrap_or(0.0);
                self.fill_table_data_matrix(ti, sr, sc, mtx, delay)
            }
            "format_table_cells" => {
                let ti = args["table_index"].as_u64().unwrap_or(0) as usize;
                let fmt = args.get("format_dict").unwrap_or(&Value::Null);
                self.format_table_cells(ti, fmt)
            }
            "set_font_style" => {
                self.set_font_style(
                    args["font_name"].as_str(),
                    args["size_pt"].as_f64(),
                    args["bold"].as_bool(),
                    args["color_rgb"].as_i64(),
                )
            }
            "auto_fit_paragraph" => {
                let cnt = args["decrease_count"].as_u64().unwrap_or(3) as usize;
                self.auto_fit_paragraph(cnt)
            }
            "append_page_from_template" => {
                let fp = args["file_path"].as_str().unwrap_or("");
                self.append_page_from_template(fp)
            }
            "manage_page_visibility" => {
                let pn = args["page_number"].as_i64().unwrap_or(1) as i32;
                let ac = args["action"].as_str().unwrap_or("hide");
                self.manage_page_visibility(pn, ac)
            }
            "export_to_pdf" => {
                let tp = args["target_path"].as_str().unwrap_or("");
                self.export_to_pdf(tp)
            }
            "execute_raw_action" => {
                let ai = args["action_id"].as_str().unwrap_or("");
                self.execute_raw_action(ai, args.get("params"))
            }

            // ── 진단 도구 (ToolTestPage 전용) ─────────────────────────
            "diag_raw_html" => {
                let html = self.get_html()?;
                let limit = args["limit"].as_u64().unwrap_or(3000) as usize;
                let lower = html.to_lowercase();
                let preview: String = html.chars().take(limit).collect();
                Ok(json!({
                    "total_bytes": html.len(),
                    "total_chars": html.chars().count(),
                    "table_tag_count": lower.matches("<table").count(),
                    "tr_tag_count": lower.matches("<tr").count(),
                    "td_tag_count": lower.matches("<td").count(),
                    "p_tag_count": lower.matches("<p").count(),
                    "br_tag_count": lower.matches("<br").count(),
                    "preview": preview,
                }).to_string())
            }
            "diag_text_file_txt" => {
                // GetTextFile("TEXT") — HTML 파싱 없는 대안. 탭(\t)으로 셀 구분.
                let hwp = self.hwp()?;
                let result = hwp.call(
                    "GetTextFile",
                    vec![
                        Variant::String("TEXT".to_string()),
                        Variant::String(String::new()),
                    ],
                )?;
                let text = result.to_string_repr();
                let limit = args["limit"].as_u64().unwrap_or(3000) as usize;
                let preview: String = text.chars().take(limit).collect();
                // 이스케이프 시각화 (\t \r \n 명시)
                let escaped: String = text.chars().take(limit).map(|c| match c {
                    '\n' => "↵\n".to_string(),
                    '\r' => "⏎".to_string(),
                    '\t' => "→\t".to_string(),
                    c    => c.to_string(),
                }).collect();
                Ok(json!({
                    "total_chars": text.chars().count(),
                    "preview_plain": preview,
                    "preview_escaped": escaped,
                }).to_string())
            }
            "diag_normalize_keyword" => {
                // html_contains_text 내부 정규화를 시뮬레이션하여
                // "false positive" 가능성을 진단합니다.
                let keyword = args["keyword"].as_str().unwrap_or("");
                let html = self.get_html()?;
                let plain = html_to_plain_text(&html);
                let normalize = |s: &str| -> String {
                    s.split_whitespace().collect::<Vec<_>>().join(" ")
                };
                let kw_norm   = normalize(keyword);
                let plain_norm = normalize(&plain);
                let found_norm = plain_norm.to_lowercase().contains(&kw_norm.to_lowercase());
                let found_raw  = plain.contains(keyword);
                Ok(json!({
                    "keyword_raw":         keyword,
                    "keyword_normalized":  kw_norm,
                    "found_by_html_check": found_norm,
                    "found_raw_in_plain":  found_raw,
                    // true = html_contains_text가 "있다"고 판단하지만 실제로는 단락 경계에 걸려있을 가능성
                    "false_positive_risk": found_norm && !found_raw,
                }).to_string())
            }
            "diag_cell_raw" => {
                // 특정 셀의 원시 텍스트를 이스케이프 표기로 반환.
                // \n \r \t 가 실제로 어떻게 들어있는지 확인.
                let ti  = args["table_index"].as_u64().unwrap_or(0) as usize;
                let row = args["row"].as_u64().unwrap_or(0) as usize;
                let col = args["col"].as_u64().unwrap_or(0) as usize;
                let tables = self.scan_tables()?;
                if ti >= tables.len() {
                    anyhow::bail!("표 {ti}번 없음 (총 {}개)", tables.len());
                }
                let cell = tables[ti]
                    .cells.get(row).and_then(|r| r.get(col))
                    .map(|s| s.as_str()).unwrap_or("");
                let escaped: String = cell.chars().map(|c| match c {
                    '\n' => "\\n".to_string(),
                    '\r' => "\\r".to_string(),
                    '\t' => "\\t".to_string(),
                    c    => c.to_string(),
                }).collect();
                Ok(json!({
                    "table_index": ti, "row": row, "col": col,
                    "char_count":  cell.chars().count(),
                    "has_newline": cell.contains('\n'),
                    "has_cr":      cell.contains('\r'),
                    "has_tab":     cell.contains('\t'),
                    "raw_escaped": escaped,
                    "actual":      cell,
                }).to_string())
            }
            "diag_html_table_extract" => {
                // 특정 표의 원시 HTML 청크를 추출합니다.
                // parse_tables_from_html이 어떤 HTML을 처리하는지 직접 확인용.
                let target = args["table_index"].as_u64().unwrap_or(0) as usize;
                let limit  = args["limit"].as_u64().unwrap_or(2000) as usize;
                let html   = self.get_html()?;
                let lower  = html.to_lowercase();
                let mut count = 0usize;
                let mut depth = 0i32;
                let mut start_pos: Option<usize> = None;
                let mut i = 0usize;
                loop {
                    let next_tag = match lower[i..].find('<') {
                        Some(off) => i + off,
                        None => break,
                    };
                    let tag_end = match lower[next_tag..].find('>') {
                        Some(off) => next_tag + off + 1,
                        None => break,
                    };
                    let tag_inner = lower[next_tag + 1..tag_end - 1].trim();
                    let closing   = tag_inner.starts_with('/');
                    let tag_name  = if closing {
                        tag_inner.trim_start_matches('/').split_whitespace().next().unwrap_or("")
                    } else {
                        tag_inner.split_whitespace().next().unwrap_or("")
                    };
                    if tag_name == "table" && !closing {
                        if depth == 0 {
                            if count == target { start_pos = Some(next_tag); }
                            count += 1;
                        }
                        depth += 1;
                    } else if tag_name == "table" && closing {
                        depth -= 1;
                        if depth == 0 {
                            if let Some(sp) = start_pos {
                                let raw = &html[sp..tag_end];
                                let preview: String = raw.chars().take(limit).collect();
                                return Ok(json!({
                                    "table_index": target,
                                    "raw_html_chars": raw.chars().count(),
                                    "preview": preview,
                                }).to_string());
                            }
                        }
                    }
                    i = tag_end;
                }
                Ok(json!({"error": format!("표 {target}번 없음 (총 {count}개)")}).to_string())
            }
            "diag_initscan_gettext" => {
                // InitScan → GetText → ReleaseScan API 테스트.
                // VT_BYREF output param을 COM 레이어가 지원하는지 확인합니다.
                let hwp = self.hwp()?;
                let init_res = hwp.call("InitScan", vec![
                    Variant::I32(0x07), Variant::I32(0x0077),
                    Variant::Empty, Variant::Empty, Variant::Empty, Variant::Empty,
                ])?;
                let init_str = init_res.to_string_repr();
                // GetText는 (상태코드, 텍스트) 두 output param을 VT_BYREF로 반환 —
                // com_dispatch 레이어가 지원하면 텍스트가 들어옴
                let gettext_res = hwp.call("GetText", vec![]);
                let _ = hwp.call("ReleaseScan", vec![]);
                match gettext_res {
                    Ok(v)  => Ok(json!({
                        "init_result":   init_str,
                        "gettext_value": v.to_string_repr(),
                        "verdict": "GetText COM 호출 성공 — 반환값 확인 필요",
                    }).to_string()),
                    Err(e) => Ok(json!({
                        "init_result":  init_str,
                        "gettext_error": e.to_string(),
                        "verdict": "GetText COM 호출 실패 (VT_BYREF 미지원 가능성)",
                    }).to_string()),
                }
            }

            "diag_get_pos" => {
                // hwp.GetPos() VT_BYREF 테스트: 현재 커서 좌표 반환
                let hwp = self.hwp()?;
                match hwp.get_pos() {
                    Ok((list, para, pos)) => Ok(json!({
                        "list": list,
                        "para": para,
                        "pos": pos,
                        "verdict": if list == 0 && para == 0 && pos == 0 {
                            "모두 0 — 커서가 맨 앞이거나 VT_BYREF 미작동"
                        } else {
                            "비영값 수신 — VT_BYREF 정상 작동"
                        }
                    }).to_string()),
                    Err(e) => Ok(json!({ "error": e.to_string() }).to_string()),
                }
            }

            "diag_key_indicator" => {
                // KeyIndicator: HWP는 output param이 있으므로 두 방식 모두 시도
                let hwp = self.hwp()?;
                // 방식A: 인자 0개로 call (반환값에 정보 없을 수 있음)
                let call0 = hwp.call("KeyIndicator", vec![])
                    .map(|v| v.to_string_repr())
                    .map_err(|e| e.to_string());
                // 방식B: VT_BYREF 10개 (이전 실패 — 참고용)
                let byref10 = hwp.key_indicator()
                    .map_err(|e| e.to_string());
                Ok(json!({
                    "call_0args": call0,
                    "byref_10args": byref10,
                    "note": "call_0args가 성공하면 반환 VARIANT 값 확인, byref_10args는 파라미터 수 문제로 실패 예상"
                }).to_string())
            }

            "diag_set_pos" => {
                // hwp.SetPos(list, para, pos) 테스트: 지정 좌표로 커서 이동
                let list = args["list"].as_i64().unwrap_or(0) as i32;
                let para = args["para"].as_i64().unwrap_or(0) as i32;
                let pos  = args["pos"].as_i64().unwrap_or(0) as i32;
                let hwp = self.hwp()?;
                let before = hwp.get_pos().ok();
                let result = hwp.set_pos(list, para, pos);
                let after  = hwp.get_pos().ok();
                match result {
                    Ok(ok) => Ok(json!({
                        "set_pos_result": ok,
                        "before": before.map(|(l,p,po)| json!({"list":l,"para":p,"pos":po})),
                        "after":  after.map(|(l,p,po)|  json!({"list":l,"para":p,"pos":po})),
                        "verdict": if ok { "SetPos 성공" } else { "SetPos false 반환" },
                    }).to_string()),
                    Err(e) => Ok(json!({ "error": e.to_string() }).to_string()),
                }
            }

            "diag_initscan_1param" => {
                // InitScan 파라미터 개수 탐색 + CreateAction 방식도 시도
                let hwp = self.hwp()?;
                // 6개 파라미터 (option, rang, spara, spos, epara, epos) — 정식 시그니처
                let r1 = hwp.call("InitScan", vec![
                    Variant::I32(0x06), Variant::I32(0x0077),
                    Variant::I32(0), Variant::I32(0), Variant::I32(0), Variant::I32(0),
                ]);
                let _ = hwp.call("ReleaseScan", vec![]);
                // 2개 파라미터 (이전 방식 — 실패 예상)
                let r2 = hwp.call("InitScan", vec![Variant::I32(0), Variant::I32(0x0077)]);
                let _ = hwp.call("ReleaseScan", vec![]);
                // 0개 파라미터 — 실패 예상
                let r3 = hwp.call("InitScan", vec![]);
                let _ = hwp.call("ReleaseScan", vec![]);
                // CreateAction 방식 시도
                let r_action = hwp.call("CreateAction", vec![Variant::String("InitScan".to_string())])
                    .map(|v| v.to_string_repr())
                    .map_err(|e| e.to_string());
                Ok(json!({
                    "1param_0x37": r1.map(|v| v.to_string_repr()).map_err(|e| e.to_string()),
                    "0param":      r2.map(|v| v.to_string_repr()).map_err(|e| e.to_string()),
                    "3params":     r3.map(|v| v.to_string_repr()).map_err(|e| e.to_string()),
                    "create_action": r_action,
                }).to_string())
            }

            "diag_scan_table_positions" => {
                // InitScan → GetText+GetPos 루프로 셀 위치 맵 구축
                let hwp = self.hwp()?;

                let init_result = hwp.call("InitScan", vec![
                    Variant::I32(0x07), Variant::I32(0x0077),
                    Variant::Empty, Variant::Empty, Variant::Empty, Variant::Empty,
                ])?;
                let init_str = init_result.to_string_repr();

                let mut entries: Vec<serde_json::Value> = Vec::new();
                let mut call_count = 0;
                let max_calls = 500;

                loop {
                    if call_count >= max_calls { break; }
                    call_count += 1;

                    match hwp.get_text_scan() {
                        Err(e) => {
                            entries.push(json!({ "error": e.to_string(), "call": call_count }));
                            break;
                        }
                        Ok((text_state, text, ret_code)) => {
                            if ret_code == 0 || text_state == 0 { break; }
                            let pos_info = hwp.get_pos().ok()
                                .map(|(l,p,po)| json!({"list":l,"para":p,"pos":po}));
                            let preview: String = text.chars().take(40).collect();
                            entries.push(json!({
                                "call": call_count,
                                "text_state": text_state,
                                "text": preview,
                                "ret_code": ret_code,
                                "pos": pos_info,
                            }));
                            if entries.len() >= 30 { break; }
                        }
                    }
                }

                let _ = hwp.call("ReleaseScan", vec![]);
                Ok(json!({
                    "init_result": init_str,
                    "total_calls": call_count,
                    "entries": entries,
                }).to_string())
            }

            "diag_table_cell_walk" => {
                // TableRightCell + GetPos로 표의 모든 물리 셀 좌표를 기록합니다.
                // navigate_to_table → TableRightCell 반복 → GetPos 수집
                // SetPos 기반 셀 이동의 핵심 데이터 수집용.
                let table_index = args["table_index"].as_u64().unwrap_or(0) as usize;
                let hwp = self.hwp()?;

                // 표 첫 셀로 이동
                self.navigate_to_table(table_index)?;

                let first_pos = hwp.get_pos()?;
                let mut cells: Vec<serde_json::Value> = vec![
                    json!({ "phys_idx": 0, "pos": {"list": first_pos.0, "para": first_pos.1, "pos": first_pos.2} })
                ];

                // Run("TableRightCell")로 순회 (최대 200 셀)
                // TableRightCell은 직접 메서드가 아닌 HAction ID
                for i in 1..200usize {
                    let moved = hwp.call("Run", vec![
                        Variant::String("TableRightCell".to_string()),
                    ])?;
                    // false 또는 Empty 반환 = 표 끝
                    if !moved.as_bool().unwrap_or(true) { break; }
                    let p = hwp.get_pos()?;
                    cells.push(json!({
                        "phys_idx": i,
                        "pos": { "list": p.0, "para": p.1, "pos": p.2 }
                    }));
                    if cells.len() >= 100 { break; }
                }

                Ok(json!({
                    "table_index": table_index,
                    "physical_cells_found": cells.len(),
                    "cells": cells,
                }).to_string())
            }

            "diag_phys_structure" => {
                // parse_physical_tables 결과 + physical_cell_offset 계산을 노출합니다.
                // 실제 navigation 계산이 맞는지 확인용.
                let table_index = args["table_index"].as_u64().unwrap_or(0) as usize;
                let target_row  = args["target_row"].as_u64().unwrap_or(4) as usize;
                let target_col  = args["target_col"].as_u64().unwrap_or(2) as usize;

                let html = self.get_html()?;
                let raw_tables = parse_physical_tables(&html);
                let raw = match raw_tables.get(table_index) {
                    Some(r) => r,
                    None => return Ok(json!({"error": format!("표 {}번 없음", table_index)}).to_string()),
                };

                // 각 행의 물리 셀 수와 (colspan, rowspan)
                let rows_info: Vec<serde_json::Value> = raw.iter().enumerate().map(|(r, row)| {
                    let cells: Vec<serde_json::Value> = row.iter().map(|&(cs, rs)| json!({"colspan":cs,"rowspan":rs})).collect();
                    json!({"row": r, "phys_count": row.len(), "cells": cells})
                }).collect();

                // physical_cell_offset for anchor (row 0, col 0)
                let anchor_offset_res = physical_cell_offset(raw, 0, 0);
                // physical_cell_offset for target
                let target_offset_res = physical_cell_offset(raw, target_row, target_col);
                let delta = match (&anchor_offset_res, &target_offset_res) {
                    (Ok(a), Ok(t)) => Some(*t as isize - *a as isize),
                    _ => None,
                };

                // is_vcol_continuation for each vcol in target_row (0..7)
                let continuation_map: Vec<serde_json::Value> = (0..8usize).map(|vc| {
                    json!({"vcol": vc, "is_continuation": is_vcol_continuation(&raw_tables, table_index, target_row, vc)})
                }).collect();

                Ok(json!({
                    "table_index": table_index,
                    "total_rows": raw.len(),
                    "rows_physical": rows_info,
                    "anchor_offset": anchor_offset_res.map_err(|e| e.to_string()),
                    "target_offset": target_offset_res.map_err(|e| e.to_string()),
                    "delta_tableright": delta,
                    "continuation_map": continuation_map,
                }).to_string())
            }

            "probe_scan" => {
                // InitScan/GetText 실측 진단 도구 (v2)
                // Phase 1: GetFieldList → MoveToField → GetPos 로 필드 위치 사전 매핑
                // Phase 2: InitScan/GetText 루프 안에서 GetPos 호출 → 스캔 커서가 위치를 따르는지 확인
                // 두 위치가 일치하면 스캔 중 필드 감지가 가능함
                let max_events = args["max_events"].as_u64().unwrap_or(300) as usize;
                let hwp = self.hwp()?;

                // ── Phase 1: 필드 위치 사전 매핑 ──────────────────────────────────
                // GetFieldList(option=0, type=1) — type 1 = 누름틀(field control)
                let field_names = self.get_field_names();
                let mut field_pos_map: Vec<serde_json::Value> = Vec::new();

                for fname in &field_names {
                    // MoveToField(name, start=True, select=True, move=True)
                    // select=True: 필드 내용을 블록으로 선택 → GetSelectedPos 사용 가능
                    // 출처: 한컴 공식 포럼 (2024-08) — 내용 유무 확인 방법
                    let moved = hwp.call("MoveToField", vec![
                        Variant::String(fname.clone()),
                        Variant::Bool(true),  // start: 필드 시작으로 이동
                        Variant::Bool(true),  // select: 필드 내용 선택 (GetSelectedPos 활성화)
                        Variant::Bool(true),  // move: 실제 이동
                    ]).map(|v| v.as_bool().unwrap_or(false)).unwrap_or(false);

                    // GetPos: 편집 커서 위치 (필드가 속한 단락의 list/para/pos)
                    let pos = if moved { hwp.get_pos().ok() } else { None };

                    // GetSelectedPos: 선택된 필드 내용의 영역 좌표
                    // 내용 없음 → None, 내용 있음 → (spara, spos, epara, epos)
                    let sel_pos = if moved { hwp.get_selected_pos().ok().flatten() } else { None };

                    // CurFieldName: 편집 커서가 필드 안에 있을 때 이름 반환
                    // 이름 없는 누름틀이면 "" → "[unnamed]"
                    let cur_field_name_after_move: Option<String> = match hwp.call("CurFieldName", vec![]) {
                        Ok(v) => {
                            let s = v.to_string_repr();
                            if s == "Empty" { None }
                            else { Some(if s.is_empty() { "[unnamed]".to_string() } else { s }) }
                        }
                        Err(_) => None,
                    };

                    // GetFieldText: 필드의 현재 텍스트 값 (표/도형은 빈 문자열)
                    let field_text = hwp.call("GetFieldText", vec![
                        Variant::String(fname.clone()),
                    ]).map(|v| v.to_string_repr()).ok();

                    field_pos_map.push(json!({
                        "name":             fname,
                        "moved":            moved,
                        "field_text":       field_text,
                        "pos_list":         pos.map(|(l,_,_)| l),
                        "pos_para":         pos.map(|(_,p,_)| p),
                        "pos_char":         pos.map(|(_,_,c)| c),
                        "sel_spara":        sel_pos.map(|(sp,_,_,_)| sp),
                        "sel_spos":         sel_pos.map(|(_,ss,_,_)| ss),
                        "sel_epara":        sel_pos.map(|(_,_,ep,_)| ep),
                        "sel_epos":         sel_pos.map(|(_,_,_,es)| es),
                        "cur_field_after_move": cur_field_name_after_move,
                    }));
                }

                // ── Phase 2: InitScan/GetText 루프 ────────────────────────────────
                // option: 0x07 = maskChar|maskInline|maskCtrl (모든 컨트롤 + 누름틀)
                // rang:   0x0077 = scanSposDocument|scanEposDocument (문서 전체)
                let init_result = hwp.call("InitScan", vec![
                    Variant::I32(0x07),   // option: 모든 마스크
                    Variant::I32(0x0077), // rang: 문서 전체
                    Variant::Empty,       // spara (rang에 Specified 없으므로 무시)
                    Variant::Empty,       // spos
                    Variant::Empty,       // epara
                    Variant::Empty,       // epos
                ]).map(|v| v.to_string_repr())
                  .unwrap_or_else(|e| format!("ERR:{e}"));

                let mut events: Vec<serde_json::Value> = Vec::new();

                for seq in 0..max_events {
                    let scan = hwp.get_text_scan();
                    let (text_state, text, ret_code) = match scan {
                        Err(e) => {
                            events.push(json!({ "seq": seq, "error": e.to_string() }));
                            break;
                        }
                        Ok(t) => t,
                    };

                    // GetPos: 스캔 중 호출 — 편집 커서를 따르는지 vs 스캔 커서를 따르는지 측정
                    let pos_during_scan = hwp.get_pos().ok();

                    // CurFieldName: 편집 커서 기준
                    // 이름 없는 누름틀은 빈 문자열 "" 반환 → "[unnamed]" 로 표시
                    // 누름틀 밖이면 COM 오류 또는 VT_EMPTY → None
                    let cur_field_name: Option<String> = match hwp.call("CurFieldName", vec![]) {
                        Ok(v) => {
                            let s = v.to_string_repr();
                            if s == "Empty" {
                                None // VT_EMPTY = 누름틀 밖
                            } else {
                                Some(if s.is_empty() { "[unnamed]".to_string() } else { s })
                            }
                        }
                        Err(_) => None,
                    };

                    // CurCtrl: IDispatch 객체 → CtrlID 프로퍼티로 컨트롤 타입 확인
                    let ctrl_id = hwp.call("CurCtrl", vec![])
                        .ok()
                        .and_then(|ctrl_var| ctrl_var.as_object())
                        .and_then(|ctrl_obj| ctrl_obj.get("CtrlID").ok())
                        .map(|v| v.to_string_repr())
                        .and_then(|s| if s.is_empty() || s == "Empty" { None } else { Some(s) });

                    // 문자 경계를 지켜서 자르기 (바이트 슬라이스 패닉 방지)
                    let text_preview: String = text.chars().take(80).collect();
                    events.push(json!({
                        "seq":          seq,
                        "ret_code":     ret_code,
                        "text_state":   text_state,
                        "text":         text_preview,
                        "pos_list":     pos_during_scan.map(|(l,_,_)| l),
                        "pos_para":     pos_during_scan.map(|(_,p,_)| p),
                        "pos_char":     pos_during_scan.map(|(_,_,c)| c),
                        "cur_field_name": cur_field_name,
                        "ctrl_id":      ctrl_id,
                    }));

                    // ret_code 0 또는 text_state 0 = 스캔 종료
                    if ret_code == 0 || text_state == 0 { break; }
                }

                let _ = hwp.call("ReleaseScan", vec![]);

                // 요약 통계
                let state_counts = {
                    let mut m: std::collections::HashMap<i32, usize> = Default::default();
                    for ev in &events {
                        if let Some(s) = ev["text_state"].as_i64() {
                            *m.entry(s as i32).or_default() += 1;
                        }
                    }
                    let mut counts: Vec<_> = m.into_iter()
                        .map(|(k,v)| json!({"state": k, "count": v}))
                        .collect();
                    counts.sort_by_key(|v| v["state"].as_i64().unwrap_or(0));
                    counts
                };

                // 필드 위치와 스캔 이벤트 위치 교차 분석
                let field_match_analysis: Vec<serde_json::Value> = field_pos_map.iter().map(|fp| {
                    let fl = fp["pos_list"].as_i64();
                    let fp_para = fp["pos_para"].as_i64();
                    // 스캔 이벤트 중 같은 (list, para) 를 가진 이벤트 찾기
                    let matching_events: Vec<&serde_json::Value> = events.iter()
                        .filter(|ev| {
                            ev["pos_list"].as_i64() == fl &&
                            ev["pos_para"].as_i64() == fp_para
                        })
                        .collect();
                    json!({
                        "field_name": fp["name"],
                        "field_pos": {"list": fl, "para": fp_para},
                        "matching_scan_events": matching_events.len(),
                        "sample_texts": matching_events.iter().take(3)
                            .map(|ev| ev["text"].clone())
                            .collect::<Vec<_>>(),
                    })
                }).collect();

                Ok(json!({
                    "init_result":         init_result,
                    "total_events":        events.len(),
                    "state_counts":        state_counts,
                    "field_count":         field_names.len(),
                    "field_names":         field_names,
                    "field_pos_map":       field_pos_map,
                    "field_match_analysis": field_match_analysis,
                    "events":              events,
                }).to_string())
            }

            // ── 누름틀 필드 관련 Tool ─────────────────────────────────────

            // HTML FieldStart:/FieldEnd: 파싱으로 누름틀 목록 반환
            // 이름 없는 누름틀도 감지됨 (GetFieldList와 달리)
            // placeholder 판별: font-style:italic + color:#ff0000
            "get_document_fields" => {
                let html = self.get_html()?;
                let fields = Self::parse_fields_from_html(&html);

                let result: Vec<serde_json::Value> = fields.iter().map(|f| json!({
                    "index":          f.index,
                    "name":           if f.name.is_empty() { json!(null) } else { json!(f.name) },
                    "content":        f.content,
                    "text_after":     f.text_after.chars().take(30).collect::<String>(),
                    "is_placeholder": f.is_placeholder,
                    "status":         if f.is_placeholder { "미입력" } else { "입력됨" },
                })).collect();

                Ok(json!({
                    "count":  fields.len(),
                    "fields": result,
                }).to_string())
            }

            // 누름틀에 값 설정
            // 전략: ForwardFind(SearchCtrl=true)로 누름틀 내부 텍스트 검색
            //       → GetPos.list != 0 이면 필드 내부 → PutString으로 교체
            // MoveToField / CurFieldName / MoveNextChar list-change — 모두 unnamed 필드 미지원 확인됨
            "fill_field" => {
                let placeholder = args["placeholder"].as_str()
                    .ok_or_else(|| anyhow::anyhow!("placeholder 파라미터가 없습니다"))?
                    .to_string();
                let value = args["value"].as_str()
                    .ok_or_else(|| anyhow::anyhow!("value 파라미터가 없습니다"))?
                    .to_string();

                // HTML에서 필드 메타 정보 (index, name, 존재 확인)
                let html = self.get_html()?;
                let fields = Self::parse_fields_from_html(&html);
                let target = fields.iter()
                    .find(|f| f.content == placeholder || (!f.name.is_empty() && f.name == placeholder));
                let Some(field) = target else {
                    let available: Vec<&str> = fields.iter().map(|f| f.content.as_str()).collect();
                    anyhow::bail!("'{placeholder}' 와 일치하는 누름틀 없음. 사용 가능: {available:?}");
                };
                let field_index = field.index;
                let field_name  = field.name.clone();

                // 목표 필드(field_index) 앞에 동일 placeholder를 가진 누름틀이 몇 개인지 계산
                // → ForwardFind 루프에서 그 수만큼 건너뛰어야 정확한 누름틀에 도달
                let skip_count = fields.iter()
                    .filter(|f| f.index < field_index && (f.content == placeholder || (!f.name.is_empty() && f.name == placeholder)))
                    .count();

                let hwp = self.hwp()?;
                // 찾기/바꾸기 관련 "문서 끝까지 찾았습니다" 다이얼로그 자동 응답
                // (다이얼로그가 COM Execute 호출을 블록하면 exec_ok=false가 됨)
                let _ = hwp.call("SetMessageBoxMode", vec![Variant::I32(0x00F0)]);
                let _ = hwp.call("Run", vec![Variant::String("MoveDocBegin".to_string())]);

                // ForwardFind를 반복 호출해서 placeholder를 찾음
                // GetPos.list != 0 이면 커서가 누름틀 내부 (성공)
                // GetPos.list == 0 이면 본문 텍스트 (건너뜀, 다시 탐색)
                let mut find_log: Vec<Value> = Vec::new();
                let mut found = false;
                let mut put_ok = false;
                let mut put_err: Option<String> = None;
                let mut sel_pos: Option<(i32,i32,i32,i32)> = None;
                let mut new_content = String::new();
                // 이미 시도한 (list,para,char) → 순환 감지용
                let mut tried_positions: Vec<(i32, i32, i32)> = Vec::new();

                // 누름틀 컨텍스트 판별
                let non_field_ctrls: &[&str] = &["표", "머리말", "꼬리말", "각주", "미주", "글상자"];

                'find_loop: for iter in 0..20usize {
                    let act = match hwp.call("CreateAction", vec![
                        Variant::String("ForwardFind".to_string()),
                    ]).ok().and_then(|v| v.as_object()) {
                        Some(a) => a,
                        None => {
                            find_log.push(json!({"iter": iter, "error": "CreateAction 실패"}));
                            break;
                        }
                    };
                    let ps = match act.call("CreateSet", vec![]).ok().and_then(|v| v.as_object()) {
                        Some(p) => p,
                        None => {
                            find_log.push(json!({"iter": iter, "error": "CreateSet 실패"}));
                            break;
                        }
                    };

                    let setitems: &[(&str, Variant)] = &[
                        ("FindString",    Variant::String(placeholder.clone())),
                        ("ReplaceString", Variant::String(String::new())),
                        ("IgnoreCase",    Variant::Bool(false)),
                        ("AllWordReplace",Variant::Bool(false)),
                        ("Direction",     Variant::I32(0)),
                        ("FindRegExp",    Variant::Bool(false)),
                        ("SearchTbl",     Variant::Bool(true)),
                        ("SearchCtrl",    Variant::Bool(true)),
                        ("FindOpt",       Variant::I32(0x3F)),
                    ];
                    for (k, v) in setitems {
                        let _ = ps.call("SetItem", vec![
                            Variant::String(k.to_string()),
                            v.clone(),
                        ]);
                    }

                    let exec_ok = act.call("Execute", vec![Variant::Object(ps)]).is_ok();
                    let pos = hwp.get_pos().ok();
                    let cur_list = pos.map(|(l,_,_)| l).unwrap_or(0);
                    let ctrl_name = hwp.key_indicator().unwrap_or_default();
                    let pos_key = pos.unwrap_or((0,0,0));

                    find_log.push(json!({
                        "iter":      iter,
                        "exec_ok":   exec_ok,
                        "list":      cur_list,
                        "para":      pos.map(|(_,p,_)| p),
                        "char":      pos.map(|(_,_,c)| c),
                        "ctrl_name": ctrl_name,
                    }));

                    if !exec_ok { break; }

                    // 순환 감지: 이미 시도한 위치면 전체 탐색 완료
                    if tried_positions.contains(&pos_key) {
                        break;
                    }

                    if cur_list != 0 && !non_field_ctrls.contains(&ctrl_name.as_str()) {
                        // 누름틀 후보 → 즉시 Paste 시도 후 HTML로 검증
                        tried_positions.push(pos_key);
                        sel_pos = hwp.get_selected_pos().ok().flatten();

                        let paste_result = crate::com_dispatch::set_clipboard_text(&value)
                            .and_then(|_| hwp.call("Run", vec![Variant::String("Paste".to_string())])
                                .map_err(|e| anyhow::anyhow!(e)));
                        put_ok = paste_result.is_ok();
                        put_err = paste_result.err().map(|e| e.to_string());

                        if put_ok {
                            // 올바른 필드가 바뀌었는지 HTML로 즉시 확인
                            let check_html = self.get_html()?;
                            let check_fields = Self::parse_fields_from_html(&check_html);
                            new_content = check_fields.get(field_index)
                                .map(|f| f.content.clone())
                                .unwrap_or_default();

                            if new_content == value {
                                found = true;
                                break 'find_loop; // 성공
                            }

                            // 틀린 위치에 붙여넣음 → Undo 후 계속 탐색
                            let _ = hwp.call("Run", vec![Variant::String("Undo".to_string())]);
                            // 현재 위치 지나쳐야 재발견 안 됨: 커서를 placeholder 길이만큼 전진
                            for _ in 0..=placeholder.chars().count() {
                                let _ = hwp.call("Run", vec![Variant::String("MoveNextChar".to_string())]);
                            }
                        }
                    }
                }

                // 메시지박스 모드 복원
                let _ = hwp.call("SetMessageBoxMode", vec![Variant::I32(0x0000)]);

                Ok(json!({
                    "success":     found,
                    "method":      "ForwardFind+Paste+Verify",
                    "field_index": field_index,
                    "field_name":  if field_name.is_empty() { json!(null) } else { json!(field_name) },
                    "placeholder": placeholder,
                    "value":       value,
                    "new_content": new_content,
                    "put_ok":      put_ok,
                    "put_err":     put_err,
                    "sel_pos":     sel_pos.map(|(sp,ss,ep,es)| json!({"spara":sp,"spos":ss,"epara":ep,"epos":es})),
                    "find_log":    find_log,
                }).to_string())
            }

            // 커서 이동 + CurFieldName 진단 도구
            // 1. MoveDocBegin → GetPos
            // 2. MoveNextChar × n → 각 단계 GetPos + CurFieldName 기록
            // 3. 커서가 이동하는지, 필드를 감지하는지 확인
            "diag_nav" => {
                let steps = args["steps"].as_u64().unwrap_or(20) as usize;
                let hwp = self.hwp()?;

                let _ = hwp.call("Run", vec![Variant::String("MoveDocBegin".to_string())]);
                let initial_pos = hwp.get_pos().ok();

                let mut log = Vec::new();
                let mut prev_list: i32 = 0;
                for step in 0..steps {
                    let _ = hwp.call("Run", vec![Variant::String("MoveNextChar".to_string())]);
                    let pos = hwp.get_pos().ok();
                    let cur_list = pos.map(|(l,_,_)| l).unwrap_or(0);
                    let list_changed = cur_list != prev_list;
                    prev_list = cur_list;
                    log.push(json!({
                        "step":         step,
                        "list":         cur_list,
                        "para":         pos.map(|(_,p,_)| p),
                        "char":         pos.map(|(_,_,c)| c),
                        "list_changed": list_changed,
                    }));
                }

                Ok(json!({
                    "initial_pos": {
                        "list": initial_pos.map(|(l,_,_)| l),
                        "para": initial_pos.map(|(_,p,_)| p),
                        "char": initial_pos.map(|(_,_,c)| c),
                    },
                    "steps": log,
                }).to_string())
            }

            "get_field_list" => {
                let hwp = self.hwp()?;

                // type=0: 모든 필드, type=1: 누름틀만 — 둘 다 시도해서 비교
                let try_field_list = |opt: i32, typ: i32| {
                    hwp.call("GetFieldList", vec![Variant::I32(opt), Variant::I32(typ)])
                       .map(|v| v.to_string_repr())
                       .unwrap_or_default()
                };

                let raw_t0 = try_field_list(0, 0);
                let raw_t1 = try_field_list(0, 1);
                let raw_t3 = try_field_list(0, 3); // 한컴 일부 버전: type=3=유형 포함

                // 구분자 진단 — \x02 외에 다른 구분자 사용 여부 확인
                let sep_counts = |s: &str| json!({
                    "len":   s.len(),
                    "0x02":  s.chars().filter(|&c| c == '\x02').count(),
                    "0x01":  s.chars().filter(|&c| c == '\x01').count(),
                    "newline": s.chars().filter(|&c| c == '\n').count(),
                    "preview": s.chars().take(200).collect::<String>(),
                });

                let fields_t0: Vec<&str> = raw_t0.split('\x02').filter(|s| !s.is_empty()).collect();
                let fields_t1: Vec<&str> = raw_t1.split('\x02').filter(|s| !s.is_empty()).collect();
                let fields_t3: Vec<&str> = raw_t3.split('\x02').filter(|s| !s.is_empty()).collect();

                Ok(json!({
                    "type0_all":    { "count": fields_t0.len(), "fields": fields_t0, "raw": sep_counts(&raw_t0) },
                    "type1_field":  { "count": fields_t1.len(), "fields": fields_t1, "raw": sep_counts(&raw_t1) },
                    "type3_typed":  { "count": fields_t3.len(), "fields": fields_t3, "raw": sep_counts(&raw_t3) },
                }).to_string())
            }

            // GetFieldText(name) — 누름틀의 현재 텍스트 값 반환
            // 표/도형이 있는 필드는 빈 문자열 반환 (GetFieldText 한계)
            // 모든 필드의 현재 값을 한 번에 반환합니다
            "get_field_values" => {
                let hwp = self.hwp()?;
                let raw = hwp.call("GetFieldList", vec![
                    Variant::I32(0),
                    Variant::I32(1),
                ]).map(|v| v.to_string_repr()).unwrap_or_default();

                let fields: Vec<&str> = raw.split('\x02').filter(|s| !s.is_empty()).collect();
                let mut values: Vec<serde_json::Value> = Vec::new();

                for fname in &fields {
                    let text = hwp.call("GetFieldText", vec![
                        Variant::String(fname.to_string()),
                    ]).map(|v| v.to_string_repr()).unwrap_or_default();

                    values.push(json!({
                        "name":  fname,
                        "value": text,
                        "empty": text.is_empty(),
                    }));
                }

                Ok(json!({
                    "count":  fields.len(),
                    "fields": values,
                }).to_string())
            }

            // SetFieldText(name, value) — 누름틀에 텍스트 값 설정
            // 필드에 표/도형이 포함된 경우 이 API로는 설정 불가 (텍스트 전용)
            "set_field" => {
                let name = args["name"].as_str()
                    .ok_or_else(|| anyhow::anyhow!("name 파라미터가 없습니다"))?;
                let value = args["value"].as_str()
                    .ok_or_else(|| anyhow::anyhow!("value 파라미터가 없습니다"))?;

                let hwp = self.hwp()?;
                let ok = hwp.call("SetFieldText", vec![
                    Variant::String(name.to_string()),
                    Variant::String(value.to_string()),
                ]).map(|v| v.as_bool().unwrap_or(true))
                  .unwrap_or(false);

                Ok(json!({
                    "success": ok,
                    "name":    name,
                    "value":   value,
                }).to_string())
            }

            other => Ok(format!("❌ 알 수 없는 도구: {other}")),
        }
    }
}

// ──────────────────────────────────────────────────────────────
// 표 데이터 구조체
// ──────────────────────────────────────────────────────────────

struct TableData {
    rows: usize,
    cols: usize,
    cells: Vec<Vec<String>>,
}

// ──────────────────────────────────────────────────────────���───
// HTML 파싱 헬퍼 (Python PoC의 _parse_all_tables / _text_exists 포트)
// ──────────────────────────────────────────────────────────────

/// HTML에서 모든 최상위 `<table>` 을 파싱하여 TableData 목록으로 반환합니다.
///
/// colspan / rowspan 을 처리하여 시각적 그리드 좌표와 일치하는 cells 배열을 반환합니다.
/// HWP HTML은 `<br>` 없이 `<p>` 태그로 단락을 구분하므로, `<p>` 시작 시 '\n' 구분자를 삽입합니다.
fn parse_tables_from_html(html: &str) -> Vec<TableData> {
    let lower = html.to_lowercase();
    let mut tables: Vec<TableData> = Vec::new();
    let mut depth = 0i32;

    // raw_rows: (cell_text, colspan, rowspan) 의 행 목록
    let mut raw_rows: Vec<Vec<(String, usize, usize)>> = Vec::new();
    let mut cur_row: Vec<(String, usize, usize)> = Vec::new();
    let mut in_cell = false;
    let mut cell_buf = String::new();
    let mut cur_colspan = 1usize;
    let mut cur_rowspan = 1usize;
    let mut i = 0usize;

    while i < html.len() {
        let next_tag = match lower[i..].find('<') {
            Some(off) => i + off,
            None => {
                if in_cell && depth == 1 { cell_buf.push_str(&html[i..]); }
                break;
            }
        };

        if in_cell && depth == 1 && next_tag > i {
            cell_buf.push_str(&html[i..next_tag]);
        }

        let tag_end = match lower[next_tag..].find('>') {
            Some(off) => next_tag + off + 1,
            None => break,
        };

        let tag_inner = lower[next_tag + 1..tag_end - 1].trim();
        let closing = tag_inner.starts_with('/');
        let tag_name = if closing {
            tag_inner.trim_start_matches('/').split_whitespace().next().unwrap_or("")
        } else {
            tag_inner.split_whitespace().next().unwrap_or("")
        };

        match (tag_name, closing) {
            ("table", false) => {
                depth += 1;
                if depth == 1 {
                    raw_rows.clear();
                    cur_row.clear();
                    in_cell = false;
                    cell_buf.clear();
                }
            }
            ("table", true) => {
                if depth == 1 {
                    if in_cell {
                        cur_row.push((decode_html_entities(cell_buf.trim()), cur_colspan, cur_rowspan));
                        cell_buf.clear(); in_cell = false;
                    }
                    if !cur_row.is_empty() { raw_rows.push(std::mem::take(&mut cur_row)); }
                    if !raw_rows.is_empty() {
                        let grid = expand_grid_with_spans(&raw_rows);
                        let row_count = grid.len();
                        let col_count = grid.iter().map(|r| r.len()).max().unwrap_or(0);
                        tables.push(TableData { rows: row_count, cols: col_count, cells: grid });
                        raw_rows.clear();
                    }
                }
                depth -= 1;
            }
            ("tr", false) => {
                if depth == 1 && !cur_row.is_empty() {
                    raw_rows.push(std::mem::take(&mut cur_row));
                }
            }
            ("tr", true) => {
                if depth == 1 {
                    if in_cell {
                        cur_row.push((decode_html_entities(cell_buf.trim()), cur_colspan, cur_rowspan));
                        cell_buf.clear(); in_cell = false;
                    }
                    if !cur_row.is_empty() { raw_rows.push(std::mem::take(&mut cur_row)); }
                }
            }
            ("td", false) | ("th", false) if depth == 1 => {
                in_cell = true;
                cell_buf.clear();
                cur_colspan = parse_span_attr(tag_inner, "colspan").max(1);
                cur_rowspan = parse_span_attr(tag_inner, "rowspan").max(1);
            }
            ("td", true) | ("th", true) if depth == 1 && in_cell => {
                cur_row.push((decode_html_entities(cell_buf.trim()), cur_colspan, cur_rowspan));
                cell_buf.clear(); in_cell = false;
                cur_colspan = 1; cur_rowspan = 1;
            }
            // HWP HTML은 <p> 태그로 셀 내 단락을 구분합니다 (<br> 미사용).
            // <p> 열릴 때 기존 내용이 있으면 '\n' 구분자를 삽입합니다.
            ("p", false) if in_cell && depth == 1 => {
                let trimmed = cell_buf.trim_end().to_string();
                if !trimmed.is_empty() {
                    cell_buf = trimmed;
                    cell_buf.push('\n');
                }
            }
            ("br", _) if in_cell && depth == 1 => {
                cell_buf.push('\n');
            }
            _ => {}
        }

        i = tag_end;
    }

    tables
}

/// 시각적 열 vcol의 HWP 물리 셀 유형을 반환합니다.
///
/// - Physical   : 이 행에서 새로 시작하는 HTML `<td>` — 커서가 멈추는 독립 셀
/// - RowspanCont: 이전 행의 rowspan이 이 행까지 연장되는 셀 — HWP는 이 셀에도 커서를 세움
/// - ColspanCont: 같은 행 colspan 확장 열 — HWP는 이 위치를 별개 셀로 취급하지 않음
#[derive(Debug, PartialEq, Eq)]
enum VcolType {
    Physical,
    RowspanCont,
    ColspanCont,
}

fn vcol_cell_type(
    raw_tables: &[Vec<Vec<(usize, usize)>>],
    table_index: usize,
    row_idx: usize,
    target_vcol: usize,
) -> VcolType {
    let raw = match raw_tables.get(table_index) {
        Some(t) => t,
        None => return VcolType::Physical,
    };

    // pending: first_vcol → (remaining_rows, colspan)
    let mut pending: std::collections::BTreeMap<usize, (usize, usize)> = Default::default();

    for r in 0..=row_idx {
        let row = match raw.get(r) {
            Some(row) => row,
            None => return VcolType::Physical,
        };

        let mut vc = 0usize;
        let mut html_idx = 0usize;

        loop {
            if let Some(&(_rem, cs)) = pending.get(&vc) {
                if r == row_idx && target_vcol >= vc && target_vcol < vc + cs {
                    // target_vcol == vc : HWP 커서가 멈추는 첫 번째 vcol → RowspanCont
                    // target_vcol  > vc : 같은 셀의 colspan 내부 → ColspanCont (HWP 스텝 아님)
                    return if target_vcol == vc { VcolType::RowspanCont } else { VcolType::ColspanCont };
                }
                vc += cs;
            } else if html_idx < row.len() {
                let (cs, rs) = row[html_idx];
                html_idx += 1;
                if r == row_idx {
                    if target_vcol == vc { return VcolType::Physical; }
                    if target_vcol > vc && target_vcol < vc + cs { return VcolType::ColspanCont; }
                }
                if rs > 1 { pending.insert(vc, (rs, cs)); }
                vc += cs;
            } else {
                let next_pend = pending.keys().find(|&&k| k >= vc).copied();
                match next_pend {
                    Some(pvc) => { vc = pvc; }
                    None => break,
                }
            }
        }

        if r < row_idx {
            let mut new_p: std::collections::BTreeMap<usize, (usize, usize)> = Default::default();
            for (k, (rem, cs)) in pending {
                if rem > 1 { new_p.insert(k, (rem - 1, cs)); }
            }
            pending = new_p;
        }
    }

    VcolType::Physical
}

/// ColspanCont 여부만 반환합니다 (diag_phys_structure 출력용).
/// RowspanCont는 HWP 물리 셀이므로 false 반환.
fn is_vcol_continuation(
    raw_tables: &[Vec<Vec<(usize, usize)>>],
    table_index: usize,
    row_idx: usize,
    vcol: usize,
) -> bool {
    matches!(vcol_cell_type(raw_tables, table_index, row_idx, vcol), VcolType::ColspanCont)
}

/// HTML에서 각 표의 물리 셀 구조를 파싱합니다. (colspan, rowspan)만 추출.
/// 반환: tables[table_idx][row_idx][phys_col_idx] = (colspan, rowspan)
fn parse_physical_tables(html: &str) -> Vec<Vec<Vec<(usize, usize)>>> {
    let lower = html.to_lowercase();
    let mut tables: Vec<Vec<Vec<(usize, usize)>>> = Vec::new();
    let mut depth = 0i32;
    let mut cur_table: Vec<Vec<(usize, usize)>> = Vec::new();
    let mut cur_row: Vec<(usize, usize)> = Vec::new();
    let mut in_row = false;
    let mut i = 0usize;

    while let Some(off) = lower[i..].find('<') {
        let start = i + off;
        let end = match lower[start..].find('>') {
            Some(o) => start + o + 1,
            None => break,
        };
        let tag_inner = lower[start + 1..end - 1].trim();
        let closing = tag_inner.starts_with('/');
        let tag_name = if closing {
            tag_inner.trim_start_matches('/').split_whitespace().next().unwrap_or("")
        } else {
            tag_inner.split_whitespace().next().unwrap_or("")
        };

        match (tag_name, closing) {
            ("table", false) => {
                depth += 1;
                if depth == 1 {
                    cur_table.clear();
                    in_row = false;
                }
            }
            ("table", true) => {
                if depth == 1 {
                    if in_row && !cur_row.is_empty() {
                        cur_table.push(std::mem::take(&mut cur_row));
                    }
                    tables.push(std::mem::take(&mut cur_table));
                }
                depth -= 1;
            }
            ("tr", false) if depth == 1 => {
                if in_row && !cur_row.is_empty() {
                    cur_table.push(std::mem::take(&mut cur_row));
                }
                in_row = true;
            }
            ("tr", true) if depth == 1 => {
                if !cur_row.is_empty() {
                    cur_table.push(std::mem::take(&mut cur_row));
                }
                in_row = false;
            }
            ("td", false) | ("th", false) if depth == 1 && in_row => {
                let orig_tag = &html[start..end];
                let orig_lower = orig_tag.to_lowercase();
                let colspan = parse_span_attr(&orig_lower, "colspan").max(1);
                let rowspan = parse_span_attr(&orig_lower, "rowspan").max(1);
                cur_row.push((colspan, rowspan));
            }
            _ => {}
        }

        i = end;
    }

    tables
}

/// 표 첫 셀(물리 오프셋 0)부터 시각적 (target_row, target_vcol)까지
/// TableRightCell 이동 횟수를 계산합니다.
///
/// HWP는 rowspan 셀을 span 하는 모든 행에서 다시 방문합니다.
/// 따라서 각 행의 HWP 스텝 수 = (해당 행의 active pending rowspan 셀 수) + (HTML <td> 수)
/// ColspanCont는 HWP 스텝이 아니므로 count 0.
fn physical_cell_offset(
    raw: &[Vec<(usize, usize)>],
    target_row: usize,
    target_vcol: usize,
) -> anyhow::Result<usize> {
    // pending: first_vcol → (remaining_rows, colspan)
    let mut pending: std::collections::BTreeMap<usize, (usize, usize)> = Default::default();
    let mut total = 0usize;

    for row_idx in 0..=target_row {
        let row = raw.get(row_idx)
            .ok_or_else(|| anyhow::anyhow!("row {} out of bounds (table has {} rows)", row_idx, raw.len()))?;

        let mut hwp_step = 0usize;
        let mut vc = 0usize;
        let mut html_idx = 0usize;
        let mut found_step: Option<usize> = None;

        loop {
            if let Some(&(_rem, cs)) = pending.get(&vc) {
                // rowspan 재방문 — 1 HWP 스텝
                if row_idx == target_row && target_vcol >= vc && target_vcol < vc + cs {
                    found_step = Some(hwp_step);
                    break;
                }
                hwp_step += 1;
                vc += cs;
            } else if html_idx < row.len() {
                let (cs, rs) = row[html_idx];
                html_idx += 1;
                if row_idx == target_row && target_vcol >= vc && target_vcol < vc + cs {
                    found_step = Some(hwp_step);
                    break;
                }
                if rs > 1 { pending.insert(vc, (rs, cs)); }
                hwp_step += 1;
                vc += cs;
            } else {
                // HTML 셀 소진 — 남은 pending 중 더 높은 vcol 있으면 건너뜀
                let next_pend = pending.keys().find(|&&k| k >= vc).copied();
                match next_pend {
                    Some(pvc) => { vc = pvc; }
                    None => break,
                }
            }
        }

        if row_idx == target_row {
            return found_step
                .map(|s| total + s)
                .ok_or_else(|| anyhow::anyhow!("셀 ({},{})를 표에서 찾을 수 없습니다", target_row, target_vcol));
        }

        total += hwp_step;

        // 다음 행을 위해 pending rowspan 1 감소
        let mut new_p: std::collections::BTreeMap<usize, (usize, usize)> = Default::default();
        for (k, (rem, cs)) in pending {
            if rem > 1 { new_p.insert(k, (rem - 1, cs)); }
        }
        pending = new_p;
    }

    anyhow::bail!("unreachable")
}

/// colspan / rowspan 을 반영하여 원시 행 목록을 시각적 그리드로 확장합니다.
///
/// rowspan 이 있는 셀은 이후 행의 같은 열에 동일 텍스트가 들어갑니다.
/// colspan 이 있는 셀은 첫 열에만 텍스트, 나머지 열은 빈 문자열로 채웁니다.
fn expand_grid_with_spans(raw_rows: &[Vec<(String, usize, usize)>]) -> Vec<Vec<String>> {
    // pending[col] = (remaining_rows, text)
    let mut pending: std::collections::BTreeMap<usize, (usize, String)> = Default::default();
    let mut result: Vec<Vec<String>> = Vec::new();

    for raw_row in raw_rows {
        let mut row_map: std::collections::BTreeMap<usize, String> = Default::default();
        let mut col = 0usize;
        let mut html_idx = 0usize;

        loop {
            // 현재 col이 rowspan 점유 중이면 pending 값 사용
            if let Some((rem, text)) = pending.get(&col).cloned() {
                row_map.insert(col, text.clone());
                if rem <= 1 { pending.remove(&col); } else { pending.insert(col, (rem - 1, text)); }
                col += 1;
                continue;
            }

            if html_idx < raw_row.len() {
                // 다음 HTML 셀 배치
                let (text, colspan, rowspan) = &raw_row[html_idx];
                html_idx += 1;
                for c in 0..*colspan {
                    let cell_text = if c == 0 { text.clone() } else { String::new() };
                    row_map.insert(col + c, cell_text.clone());
                    if *rowspan > 1 {
                        pending.insert(col + c, (*rowspan - 1, cell_text));
                    }
                }
                col += colspan;
            } else {
                // HTML 셀 소진 — 남은 pending 열이 있으면 건너뜀
                let next_pending = pending.keys().find(|&&k| k >= col).copied();
                match next_pending {
                    Some(pc) => col = pc,
                    None => break,
                }
            }
        }

        if row_map.is_empty() { continue; }
        let max_col = *row_map.keys().next_back().unwrap();
        let row: Vec<String> = (0..=max_col)
            .map(|c| row_map.get(&c).cloned().unwrap_or_default())
            .collect();
        result.push(row);
    }

    result
}

/// HTML 태그 속성에서 colspan / rowspan 값을 파싱합니다.
/// tag 문자열은 이미 lowercase 처리된 상태여야 합니다.
fn parse_span_attr(tag: &str, attr: &str) -> usize {
    let search = format!("{attr}=");
    let pos = match tag.find(&search) {
        Some(p) => p + search.len(),
        None => return 1,
    };
    let after = tag[pos..].trim_start_matches('"').trim_start_matches('\'');
    let digits: String = after.chars().take_while(|c| c.is_ascii_digit()).collect();
    digits.parse().unwrap_or(1)
}

/// 문자 단위로 N자까지 자르고 초과 시 "..." 추가. 바이트 슬라이싱 대신 사용.
fn truncate_chars(s: &str, max_chars: usize) -> String {
    let mut chars = s.chars();
    let collected: String = chars.by_ref().take(max_chars).collect();
    if chars.next().is_some() {
        format!("{}...", collected)
    } else {
        collected
    }
}

/// HTML 특수문자 엔티티 디코딩
fn decode_html_entities(s: &str) -> String {
    s.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&nbsp;", " ")    // non-breaking space → 일반 공백
        .replace("&middot;", "·") // 가운뎃점
        .replace("&bull;", "•")
        .replace("&ndash;", "–")
        .replace("&mdash;", "—")
        .replace("&laquo;", "«")
        .replace("&raquo;", "»")
        .replace("&#39;", "'")
        .replace("&#160;", " ")   // non-breaking space numeric
}

/// HWP HTML에서 파싱된 누름틀 필드 정보
#[derive(Debug)]
struct HtmlField {
    index: usize,         // 문서 내 순서 (0-based)
    name: String,         // 필드 이름 (비어있으면 unnamed)
    content: String,      // 현재 내용 (placeholder 또는 실제 값)
    text_after: String,   // FieldEnd 직후 리터럴 텍스트 (find_replace 문맥용)
    is_placeholder: bool, // italic+red = 미입력 placeholder
}

/// HTML 청크(태그 포함)에서 plain text만 추출
fn extract_plain_text_from_html_chunk(html: &str) -> String {
    let mut out = String::new();
    let mut in_tag = false;
    for ch in html.chars() {
        match ch {
            '<' => { in_tag = true; }
            '>' => { in_tag = false; }
            _ if !in_tag => { out.push(ch); }
            _ => {}
        }
    }
    decode_html_entities(&out)
}

/// HTML에 특정 텍스트가 포함되어 있는지 확인합니다.
/// case_sensitive=false 이면 대소문자 무시 비교.
/// ForwardFind 호출 전 사전 검증에 사용 (다이얼로그 블로킹 방지).
fn html_contains_text(html: &str, text: &str, case_sensitive: bool) -> bool {
    // HTML 태그를 제거한 plain text에서 검색
    let plain = html_to_plain_text(html);
    // 검색 패턴과 plain text 양쪽의 공백/개행을 정규화하여 비교
    // (HWP HTML에서 개행은 태그 경계로 표현되므로 \n은 공백으로 처리)
    let normalize = |s: &str| -> String {
        s.split_whitespace().collect::<Vec<_>>().join(" ")
    };
    let plain_norm = normalize(&plain);
    let text_norm = normalize(text);
    if case_sensitive {
        plain_norm.contains(&text_norm)
    } else {
        plain_norm.to_lowercase().contains(&text_norm.to_lowercase())
    }
}

/// HTML에서 태그를 제거하고 plain text를 반환합니다.
///
/// 블록 요소(<p> <br> <tr> <td> <div>)만 개행/공백 삽입.
/// <a> <span> 같은 인라인 태그는 공백 삽입 없이 처리.
/// → "학년</span><a ...>학년 (" 이 "학년학년 (" 으로 올바르게 합쳐짐.
fn html_to_plain_text(html: &str) -> String {
    let mut out = String::with_capacity(html.len());
    let mut in_tag = false;
    let mut tag_buf = String::new();
    for ch in html.chars() {
        match ch {
            '<' => { in_tag = true; tag_buf.clear(); }
            '>' => {
                in_tag = false;
                // 태그 이름 추출 (선행 '/' 제거 후 소문자화)
                let tag_name = tag_buf
                    .trim_start_matches('/')
                    .split(|c: char| c.is_whitespace())
                    .next()
                    .unwrap_or("")
                    .to_lowercase();
                // 블록 요소에만 개행 삽입
                if matches!(tag_name.as_str(), "p" | "br" | "tr" | "td" | "th" | "div" | "li") {
                    out.push('\n');
                }
                // 인라인 요소(<a>, <span>, <b>, <i> 등) → 공백 없음
            }
            _ if in_tag => { tag_buf.push(ch); }
            _ => { out.push(ch); }
        }
    }
    decode_html_entities(&out)
}

