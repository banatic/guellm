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
            let act = hwp
                .call("CreateAction", vec![Variant::String("FileSaveAs_S".to_string())])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CreateAction 실패"))?;
            let ps = act
                .call("CreateSet", vec![])?
                .as_object()
                .ok_or_else(|| anyhow::anyhow!("CreateSet 실패"))?;
            ps.call("SetItem", vec![
                Variant::String("filename".to_string()),
                Variant::String(path.to_string()),
            ])?;
            ps.call("SetItem", vec![
                Variant::String("Format".to_string()),
                Variant::String("HWP".to_string()),
            ])?;
            ps.call("SetItem", vec![
                Variant::String("Attributes".to_string()),
                Variant::I32(0),
            ])?;
            act.call("Execute", vec![Variant::Object(ps)])?;
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

    // ─────────────── 텍스트 스캔 (InitScan/GetText) ─────────

    /// InitScan(0x07) + GetText()로 문서 전체 텍스트를 단락 단위로 추출.
    /// 표 내부 텍스트도 포함됩니다.
    fn scan_all_text(&self) -> anyhow::Result<Vec<String>> {
        let hwp = self.hwp()?;
        // option 0x07 = 본문(0x01) + 각주(0x02) + 표(0x04)
        hwp.call("InitScan", vec![Variant::I32(0x07)])?;

        let mut paragraphs = Vec::new();
        let mut buf = Vec::new();

        for _ in 0..20_000 {
            let result = hwp.call("GetText", vec![])?;
            // GetText returns a tuple-like: (state, text)
            // In COM it may come as a string "state\x00text" or we get it from the return
            let repr = result.to_string_repr();
            // HWP GetText returns (state_code, text_string)
            // state: 0=끝, 1=일반텍스트, 2=단락끝, 3=섹션끝
            let (state, text) = parse_gettext_result(&result, &repr);

            if state == 0 {
                if !buf.is_empty() {
                    paragraphs.push(buf.join(""));
                }
                break;
            }
            if state == 1 || state == 2 || state == 3 {
                let tok = text.trim().to_string();
                if !tok.is_empty() {
                    buf.push(tok);
                }
            }
            if state == 2 || state == 3 {
                let para = buf.join(" ").trim().to_string();
                if !para.is_empty() {
                    paragraphs.push(para);
                }
                buf.clear();
            }
        }

        let _ = hwp.call("ReleaseScan", vec![]);
        Ok(paragraphs)
    }

    /// 문서에서 텍스트가 존재하는지 ForwardFind로 확인
    fn text_exists(&self, text: &str) -> bool {
        let hwp = match self.hwp() {
            Ok(h) => h,
            Err(_) => return false,
        };
        // 문서 처음으로 이동
        let _ = hwp.call("Run", vec![Variant::String("MoveDocBegin".to_string())]);
        let act = match hwp.call("CreateAction", vec![Variant::String("ForwardFind".to_string())]) {
            Ok(v) => match v.as_object() {
                Some(o) => o,
                None => return false,
            },
            Err(_) => return false,
        };
        let ps = match act.call("CreateSet", vec![]) {
            Ok(v) => match v.as_object() {
                Some(o) => o,
                None => return false,
            },
            Err(_) => return false,
        };
        let _ = ps.call("SetItem", vec![
            Variant::String("FindString".to_string()),
            Variant::String(text.to_string()),
        ]);
        let _ = ps.call("SetItem", vec![
            Variant::String("IgnoreCase".to_string()),
            Variant::Bool(true),
        ]);
        let _ = ps.call("SetItem", vec![
            Variant::String("Direction".to_string()),
            Variant::I32(3), // 전체 문서
        ]);
        act.call("Execute", vec![Variant::Object(ps)])
            .ok()
            .and_then(|v| v.as_bool())
            .unwrap_or(false)
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
        if !self.text_exists(find) {
            return Ok(());
        }
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
        ps.call("SetItem", vec![
            Variant::String("AllWordReplace".to_string()),
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

        // InitScan으로 단락 수 계산
        let paragraphs = self.scan_all_text().unwrap_or_default();
        result["paragraph_count"] = json!(paragraphs.len());

        // 표 수는 네비게이션으로 카운트
        let tables = self.scan_tables()?;
        result["table_count"] = json!(tables.len());

        let fields = self.get_field_names();
        result["fields"] = json!(fields);

        Ok(result)
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

    // ─────────────── 표 스캔 (테이블 네비게이션) ─────────────

    /// 문서의 모든 표를 순회하며 구조를 읽어옵니다.
    /// 각 표: { rows, cols, cells: [[text, ...], ...] }
    fn scan_tables(&self) -> anyhow::Result<Vec<TableData>> {
        let hwp = self.hwp()?;
        let mut tables = Vec::new();

        // 문서 처음으로 이동
        hwp.call("Run", vec![Variant::String("MoveDocBegin".to_string())])?;

        // 표를 순차적으로 찾아간다
        loop {
            // 현재 위치에서 다음 표를 찾음
            // ShapeObjTableSelAll: 현재 커서가 표 안이면 표 전체 선택
            // 커서가 표 밖이면 Ctrl 기반으로 다음 표로 이동해야 함
            //
            // 전략: MoveNextParaBegin을 반복하면서 KeyIndicator로 표 진입을 감지
            // 또는 직접 표에 진입: Ctrl + 표 진입 방식

            // 현재 위치가 표인지 확인
            let in_table = self.is_cursor_in_table()?;
            if in_table {
                // 이미 표 안에 있으면 읽기
                if let Ok(table) = self.read_current_table() {
                    tables.push(table);
                }
                // 표 밖으로 나가기
                hwp.call("Run", vec![Variant::String("MoveNextParaBegin".to_string())])?;
                hwp.call("Run", vec![Variant::String("MoveNextParaBegin".to_string())])?;
                continue;
            }

            // 다음 표를 찾기: Ctrl(표 객체)을 이용
            // HWP에서는 표 진입을 위해 표 위에서 Enter 또는 표를 클릭
            // 프로그래밍적으로는: 다음 컨트롤(표/이미지 등)로 이동 후 표인지 확인
            let moved = hwp
                .call("Run", vec![Variant::String("MoveNextParaBegin".to_string())])
                .ok()
                .and_then(|v| v.as_bool())
                .unwrap_or(false);

            if !moved {
                break; // 문서 끝
            }

            // 이동 후 표 안인지 확인
            if self.is_cursor_in_table()? {
                if let Ok(table) = self.read_current_table() {
                    tables.push(table);
                }
                // 표 밖으로
                hwp.call("Run", vec![Variant::String("MoveNextParaBegin".to_string())])?;
                hwp.call("Run", vec![Variant::String("MoveNextParaBegin".to_string())])?;
            }
        }

        // scan_tables가 비었으면 대체 방법: GetTextFile("UNICODE")에서 표 카운트
        // (이미 위에서 네비게이션으로 탐색했으므로 보통 여기까지 오지 않음)

        Ok(tables)
    }

    /// KeyIndicator로 현재 커서가 표 안에 있는지 확인
    fn is_cursor_in_table(&self) -> anyhow::Result<bool> {
        let hwp = self.hwp()?;
        // KeyIndicator 반환: (seccnt, secno, prgcnt, prgno, colcnt, colno, line, pos, over, ctrlname)
        // ctrlname이 "표"이면 표 안에 있음
        match hwp.call("KeyIndicator", vec![]) {
            Ok(result) => {
                let repr = result.to_string_repr();
                // KeyIndicator가 반환하는 형태에 따라 파싱
                // 일반적으로 컨트롤 이름이 포함됨
                Ok(repr.contains("표") || repr.contains("table") || repr.contains("Table"))
            }
            Err(_) => Ok(false),
        }
    }

    /// 현재 커서가 위치한 표의 전체 셀을 읽습니다.
    /// 전략: 표의 첫 셀로 이동 → 셀 읽기 → TableRightCell로 다음 셀 →
    ///        위치가 변하지 않으면 행 끝 → 다시 이동하면 다음 행 → 위치가 안 변하면 표 끝
    fn read_current_table(&self) -> anyhow::Result<TableData> {
        let hwp = self.hwp()?;

        // 표 첫 셀로 이동 (Ctrl+Home in table)
        hwp.call("Run", vec![Variant::String("TableColBegin".to_string())])?;
        hwp.call("Run", vec![Variant::String("TableCellBlock".to_string())])?;
        hwp.call("Run", vec![Variant::String("TableCellBlock".to_string())])?;
        hwp.call("Run", vec![Variant::String("MoveTopLevelBegin".to_string())])?;

        let mut rows: Vec<Vec<String>> = Vec::new();
        let mut current_row: Vec<String> = Vec::new();
        let mut prev_pos = self.get_cursor_pos();
        let mut total_cells_read = 0usize;
        let max_cells = 5000; // 안전 제한

        loop {
            if total_cells_read >= max_cells {
                break;
            }

            // 현재 셀 텍스트 읽기
            let cell_text = self.read_current_cell_text()?;
            current_row.push(cell_text);
            total_cells_read += 1;

            // 다음 셀로 이동
            hwp.call("Run", vec![Variant::String("TableRightCell".to_string())])?;
            let new_pos = self.get_cursor_pos();

            // 위치가 안 변했으면 표 끝
            if new_pos == prev_pos {
                if !current_row.is_empty() {
                    rows.push(std::mem::take(&mut current_row));
                }
                break;
            }

            // 행이 바뀌었는지 확인 (row 번호 비교)
            if new_pos.1 != prev_pos.1 {
                rows.push(std::mem::take(&mut current_row));
            }

            prev_pos = new_pos;
        }

        let row_count = rows.len();
        let col_count = rows.iter().map(|r| r.len()).max().unwrap_or(0);

        Ok(TableData {
            rows: row_count,
            cols: col_count,
            cells: rows,
        })
    }

    /// 현재 셀의 텍스트를 읽습니다.
    fn read_current_cell_text(&self) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        // 셀 전체 선택 후 선택 텍스트 가져오기
        hwp.call("Run", vec![Variant::String("SelectAll".to_string())])?;
        let text = hwp
            .call("GetSelectedText", vec![])
            .ok()
            .map(|v| v.to_string_repr())
            .unwrap_or_default();
        // 선택 해제
        hwp.call("Run", vec![Variant::String("Cancel".to_string())])?;
        Ok(text.trim().to_string())
    }

    /// 커서 위치를 (list, para, pos) 튜플로 반환
    fn get_cursor_pos(&self) -> (i32, i32, i32) {
        let hwp = match self.hwp() {
            Ok(h) => h,
            Err(_) => return (-1, -1, -1),
        };
        // GetPos returns (list, para, pos)
        match hwp.call("GetPos", vec![]) {
            Ok(v) => {
                let repr = v.to_string_repr();
                // 반환 형태에 따라 파싱: "(list, para, pos)" 또는 개별 값
                parse_pos_tuple(&repr)
            }
            Err(_) => (-1, -1, -1),
        }
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
                    .map(|h| {
                        if h.len() > 30 {
                            format!("{}...", &h[..30])
                        } else {
                            h.clone()
                        }
                    })
                    .collect();

                let is_title = table.rows <= 2
                    && table.cols <= 2
                    && table.cells.iter().all(|r| r.iter().all(|c| c.len() < 80));

                let mut entry = json!({
                    "table_index": idx,
                    "rows": table.rows,
                    "cols": table.cols,
                    "headers": headers,
                });

                if is_title {
                    entry["role"] = json!("title");
                }
                // 인접 표 관계
                if is_title && idx + 1 < tables.len() {
                    let next = &tables[idx + 1];
                    let next_headers: Vec<String> = next
                        .cells
                        .first()
                        .cloned()
                        .unwrap_or_default()
                        .iter()
                        .take(4)
                        .map(|h| {
                            if h.len() > 20 {
                                format!("{}...", &h[..20])
                            } else {
                                h.clone()
                            }
                        })
                        .collect();
                    entry["data_table"] = json!(format!(
                        "→ table_index:{} ({}x{}, {})",
                        idx + 1,
                        next.rows,
                        next.cols,
                        next_headers.join("/")
                    ));
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

        let cells: Vec<Vec<Value>> = table
            .cells
            .iter()
            .map(|row| row.iter().map(|c| json!(c)).collect())
            .collect();

        Ok(json!({
            "table_index": table_index,
            "rows": table.rows,
            "cols": table.cols,
            "headers": headers,
            "cells": cells,
        }))
    }

    fn find_text_anchor(&self, keyword: &str) -> anyhow::Result<Value> {
        if !self.text_exists(keyword) {
            return Ok(json!({"found": false, "keyword": keyword}));
        }
        let hwp = self.hwp()?;
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
            Variant::String(keyword.to_string()),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("IgnoreCase".to_string()),
            Variant::Bool(true),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("Direction".to_string()),
            Variant::I32(3),
        ])?;
        ps.call("SetItem", vec![
            Variant::String("FindReplace".to_string()),
            Variant::I32(0),
        ])?;
        let found = act
            .call("Execute", vec![Variant::Object(ps)])?
            .as_bool()
            .unwrap_or(false);
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
        let mut results = vec![];
        for (pattern, value) in map {
            let val_str = value.as_str().map(str::to_string).unwrap_or_else(|| value.to_string());
            match self.find_replace(pattern, &val_str, true) {
                Ok(_) => results.push(format!("✅ '{pattern}' → '{val_str}'")),
                Err(e) => results.push(format!("❌ '{pattern}': {e}")),
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
    fn navigate_to_table(&self, table_index: usize) -> anyhow::Result<()> {
        let hwp = self.hwp()?;
        hwp.call("Run", vec![Variant::String("MoveDocBegin".to_string())])?;

        let mut found_tables = 0usize;
        for _ in 0..50_000 {
            if self.is_cursor_in_table()? {
                if found_tables == table_index {
                    return Ok(());
                }
                // 이 표를 벗어나기
                loop {
                    let moved = hwp
                        .call("Run", vec![Variant::String("MoveNextParaBegin".to_string())])
                        .ok()
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false);
                    if !moved || !self.is_cursor_in_table()? {
                        break;
                    }
                }
                found_tables += 1;
                continue;
            }
            let moved = hwp
                .call("Run", vec![Variant::String("MoveNextParaBegin".to_string())])
                .ok()
                .and_then(|v| v.as_bool())
                .unwrap_or(false);
            if !moved {
                break;
            }
        }

        anyhow::bail!("표 {table_index}번을 찾을 수 없습니다 (발견: {found_tables}개)")
    }

    fn fill_table_data_matrix(
        &self,
        table_index: usize,
        start_row: usize,
        matrix: &Value,
        cell_delay_secs: f64,
    ) -> anyhow::Result<String> {
        let hwp = self.hwp()?;
        let schema = self.get_table_schema(table_index)?;
        let total_rows = schema["rows"].as_u64().unwrap_or(0) as usize;
        if total_rows == 0 {
            return Ok(format!("❌ 표 {table_index}번을 찾을 수 없습니다."));
        }
        let cols = schema["cols"].as_u64().unwrap_or(0) as usize;

        let data_rows = matrix
            .as_array()
            .ok_or_else(|| anyhow::anyhow!("matrix가 배열이어야 합니다"))?;

        // 해당 표로 이동
        self.navigate_to_table(table_index)?;

        // 표 첫 셀로 이동
        hwp.call("Run", vec![Variant::String("MoveTopLevelBegin".to_string())])?;

        // start_row행 첫 열까지 이동
        let tabs = start_row * cols;
        for _ in 0..tabs {
            hwp.call("Run", vec![Variant::String("TableRightCell".to_string())])?;
        }

        let mut filled = 0usize;
        let mut skipped = vec![];

        for (r_offset, row_data) in data_rows.iter().enumerate() {
            let row_arr = match row_data.as_array() {
                Some(a) => a,
                None => continue,
            };
            for (c_idx, cell_val) in row_arr.iter().enumerate() {
                if c_idx >= cols {
                    skipped.push(format!("행{} col{} 열 범위 초과", start_row + r_offset, c_idx));
                    continue;
                }
                let new_text = cell_val
                    .as_str()
                    .map(str::to_string)
                    .unwrap_or_else(|| cell_val.to_string());

                hwp.call("Run", vec![Variant::String("SelectAll".to_string())])?;

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
                    Variant::String(new_text),
                ])?;
                ins.call("Execute", vec![Variant::Object(ips)])?;
                filled += 1;

                if cell_delay_secs > 0.0 {
                    std::thread::sleep(Duration::from_secs_f64(cell_delay_secs));
                }

                if c_idx < row_arr.len() - 1 {
                    hwp.call("Run", vec![Variant::String("TableRightCell".to_string())])?;
                }
            }
            if r_offset < data_rows.len() - 1 {
                hwp.call("Run", vec![Variant::String("TableRightCell".to_string())])?;
            }
        }

        let mut msg = format!("✅ {filled}개 셀 채움");
        if !skipped.is_empty() {
            msg.push_str(&format!("\n⚠️ 건너뜀: {}", skipped[..skipped.len().min(5)].join(", ")));
        }
        Ok(msg)
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

        let paragraphs = self.scan_all_text().unwrap_or_default();
        if !paragraphs.is_empty() {
            let text = paragraphs.join("\n");
            let truncated = if text.chars().count() > 2000 {
                format!("{}...", text.chars().take(2000).collect::<String>())
            } else {
                text
            };
            lines.push(String::new());
            lines.push("=== 본문 텍스트 ===".to_string());
            lines.push(truncated);
        }

        lines.join("\n")
    }

    // ─────────────── Tool Dispatcher ────────────────────────

    pub fn dispatch_tool(&mut self, name: &str, args: &Value) -> anyhow::Result<String> {
        match name {
            "analyze_document_structure" => {
                Ok(self.analyze_document_structure()?.to_string())
            }
            "get_field_info" => Ok(self.get_field_info()?.to_string()),
            "get_all_tables_overview" => Ok(self.get_all_tables_overview()?.to_string()),
            "get_table_schema" => {
                let idx = args["table_index"].as_u64().unwrap_or(0) as usize;
                Ok(self.get_table_schema(idx)?.to_string())
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
                let mapping = if args.get("mapping").is_some() && !args["mapping"].is_null() {
                    &args["mapping"]
                } else {
                    args
                };
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
                let mtx = &args["matrix"];
                let delay = args["cell_delay"].as_f64().unwrap_or(0.0);
                self.fill_table_data_matrix(ti, sr, mtx, delay)
            }
            "format_table_cells" => {
                let ti = args["table_index"].as_u64().unwrap_or(0) as usize;
                let fmt = &args["format"];
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

// ──────────────────────────────────────────────────────────────
// 헬퍼 함수
// ──────────────────────────────────────────────────────────────

/// GetText 반환값에서 (state, text) 추출
/// HWP COM GetText()는 다양한 형태로 반환할 수 있음:
/// - 튜플 (i32, String)
/// - 또는 문자열 표현
fn parse_gettext_result(variant: &Variant, repr: &str) -> (i32, String) {
    // Variant가 직접 튜플이면
    if let Some(n) = variant.as_i32() {
        return (n, String::new());
    }
    // 문자열 표현에서 파싱 시도: "(1, 텍스트)" 형태
    let trimmed = repr.trim();
    if trimmed.starts_with('(') && trimmed.contains(',') {
        let inner = trimmed.trim_start_matches('(').trim_end_matches(')');
        if let Some((state_part, text_part)) = inner.split_once(',') {
            if let Ok(state) = state_part.trim().parse::<i32>() {
                return (state, text_part.trim().trim_matches('"').trim_matches('\'').to_string());
            }
        }
    }
    // 숫자만 있으면 state
    if let Ok(n) = trimmed.parse::<i32>() {
        return (n, String::new());
    }
    // 그 외: 텍스트로 간주
    (1, repr.to_string())
}

/// GetPos 반환값에서 (list, para, pos) 추출
fn parse_pos_tuple(repr: &str) -> (i32, i32, i32) {
    let trimmed = repr.trim().trim_start_matches('(').trim_end_matches(')');
    let parts: Vec<&str> = trimmed.split(',').collect();
    if parts.len() >= 3 {
        let a = parts[0].trim().parse().unwrap_or(-1);
        let b = parts[1].trim().parse().unwrap_or(-1);
        let c = parts[2].trim().parse().unwrap_or(-1);
        (a, b, c)
    } else if parts.len() == 1 {
        let a = parts[0].trim().parse().unwrap_or(-1);
        (a, -1, -1)
    } else {
        (-1, -1, -1)
    }
}
