/// commands.rs
/// Tauri 명령 핸들러 — 프론트엔드 ↔ 러스트 백엔드 브릿지

use crate::{
    config::{save_config, AppConfig},
    hwp_controller::{HwpCommand, HwpRequest},
    llm_client::{AgentEvent, LlmClient},
    tools::WRITE_TOOLS,
    AppState,
};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use tauri::Emitter;

async fn send_hwp(
    state: &tauri::State<'_, AppState>,
    command: HwpCommand,
) -> anyhow::Result<String> {
    let (tx, rx) = tokio::sync::oneshot::channel();
    state
        .hwp_tx
        .send(HwpRequest { command, reply: tx })
        .await
        .map_err(|e| anyhow::anyhow!("HWP 스레드 전송 실패: {e}"))?;
    rx.await
        .map_err(|e| anyhow::anyhow!("HWP 결과 수신 실패: {e}"))?
}

// ── HWP 연결 ──────────────────────────────────────────────

#[tauri::command]
pub async fn connect_hwp(
    state: tauri::State<'_, AppState>,
    visible: bool,
) -> Result<(), String> {
    send_hwp(&state, HwpCommand::Connect { visible })
        .await
        .map(|_| ())
        .map_err(|e| e.to_string())
}

#[tauri::command]
pub async fn open_file_in_hwp(
    state: tauri::State<'_, AppState>,
    path: String,
) -> Result<(), String> {
    send_hwp(&state, HwpCommand::OpenFile { path })
        .await
        .map(|_| ())
        .map_err(|e| e.to_string())
}

#[tauri::command]
pub async fn is_connected(
    state: tauri::State<'_, AppState>,
) -> Result<bool, String> {
    let result = send_hwp(&state, HwpCommand::IsConnected)
        .await
        .map_err(|e| e.to_string())?;
    Ok(result == "true")
}

#[tauri::command]
pub async fn save_document(
    state: tauri::State<'_, AppState>,
    save_path: Option<String>,
) -> Result<String, String> {
    send_hwp(&state, HwpCommand::Save { save_path })
        .await
        .map_err(|e| e.to_string())
}

#[tauri::command]
pub async fn preview_structure(
    state: tauri::State<'_, AppState>,
) -> Result<String, String> {
    send_hwp(
        &state,
        HwpCommand::DispatchTool {
            name: "analyze_document_structure".to_string(),
            args: serde_json::json!({}),
        },
    )
    .await
    .map_err(|e| e.to_string())
}

// ── 에이전트 취소 ────────────────────────────────────────

#[tauri::command]
pub async fn cancel_agent(state: tauri::State<'_, AppState>) -> Result<(), String> {
    if let Some(token) = state.cancel_token.lock().unwrap().take() {
        token.cancel();
    }
    // Also unblock any pending confirmation
    let _ = state.pending_confirm.lock().unwrap().take();
    Ok(())
}

// ── 도구 실행 승인/거부 (Human-in-the-Loop) ──────────────

#[tauri::command]
pub async fn confirm_tool(
    state: tauri::State<'_, AppState>,
    approved: bool,
) -> Result<(), String> {
    if let Some(tx) = state.pending_confirm.lock().unwrap().take() {
        let _ = tx.send(approved);
    }
    Ok(())
}

// ── 문서 롤백 ────────────────────────────────────────────

#[tauri::command]
pub async fn rollback_agent(
    state: tauri::State<'_, AppState>,
) -> Result<(), String> {
    let path = state
        .last_backup_path
        .lock()
        .unwrap()
        .clone()
        .ok_or_else(|| "백업이 없습니다.".to_string())?;
    send_hwp(&state, HwpCommand::Rollback { backup_path: path })
        .await
        .map(|_| ())
        .map_err(|e| e.to_string())
}

// ── 에이전틱 실행 ─────────────────────────────────────────

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct HistoryMessage {
    pub role: String,
    pub content: String,
}

#[derive(Debug, Deserialize)]
pub struct RunAgentParams {
    pub provider: String,
    pub api_key: String,
    pub model: String,
    pub query: String,
    #[serde(default)]
    pub history: Vec<HistoryMessage>,
}

#[tauri::command]
pub async fn run_agent(
    app: tauri::AppHandle,
    state: tauri::State<'_, AppState>,
    params: RunAgentParams,
) -> Result<(), String> {
    // ── 1) 문서 요약 취득 ──
    let doc_summary = send_hwp(
        &state,
        HwpCommand::DispatchTool {
            name: "analyze_document_structure".to_string(),
            args: serde_json::json!({}),
        },
    )
    .await
    .map_err(|e| e.to_string())?;

    // ── 2) 문서 스냅샷 (Undo/Rollback 용) ──
    let backup_path = format!(
        "{}\\hwp_agent_backup_{}.hwp",
        std::env::temp_dir().display(),
        chrono_timestamp()
    );
    match send_hwp(
        &state,
        HwpCommand::Snapshot {
            backup_path: backup_path.clone(),
        },
    )
    .await
    {
        Ok(_) => {
            *state.last_backup_path.lock().unwrap() = Some(backup_path);
        }
        Err(e) => {
            // Snapshot failure is non-fatal — warn but continue
            eprintln!("스냅샷 저장 실패 (계속 진행): {e}");
        }
    }

    // ── 3) CancellationToken 생성 ──
    let cancel_token = tokio_util::sync::CancellationToken::new();
    *state.cancel_token.lock().unwrap() = Some(cancel_token.clone());

    let app_clone = app.clone();
    let hwp_tx = state.hwp_tx.clone();
    let pending_confirm = state.pending_confirm.clone();
    let cancel_for_cleanup = cancel_token.clone();

    // ── 4) 에이전틱 루프를 별도 tokio 태스크로 실행 ──
    tokio::spawn(async move {
        let client = match LlmClient::new(&params.provider, &params.api_key, Some(&params.model)) {
            Ok(c) => c,
            Err(e) => {
                let _ = app_clone.emit(
                    "agent-event",
                    AgentEvent::Error {
                        message: e.to_string(),
                    },
                );
                return;
            }
        };

        let app_for_cb = app_clone.clone();
        let event_cb: crate::llm_client::EventCallback =
            std::sync::Arc::new(move |event: AgentEvent| {
                let _ = app_for_cb.emit("agent-event", event);
            });

        // ── Tool executor with Human-in-the-Loop ──
        let hwp_tx_for_tool = hwp_tx.clone();
        let pending_confirm_for_tool = pending_confirm.clone();
        let app_for_tool = app_clone.clone();
        let cancel_for_tool = cancel_token.clone();

        let tool_executor: crate::llm_client::ToolExecutor =
            std::sync::Arc::new(move |name: String, args: Value| {
                let tx = hwp_tx_for_tool.clone();
                let pc = pending_confirm_for_tool.clone();
                let app_h = app_for_tool.clone();
                let ct = cancel_for_tool.clone();
                Box::pin(async move {
                    // ── Human-in-the-Loop: write tool인 경우 사용자 확인 대기 ──
                    if WRITE_TOOLS.contains(&name.as_str()) {
                        let (confirm_tx, confirm_rx) = tokio::sync::oneshot::channel::<bool>();
                        {
                            *pc.lock().unwrap() = Some(confirm_tx);
                        }
                        // 프론트엔드에 확인 요청 이벤트 전송
                        let _ = app_h.emit(
                            "agent-event",
                            AgentEvent::ToolConfirmRequest {
                                name: name.clone(),
                                args: args.clone(),
                            },
                        );

                        // 사용자 응답 또는 취소 대기
                        let approved = tokio::select! {
                            result = confirm_rx => {
                                result.unwrap_or(false)
                            }
                            _ = ct.cancelled() => {
                                return format!("⛔ 에이전트가 취소되었습니다.");
                            }
                        };

                        if !approved {
                            return format!("⛔ 사용자가 도구 실행을 거부했습니다: {name}");
                        }
                    }

                    // ── 실제 HWP 도구 실행 ──
                    let (reply_tx, reply_rx) = tokio::sync::oneshot::channel();
                    if tx
                        .send(HwpRequest {
                            command: HwpCommand::DispatchTool {
                                name: name.clone(),
                                args,
                            },
                            reply: reply_tx,
                        })
                        .await
                        .is_err()
                    {
                        return format!("❌ HWP 채널 전송 실패: {name}");
                    }
                    reply_rx
                        .await
                        .unwrap_or_else(|_| Ok("❌ 응답 수신 실패".to_string()))
                        .unwrap_or_else(|e| format!("❌ {name} 오류: {e}"))
                })
            });

        match client
            .call_agentic(
                &doc_summary,
                &params.query,
                &params.history,
                tool_executor,
                event_cb,
                25,
                cancel_for_cleanup,
            )
            .await
        {
            Ok(response) => {
                let _ = app_clone.emit(
                    "agent-event",
                    AgentEvent::FinalResponse { text: response },
                );
            }
            Err(e) => {
                let _ = app_clone.emit(
                    "agent-event",
                    AgentEvent::Error {
                        message: e.to_string(),
                    },
                );
            }
        }
    });

    Ok(())
}

/// Simple timestamp for backup filenames
fn chrono_timestamp() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let d = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();
    format!("{}", d.as_secs())
}

// ── 설정 관리 ─────────────────────────────────────────────

#[tauri::command]
pub async fn get_config(state: tauri::State<'_, AppState>) -> Result<AppConfig, String> {
    let cfg = state.config.lock().unwrap().clone();
    Ok(cfg)
}

#[tauri::command]
pub async fn update_config(
    state: tauri::State<'_, AppState>,
    new_config: AppConfig,
) -> Result<(), String> {
    {
        let mut cfg = state.config.lock().unwrap();
        *cfg = new_config.clone();
    }
    save_config(&new_config).map_err(|e| e.to_string())
}

// ── 파일 다이얼로그 ───────────────────────────────────────

#[tauri::command]
pub async fn open_file_dialog(app: tauri::AppHandle) -> Result<Option<String>, String> {
    use tauri_plugin_dialog::DialogExt;
    let file = app
        .dialog()
        .file()
        .add_filter("한글 문서", &["hwp", "hwpx"])
        .add_filter("모든 파일", &["*"])
        .blocking_pick_file();

    Ok(file.map(|p| p.to_string()))
}
