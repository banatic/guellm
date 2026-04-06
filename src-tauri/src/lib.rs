mod commands;
mod com_dispatch;
mod config;
mod hwp_controller;
mod llm_client;
mod tools;

use config::{load_config, AppConfig};
use hwp_controller::{spawn_hwp_thread, HwpRequest};
use std::sync::{Arc, Mutex};

pub struct AppState {
    pub hwp_tx: tokio::sync::mpsc::Sender<HwpRequest>,
    pub config: Mutex<AppConfig>,
    pub cancel_token: Mutex<Option<tokio_util::sync::CancellationToken>>,
    pub pending_confirm: Arc<Mutex<Option<tokio::sync::oneshot::Sender<bool>>>>,
    pub last_backup_path: Mutex<Option<String>>,
}

pub fn run() {
    let (hwp_tx, hwp_rx) = tokio::sync::mpsc::channel::<HwpRequest>(64);

    // 전용 COM STA 스레드 생성
    spawn_hwp_thread(hwp_rx);

    let cfg = load_config();
    let state = AppState {
        hwp_tx,
        config: Mutex::new(cfg),
        cancel_token: Mutex::new(None),
        pending_confirm: Arc::new(Mutex::new(None)),
        last_backup_path: Mutex::new(None),
    };

    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .manage(state)
        .invoke_handler(tauri::generate_handler![
            commands::connect_hwp,
            commands::open_file_in_hwp,
            commands::is_connected,
            commands::save_document,
            commands::preview_structure,
            commands::run_agent,
            commands::cancel_agent,
            commands::confirm_tool,
            commands::rollback_agent,
            commands::get_config,
            commands::update_config,
            commands::open_file_dialog,
            commands::call_tool,
        ])
        .run(tauri::generate_context!())
        .expect("Tauri 앱 실행 오류");
}
