#[tauri::command]
fn load_default_csv() -> Result<String, String> {
    // 릴리즈(포터블): exe 옆에 있는 default.csv 우선 시도
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            let path = dir.join("default.csv");
            if path.exists() {
                return std::fs::read_to_string(path).map_err(|e| e.to_string());
            }
        }
    }
    // 개발 모드: 프로젝트 루트의 default.csv
    std::fs::read_to_string("default.csv").map_err(|e| e.to_string())
}

#[tauri::command]
fn read_file_bytes(path: String) -> Result<Vec<u8>, String> {
    std::fs::read(&path).map_err(|e| e.to_string())
}

#[tauri::command]
fn write_file_bytes(path: String, data: Vec<u8>) -> Result<(), String> {
    std::fs::write(&path, data).map_err(|e| e.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            load_default_csv,
            read_file_bytes,
            write_file_bytes,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
