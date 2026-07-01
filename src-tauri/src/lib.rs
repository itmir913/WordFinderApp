const EMBEDDED_KEYWORDS: &str = include_str!("resources/embedded_keywords.csv");

#[derive(serde::Serialize)]
struct CsvLoadResult {
    content: String,
    source: String, // "default_file" | "embedded"
}

#[tauri::command]
fn load_default_csv() -> CsvLoadResult {
    // 릴리즈(포터블): exe 옆에 있는 default.csv 우선 시도
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            let path = dir.join("default.csv");
            if let Ok(content) = std::fs::read_to_string(path) {
                return CsvLoadResult { content, source: "default_file".into() };
            }
        }
    }
    // 개발 모드: 프로젝트 루트의 default.csv
    if let Ok(content) = std::fs::read_to_string("default.csv") {
        return CsvLoadResult { content, source: "default_file".into() };
    }
    // Fallback: 프로그램 내장 단어 목록
    CsvLoadResult {
        content: EMBEDDED_KEYWORDS.into(),
        source: "embedded".into(),
    }
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
