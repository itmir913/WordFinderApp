const EMBEDDED_KEYWORDS: &str = include_str!("resources/embedded_keywords.csv");

#[derive(serde::Serialize)]
struct CsvLoadResult {
    // 바이트 그대로 넘긴다. read_to_string은 UTF-8만 받으므로, 한국 윈도우
    // Excel이 저장한 CP949 default.csv를 "파일이 없는 것"처럼 취급해
    // 조용히 내장 목록으로 넘어가 버린다. 디코딩은 프런트엔드에 한 곳으로 모았다.
    content: Vec<u8>,
    source: String, // "default_file" | "embedded"
}

#[tauri::command]
fn load_default_csv() -> CsvLoadResult {
    // 릴리즈(포터블): exe 옆에 있는 default.csv 우선 시도
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            let path = dir.join("default.csv");
            if let Ok(content) = std::fs::read(path) {
                return CsvLoadResult {
                    content,
                    source: "default_file".into(),
                };
            }
        }
    }
    // 개발 모드: 프로젝트 루트의 default.csv
    if let Ok(content) = std::fs::read("default.csv") {
        return CsvLoadResult {
            content,
            source: "default_file".into(),
        };
    }
    // Fallback: 프로그램 내장 단어 목록
    CsvLoadResult {
        content: EMBEDDED_KEYWORDS.as_bytes().to_vec(),
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
        .invoke_handler(tauri::generate_handler![
            load_default_csv,
            read_file_bytes,
            write_file_bytes,
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
