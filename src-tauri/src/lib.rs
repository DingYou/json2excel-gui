mod export;
mod output;
mod request;

use tauri_plugin_dialog::{DialogExt, MessageDialogButtons, MessageDialogKind};
use tauri_plugin_opener::OpenerExt;

#[tauri::command]
async fn get_json_fields(input: String) -> Result<Vec<export::ArrayFieldInfo>, String> {
    let (content, _) = output::read_source(&output::expand_tilde(&input))?;
    export::fields(&content)
}

#[tauri::command]
async fn get_json_fields_from_content(
    content: String,
) -> Result<Vec<export::ArrayFieldInfo>, String> {
    export::fields(&content)
}

fn confirm_replace(app: &tauri::AppHandle, path: &std::path::Path) -> bool {
    app.dialog().message(format!("输出文件已存在，是否替换？原文件将被覆盖。\nThe output file already exists. Replace it?\n\n{}", path.display()))
        .title("确认覆盖 / Confirm replacement")
        .kind(MessageDialogKind::Warning)
        .buttons(MessageDialogButtons::OkCancel)
        .blocking_show()
}

#[tauri::command]
async fn export_to_excel(
    app: tauri::AppHandle,
    input: String,
    output: String,
    selected_paths: Vec<String>,
) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let (content, source) = output::read_source(&output::expand_tilde(&input))?;
        output::export(
            &content,
            &output::expand_tilde(&output),
            &selected_paths,
            Some(&source),
            |path| confirm_replace(&app, path),
        )
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn export_to_excel_from_content(
    app: tauri::AppHandle,
    content: String,
    output: String,
    selected_paths: Vec<String>,
) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || {
        output::export(
            &content,
            &output::expand_tilde(&output),
            &selected_paths,
            None,
            |path| confirm_replace(&app, path),
        )
    })
    .await
    .map_err(|e| e.to_string())?
}

#[tauri::command]
async fn execute_curl(command: String) -> Result<String, String> {
    request::execute(&command).await
}

#[tauri::command]
fn reveal_in_explorer(app: tauri::AppHandle, path: String) -> Result<(), String> {
    app.opener()
        .reveal_item_in_dir(output::expand_tilde(&path))
        .map_err(|e| e.to_string())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_window_state::Builder::default().build())
        .invoke_handler(tauri::generate_handler![
            get_json_fields,
            get_json_fields_from_content,
            export_to_excel,
            export_to_excel_from_content,
            execute_curl,
            reveal_in_explorer
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
