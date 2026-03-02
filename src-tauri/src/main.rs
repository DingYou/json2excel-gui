// Tauri 应用入口
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    json2excel_gui_lib::run()
}
