#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]
mod files;
fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![files::read_binary_file, files::write_binary_file_atomic])
        .run(tauri::generate_context!())
        .expect("Could not start Sheet Cloud");
}
