use std::fs;
use std::io::Write;
use std::path::Path;

#[tauri::command]
pub fn read_binary_file(path: String) -> Result<Vec<u8>, String> {
    let metadata = fs::metadata(&path).map_err(|e| e.to_string())?;
    if metadata.len() > 50 * 1024 * 1024 { return Err("File exceeds the 50 MB import limit".into()); }
    fs::read(path).map_err(|e| e.to_string())
}

#[tauri::command]
pub fn write_binary_file_atomic(path: String, data: Vec<u8>) -> Result<(), String> {
    let destination = Path::new(&path);
    let parent = destination.parent().filter(|p| !p.as_os_str().is_empty()).unwrap_or(Path::new("."));
    fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    let mut temporary = tempfile::NamedTempFile::new_in(parent).map_err(|e| e.to_string())?;
    temporary.write_all(&data).map_err(|e| e.to_string())?;
    temporary.as_file().sync_all().map_err(|e| e.to_string())?;
    temporary.persist(destination).map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn atomic_save_replaces_existing_workbook() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("workbook.scloud").to_str().unwrap().to_string();
        write_binary_file_atomic(path.clone(), b"first".to_vec()).unwrap();
        write_binary_file_atomic(path.clone(), b"second".to_vec()).unwrap();
        assert_eq!(read_binary_file(path).unwrap(), b"second");
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 1);
    }
    #[test]
    fn saves_binary_data_in_nested_directory() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("nested/data.xlsx").to_str().unwrap().to_string();
        let bytes = vec![0, 255, 123, 42];
        write_binary_file_atomic(path.clone(), bytes.clone()).unwrap();
        assert_eq!(read_binary_file(path).unwrap(), bytes);
    }
}
