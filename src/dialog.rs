use native_dialog::{FileDialog, MessageDialog};
use std::path::PathBuf;

const WINDOWSTITLE: &str = "DOCX splitter";

pub fn alert(message: &str) {
    let _no= MessageDialog::new().set_title(&WINDOWSTITLE).set_text(&message).show_alert();
}

pub fn get_path(filter_name: &str, filter_ext: &str) -> PathBuf {
    return FileDialog::new()
        .set_location("~/Desktop")
        .add_filter(filter_name, &[filter_ext])
        .show_open_single_file()
        .unwrap()
        .unwrap();
}

pub fn get_dir_path() -> PathBuf {
    return FileDialog::new()
        .set_location("~/Desktop")
        .show_open_single_dir()
        .unwrap()
        .unwrap();
}