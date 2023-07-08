mod read;
mod dialog;

use calamine::{Reader, Xlsx, open_workbook};

fn main() {

    dialog::alert("Select a .docx file to split");
    let path = dialog::get_path("Documents", "docx");
    

    dialog::alert("Choose an .xlsx file with keywords to split with");
    let keyword_path = dialog::get_path("Spreadsheets", "xlsx");
    let mut keywords: Vec<String> = vec![];
    let mut excel: Xlsx<_> = open_workbook(keyword_path).unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("Sheet1") {
        for row in r.rows() {
            for element in row.iter() {
                let mut string = element.to_string().to_ascii_lowercase();
                string.retain(|c| !c.is_whitespace());
                if !string.is_empty() {
                    keywords.push(string);
                }
            }
        }
    }

    println!("{:?}", keywords);

    dialog::alert("Select an output folder");

    let folder_path = dialog::get_dir_path();
    let folder_string = (folder_path).clone();

    read::run(path, folder_path, keywords);
    dialog::alert((format!("Splitting Completed, files saved to {}", folder_string.to_string_lossy())).as_str());
}


