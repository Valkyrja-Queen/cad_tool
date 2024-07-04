use std::{thread, time::Duration};
use clipboard_win::{formats, Clipboard, Getter};
use calamine::{open_workbook, DataType, Range, Reader, Xlsx};
use std::fs::OpenOptions;
use std::io::Write;
use std::path::Path;

fn main() {
    // Excelファイルのパスとシート名を指定
    let excel_path = "example.xlsx";
    let sheet_name = "Sheet1";
    let target_cell = "A1";

    // 前回のクリップボード内容を保持する変数
    let mut previous_clipboard_content = String::new();

    // クリップボードを定期的にチェック
    loop {
        // クリップボードの内容を取得
        if let Ok(mut clipboard) = Clipboard::new() {
            if let Ok(current_clipboard_content) = clipboard.get_string() {
                // クリップボードの内容が変わった場合
                if current_clipboard_content != previous_clipboard_content {
                    println!("Clipboard updated: {}", current_clipboard_content);

                    // Excelファイルを更新
                    if let Err(e) = update_excel(excel_path, sheet_name, target_cell, &current_clipboard_content) {
                        eprintln!("Failed to update Excel: {}", e);
                    }

                    // 前回のクリップボード内容を更新
                    previous_clipboard_content = current_clipboard_content;
                }
            }
        }

        // 1秒待つ
        thread::sleep(Duration::from_secs(1));
    }
}

fn update_excel(excel_path: &str, sheet_name: &str, cell: &str, content: &str) -> Result<(), Box<dyn std::error::Error>> {
    // Excelファイルを開く
    let mut workbook: Xlsx<_> = open_workbook(excel_path)?;

    // 指定したシートを取得
    if let Some(Ok(mut range)) = workbook.worksheet_range_mut(sheet_name) {
        // セルの位置を取得
        if let Some(cell_idx) = range.get_start().zip(range.get_end()).and_then(|(start, end)| {
            range.get_range(start, end).unwrap().get(cell)
        }) {
            let (row, col) = cell_idx;

            // セルの内容を更新
            range.set_value((row, col), DataType::String(content.to_string()));

            // 更新されたExcelファイルを保存
            let mut file = OpenOptions::new()
                .write(true)
                .truncate(true)
                .open(excel_path)?;
            workbook.write(&mut file)?;
        }
    }

    Ok(())
}

