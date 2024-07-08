use std::{thread, time::{Duration, Instant}};
use clipboard_win::Clipboard;
use calamine::{open_workbook, DataType, Xlsx};
use std::fs::OpenOptions;


fn main() {
    // Excelファイルのパスとシート名を指定
    let excel_path = "example.xlsx";
    let sheet_name = "Sheet1";
    let target_cell = "A1";

    // 前回のクリップボード内容を保持する変数
    let mut previous_clipboard_content = String::new();

    // プログラムの開始時間を記録
    let start_time = Instant::now();
    let duration = Duration::from_secs(60); // 1分

    // クリップボードを定期的にチェック
    loop {
        // 経過時間をチェック
        if start_time.elapsed() >= duration {
            println!("Exiting program after 1 minute.");
            break;
        }

        // クリップボードの内容を取得
        if let Ok(mut clipboard) = Clipboard::new() {
            if let Ok(&current_clipboard_content) = clipboard.get_string() {
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
        let cell_idx = cell_to_index(cell)?;

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

    Ok(())
}

fn cell_to_index(cell: &str) -> Result<(usize, usize), Box<dyn std::error::Error>> {
    // セル位置を行と列に変換する
    let mut chars = cell.chars();
    let col = chars.by_ref().take_while(|c| c.is_alphabetic()).collect::<String>();
    let row = chars.collect::<String>();

    let col_num = col.chars().fold(0, |acc, c| acc * 26 + (c as usize - 'A' as usize + 1)) - 1;
    let row_num = row.parse::<usize>()? - 1;

    Ok((row_num, col_num))
}
