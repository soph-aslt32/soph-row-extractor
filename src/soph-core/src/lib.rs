use std::collections::BTreeSet;
use std::path::Path;

// ──────────────────────────────────────────────
// セル参照パーサー
// ──────────────────────────────────────────────

/// セル参照文字列（例: "B3"）を (列番号, 行番号) の 1-indexed タプルに変換する
pub fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
    let s = s.trim().to_uppercase();
    let col_end = s.chars().take_while(|c| c.is_alphabetic()).count();
    if col_end == 0 || col_end == s.len() {
        return None;
    }
    let col = col_letters_to_num(&s[..col_end])?;
    let row: u32 = s[col_end..].parse().ok().filter(|&r: &u32| r > 0)?;
    Some((col, row))
}

fn col_letters_to_num(s: &str) -> Option<u32> {
    if s.is_empty() {
        return None;
    }
    let mut num: u32 = 0;
    for c in s.chars() {
        num = num
            .checked_mul(26)?
            .checked_add(c as u32 - b'A' as u32 + 1)?;
    }
    (num > 0).then_some(num)
}

// ──────────────────────────────────────────────
// 行抽出・結合
// ──────────────────────────────────────────────

pub struct ExtractionConfig {
    pub input_path: String,
    pub output_path: String,
    pub search_string: String,
    /// (列番号, 行番号) 1-indexed
    pub search_tl: (u32, u32),
    pub search_br: (u32, u32),
    pub prot_top: u32,
    pub prot_bottom: u32,
}

/// Excelファイルから条件に合う行を抽出・結合して output_path に保存する。
/// 成功時は抽出された行数（保護範囲を除く）を返す。
pub fn extract_and_combine(config: &ExtractionConfig) -> Result<usize, String> {
    let mut book = umya_spreadsheet::reader::xlsx::read(Path::new(&config.input_path))
        .map_err(|e| format!("ファイル読み込みエラー: {e}"))?;

    // 検索範囲内で文字列と完全一致するセルが存在する行番号を収集
    let search_str = config.search_string.as_str();
    let matched_rows: BTreeSet<u32> = {
        let sheet = book
            .get_sheet(&0)
            .ok_or("ワークシートが見つかりません")?;

        (config.search_tl.1..=config.search_br.1)
            .flat_map(|row| {
                (config.search_tl.0..=config.search_br.0).filter_map(move |col| {
                    sheet
                        .get_cell((col, row))
                        .filter(|c| c.get_value() == search_str)
                        .map(|_| row)
                })
            })
            .collect()
    };
    let matched_count = matched_rows.len();

    // 残す行番号セット = 保護範囲 ∪ 一致行
    let keep_rows: BTreeSet<u32> = (config.prot_top..=config.prot_bottom)
        .chain(matched_rows)
        .collect();

    // 最大行番号を取得
    let max_row = book
        .get_sheet(&0)
        .ok_or("ワークシートが見つかりません")?
        .get_highest_row();

    // 不要行を下から削除（削除による行インデックスのずれを防ぐ）
    let sheet = book
        .get_sheet_mut(&0)
        .ok_or("ワークシートが見つかりません")?;
    for row in (1..=max_row).rev() {
        if !keep_rows.contains(&row) {
            sheet.remove_row(&row, &1);
        }
    }

    umya_spreadsheet::writer::xlsx::write(&book, Path::new(&config.output_path))
        .map_err(|e| format!("ファイル書き込みエラー: {e}"))?;

    Ok(matched_count)
}

// ──────────────────────────────────────────────
// テスト
// ──────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_basic() {
        assert_eq!(parse_cell_ref("B3"), Some((2, 3)));
        assert_eq!(parse_cell_ref("Q11"), Some((17, 11)));
        assert_eq!(parse_cell_ref("A1"), Some((1, 1)));
        assert_eq!(parse_cell_ref("AA1"), Some((27, 1)));
    }

    #[test]
    fn parse_invalid() {
        assert_eq!(parse_cell_ref(""), None);
        assert_eq!(parse_cell_ref("3B"), None);
        assert_eq!(parse_cell_ref("B0"), None);
        assert_eq!(parse_cell_ref("B"), None);
    }
}
