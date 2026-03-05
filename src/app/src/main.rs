use eframe::egui;
use std::path::PathBuf;
use std::sync::Arc;

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size([520.0, 300.0]),
        ..Default::default()
    };
    eframe::run_native(
        "Excel 行抽出ツール",
        options,
        Box::new(|cc| {
            setup_fonts(&cc.egui_ctx);
            Ok(Box::new(MyApp::default()))
        }),
    )
}

fn setup_fonts(ctx: &egui::Context) {
    let mut fonts = egui::FontDefinitions::default();

    let font_paths = [
        "C:/Windows/Fonts/YuGothR.ttc",
        "C:/Windows/Fonts/meiryo.ttc",
        "C:/Windows/Fonts/msgothic.ttc",
    ];
    for path in font_paths {
        if let Ok(data) = std::fs::read(path) {
            fonts
                .font_data
                .insert("ja".to_owned(), Arc::new(egui::FontData::from_owned(data)));
            fonts
                .families
                .entry(egui::FontFamily::Proportional)
                .or_default()
                .insert(0, "ja".to_owned());
            fonts
                .families
                .entry(egui::FontFamily::Monospace)
                .or_default()
                .push("ja".to_owned());
            break;
        }
    }

    ctx.set_fonts(fonts);
}

#[derive(Default)]
struct MyApp {
    file_path: Option<PathBuf>,
    search_string: String,
    search_tl: String,
    search_br: String,
    prot_top: String,
    prot_bottom: String,
    status: String,
}

impl eframe::App for MyApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Excel 行抽出ツール");
            ui.separator();
            ui.add_space(6.0);

            // ── ファイル選択 ──
            ui.horizontal(|ui| {
                if ui.button("📂 Excelファイルを選択").clicked() {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("Excel", &["xlsx"])
                        .pick_file()
                    {
                        self.file_path = Some(path);
                        self.status.clear();
                    }
                }
                let label = self
                    .file_path
                    .as_ref()
                    .and_then(|p| p.file_name())
                    .map(|n| n.to_string_lossy().into_owned())
                    .unwrap_or_else(|| "未選択".to_string());
                ui.label(label);
            });

            ui.add_space(8.0);

            // ── 検索文字列 ──
            egui::Grid::new("inputs")
                .num_columns(2)
                .spacing([8.0, 6.0])
                .show(ui, |ui| {
                    ui.label("検索文字列");
                    ui.add(
                        egui::TextEdit::singleline(&mut self.search_string)
                            .desired_width(260.0)
                            .hint_text("完全一致する文字列"),
                    );
                    ui.end_row();

                    ui.label("検索範囲");
                    ui.horizontal(|ui| {
                        ui.add(
                            egui::TextEdit::singleline(&mut self.search_tl)
                                .desired_width(60.0)
                                .hint_text("B3"),
                        );
                        ui.label("～");
                        ui.add(
                            egui::TextEdit::singleline(&mut self.search_br)
                                .desired_width(60.0)
                                .hint_text("Q11"),
                        );
                    });
                    ui.end_row();

                    ui.label("保護範囲（行）");
                    ui.horizontal(|ui| {
                        ui.add(
                            egui::TextEdit::singleline(&mut self.prot_top)
                                .desired_width(50.0)
                                .hint_text("1"),
                        );
                        ui.label("～");
                        ui.add(
                            egui::TextEdit::singleline(&mut self.prot_bottom)
                                .desired_width(50.0)
                                .hint_text("3"),
                        );
                    });
                    ui.end_row();
                });

            ui.add_space(10.0);
            ui.separator();
            ui.add_space(8.0);

            // ── 実行ボタン ──
            let can_run = self.file_path.is_some()
                && !self.search_string.is_empty()
                && !self.search_tl.is_empty()
                && !self.search_br.is_empty()
                && !self.prot_top.is_empty()
                && !self.prot_bottom.is_empty();

            if ui
                .add_enabled(can_run, egui::Button::new("  実行  "))
                .clicked()
            {
                self.run();
            }

            // ── ステータス ──
            if !self.status.is_empty() {
                ui.add_space(8.0);
                ui.label(&self.status);
            }
        });
    }
}

impl MyApp {
    fn run(&mut self) {
        // 入力検証
        let tl = match soph_core::parse_cell_ref(&self.search_tl) {
            Some(v) => v,
            None => {
                self.status = format!(
                    "エラー: 検索範囲の左上 \"{}\" が無効な形式です（例: B3）",
                    self.search_tl
                );
                return;
            }
        };
        let br = match soph_core::parse_cell_ref(&self.search_br) {
            Some(v) => v,
            None => {
                self.status = format!(
                    "エラー: 検索範囲の右下 \"{}\" が無効な形式です（例: Q11）",
                    self.search_br
                );
                return;
            }
        };
        let prot_top: u32 = match self.prot_top.trim().parse().ok().filter(|&r: &u32| r > 0) {
            Some(v) => v,
            None => {
                self.status = "エラー: 保護範囲の開始行が無効です".to_string();
                return;
            }
        };
        let prot_bottom: u32 =
            match self.prot_bottom.trim().parse().ok().filter(|&r: &u32| r > 0) {
                Some(v) => v,
                None => {
                    self.status = "エラー: 保護範囲の終了行が無効です".to_string();
                    return;
                }
            };

        if prot_top > prot_bottom {
            self.status = "エラー: 保護範囲の開始行が終了行より大きいです".to_string();
            return;
        }
        if tl.0 > br.0 || tl.1 > br.1 {
            self.status = "エラー: 検索範囲の左上が右下より後ろになっています".to_string();
            return;
        }

        // 保存先ダイアログ
        let input_path = self.file_path.as_ref().unwrap();
        let default_name = input_path
            .file_stem()
            .map(|s| format!("{}_output.xlsx", s.to_string_lossy()))
            .unwrap_or_else(|| "output.xlsx".to_string());

        let output_path = match rfd::FileDialog::new()
            .add_filter("Excel", &["xlsx"])
            .set_file_name(&default_name)
            .save_file()
        {
            Some(p) => p,
            None => return, // キャンセル
        };

        let config = soph_core::ExtractionConfig {
            input_path: input_path.to_string_lossy().into_owned(),
            output_path: output_path.to_string_lossy().into_owned(),
            search_string: self.search_string.clone(),
            search_tl: tl,
            search_br: br,
            prot_top,
            prot_bottom,
        };

        match soph_core::extract_and_combine(&config) {
            Ok(count) => {
                self.status = format!(
                    "✓ 完了: {} 行が抽出されました。\n保存先: {}",
                    count,
                    output_path.display()
                );
            }
            Err(e) => {
                self.status = format!("エラー: {e}");
            }
        }
    }
}
