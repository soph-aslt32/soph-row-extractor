#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use eframe::egui;
use std::path::PathBuf;
use std::sync::Arc;

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([520.0, 420.0])
            .with_min_inner_size([400.0, 360.0]),
        ..Default::default()
    };
    eframe::run_native(
        "Excel 行抽出ツール",
        options,
        Box::new(|cc| {
            setup_fonts(&cc.egui_ctx);
            setup_style(&cc.egui_ctx);
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

fn setup_style(ctx: &egui::Context) {
    let mut visuals = egui::Visuals::dark();
    visuals.panel_fill = egui::Color32::from_rgb(28, 32, 48);
    visuals.window_fill = egui::Color32::from_rgb(28, 32, 48);
    visuals.widgets.noninteractive.corner_radius = egui::CornerRadius::same(5);
    visuals.widgets.inactive.corner_radius = egui::CornerRadius::same(5);
    visuals.widgets.hovered.corner_radius = egui::CornerRadius::same(5);
    visuals.widgets.active.corner_radius = egui::CornerRadius::same(5);
    visuals.widgets.inactive.bg_fill = egui::Color32::from_rgb(44, 50, 74);
    visuals.widgets.inactive.weak_bg_fill = egui::Color32::from_rgb(44, 50, 74);
    visuals.widgets.hovered.bg_fill = egui::Color32::from_rgb(58, 65, 95);
    visuals.widgets.hovered.weak_bg_fill = egui::Color32::from_rgb(58, 65, 95);
    visuals.widgets.active.bg_fill = egui::Color32::from_rgb(70, 80, 115);
    visuals.selection.bg_fill = egui::Color32::from_rgb(82, 110, 200);
    ctx.set_visuals(visuals);

    let mut style = (*ctx.style()).clone();
    style.spacing.item_spacing = egui::vec2(8.0, 8.0);
    style.spacing.button_padding = egui::vec2(10.0, 6.0);
    ctx.set_style(style);
}

#[derive(Default)]
struct MyApp {
    file_path: Option<PathBuf>,
    file_path2: Option<PathBuf>,
    search_string: String,
    search_tl: String,
    search_br: String,
    prot_top: String,
    prot_bottom: String,
    status: String,
}

impl eframe::App for MyApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default()
            .frame(
                egui::Frame::new()
                    .fill(egui::Color32::from_rgb(28, 32, 48))
                    .inner_margin(egui::Margin::same(20)),
            )
            .show(ctx, |ui| {
                // ── ヘッダー ──
                ui.label(
                    egui::RichText::new("Excel 行抽出ツール")
                        .size(20.0)
                        .strong()
                        .color(egui::Color32::from_rgb(180, 200, 255)),
                );
                ui.label(
                    egui::RichText::new("指定した文字列に一致する行を抽出・結合します")
                        .size(11.0)
                        .color(egui::Color32::from_gray(140)),
                );
                ui.add_space(14.0);

                // ── ファイル選択カード ──
                egui::Frame::new()
                    .fill(egui::Color32::from_rgb(38, 42, 62))
                    .corner_radius(egui::CornerRadius::same(8))
                    .inner_margin(egui::Margin::same(14))
                    .show(ui, |ui| {
                        ui.label(
                            egui::RichText::new("INPUT FILES")
                                .size(10.0)
                                .strong()
                                .color(egui::Color32::from_rgb(130, 150, 200)),
                        );
                        ui.add_space(6.0);
                        // ファイル 1
                        ui.horizontal(|ui| {
                            if ui
                                .add(
                                    egui::Button::new(
                                        egui::RichText::new("\u{1f4c2}  ファイル 1")
                                            .color(egui::Color32::WHITE),
                                    )
                                    .fill(egui::Color32::from_rgb(50, 60, 95)),
                                )
                                .clicked()
                            {
                                if let Some(path) = rfd::FileDialog::new()
                                    .add_filter("Excel", &["xlsx"])
                                    .pick_file()
                                {
                                    self.file_path = Some(path);
                                    self.status.clear();
                                }
                            }
                            let (label, color) = match &self.file_path {
                                Some(p) => (
                                    p.file_name()
                                        .map(|n| n.to_string_lossy().into_owned())
                                        .unwrap_or_default(),
                                    egui::Color32::from_rgb(166, 227, 161),
                                ),
                                None => (
                                    "未選択".to_string(),
                                    egui::Color32::from_gray(120),
                                ),
                            };
                            ui.label(egui::RichText::new(label).color(color));
                        });
                        // ファイル 2
                        ui.horizontal(|ui| {
                            if ui
                                .add(
                                    egui::Button::new(
                                        egui::RichText::new("\u{1f4c2}  ファイル 2")
                                            .color(egui::Color32::WHITE),
                                    )
                                    .fill(egui::Color32::from_rgb(50, 60, 95)),
                                )
                                .clicked()
                            {
                                if let Some(path) = rfd::FileDialog::new()
                                    .add_filter("Excel", &["xlsx"])
                                    .pick_file()
                                {
                                    self.file_path2 = Some(path);
                                    self.status.clear();
                                }
                            }
                            let (label, color) = match &self.file_path2 {
                                Some(p) => (
                                    p.file_name()
                                        .map(|n| n.to_string_lossy().into_owned())
                                        .unwrap_or_default(),
                                    egui::Color32::from_rgb(166, 227, 161),
                                ),
                                None => (
                                    "未選択".to_string(),
                                    egui::Color32::from_gray(120),
                                ),
                            };
                            ui.label(egui::RichText::new(label).color(color));
                        });
                    });

                ui.add_space(8.0);

                // ── 検索設定カード ──
                egui::Frame::new()
                    .fill(egui::Color32::from_rgb(38, 42, 62))
                    .corner_radius(egui::CornerRadius::same(8))
                    .inner_margin(egui::Margin::same(14))
                    .show(ui, |ui| {
                        ui.label(
                            egui::RichText::new("検索設定")
                                .size(10.0)
                                .strong()
                                .color(egui::Color32::from_rgb(130, 150, 200)),
                        );
                        ui.add_space(8.0);
                        egui::Grid::new("inputs")
                            .num_columns(2)
                            .spacing([12.0, 8.0])
                            .show(ui, |ui| {
                                ui.label(
                                    egui::RichText::new("検索文字列")
                                        .color(egui::Color32::from_gray(180)),
                                );
                                ui.add(
                                    egui::TextEdit::singleline(&mut self.search_string)
                                        .desired_width(280.0)
                                        .hint_text("完全一致する文字列"),
                                );
                                ui.end_row();

                                ui.label(
                                    egui::RichText::new("検索範囲")
                                        .color(egui::Color32::from_gray(180)),
                                );
                                ui.horizontal(|ui| {
                                    ui.add(
                                        egui::TextEdit::singleline(&mut self.search_tl)
                                            .desired_width(55.0)
                                            .hint_text("B3"),
                                    );
                                    ui.label(
                                        egui::RichText::new("→")
                                            .color(egui::Color32::from_gray(140)),
                                    );
                                    ui.add(
                                        egui::TextEdit::singleline(&mut self.search_br)
                                            .desired_width(55.0)
                                            .hint_text("Q11"),
                                    );
                                });
                                ui.end_row();

                                ui.label(
                                    egui::RichText::new("保護範囲（行）")
                                        .color(egui::Color32::from_gray(180)),
                                );
                                ui.horizontal(|ui| {
                                    ui.add(
                                        egui::TextEdit::singleline(&mut self.prot_top)
                                            .desired_width(50.0)
                                            .hint_text("1"),
                                    );
                                    ui.label(
                                        egui::RichText::new("→")
                                            .color(egui::Color32::from_gray(140)),
                                    );
                                    ui.add(
                                        egui::TextEdit::singleline(&mut self.prot_bottom)
                                            .desired_width(50.0)
                                            .hint_text("3"),
                                    );
                                });
                                ui.end_row();
                            });
                    });

                ui.add_space(16.0);

                // ── 実行ボタン ──
                let can_run = self.file_path.is_some()
                    && self.file_path2.is_some()
                    && !self.search_string.is_empty()
                    && !self.search_tl.is_empty()
                    && !self.search_br.is_empty()
                    && !self.prot_top.is_empty()
                    && !self.prot_bottom.is_empty();

                ui.with_layout(egui::Layout::top_down(egui::Align::Center), |ui| {
                    let btn = egui::Button::new(
                        egui::RichText::new("  実  行  ")
                            .size(14.0)
                            .color(egui::Color32::WHITE),
                    )
                    .fill(egui::Color32::from_rgb(72, 110, 210))
                    .min_size(egui::vec2(160.0, 38.0));
                    if ui.add_enabled(can_run, btn).clicked() {
                        self.run();
                    }
                });

                // ── ステータス ──
                if !self.status.is_empty() {
                    ui.add_space(12.0);
                    let is_error = self.status.starts_with("エラー");
                    let (text_color, bg_color) = if is_error {
                        (
                            egui::Color32::from_rgb(243, 139, 168),
                            egui::Color32::from_rgb(55, 28, 35),
                        )
                    } else {
                        (
                            egui::Color32::from_rgb(166, 227, 161),
                            egui::Color32::from_rgb(28, 50, 35),
                        )
                    };
                    egui::Frame::new()
                        .fill(bg_color)
                        .corner_radius(egui::CornerRadius::same(6))
                        .inner_margin(egui::Margin::same(10))
                        .show(ui, |ui| {
                            ui.colored_label(text_color, &self.status);
                        });
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

        // 出力パスを自動生成: 元ファイルと同ディレクトリ、ファイル名末尾に _(検索文字列)
        let make_output_path = |input: &std::path::Path| -> PathBuf {
            let stem = input
                .file_stem()
                .unwrap_or_default()
                .to_string_lossy()
                .into_owned();
            let name = format!("{}_{}.xlsx", stem, self.search_string);
            input.parent().unwrap_or(std::path::Path::new(".")).join(name)
        };

        let input1 = self.file_path.as_deref().unwrap();
        let input2 = self.file_path2.as_deref().unwrap();
        let output1 = make_output_path(input1);
        let output2 = make_output_path(input2);

        let mut messages = Vec::new();
        for (input, output) in [(input1, &output1), (input2, &output2)] {
            let config = soph_core::ExtractionConfig {
                input_path: input.to_string_lossy().into_owned(),
                output_path: output.to_string_lossy().into_owned(),
                search_string: self.search_string.clone(),
                search_tl: tl,
                search_br: br,
                prot_top,
                prot_bottom,
            };
            match soph_core::extract_and_combine(&config) {
                Ok(count) => messages.push(format!(
                    "✓ {} → {} 行抽出",
                    input.file_name().unwrap_or_default().to_string_lossy(),
                    count,
                )),
                Err(e) => messages.push(format!(
                    "エラー [{}]: {e}",
                    input.file_name().unwrap_or_default().to_string_lossy(),
                )),
            }
        }
        self.status = messages.join("\n");
    }
}
