use std::path::PathBuf;

use anyhow;
use calamine::{open_workbook, Reader, Xlsx};
use chrono::NaiveDate;
use polars::prelude::*;
use polars_excel_writer::PolarsXlsxWriter;
use rust_xlsxwriter::{
    chart::{
        Chart, ChartFont, ChartFormat, ChartLayout, ChartLine, ChartMarker, ChartMarkerType,
        ChartSolidFill,
    },
    conditional_format::ConditionalFormatDataBar,
    worksheet::Worksheet,
    Color, Workbook,
};
// Learn more about Tauri commands at https://tauri.app/develop/calling-rust/

struct YearlyData {
    year: i32,
    ldf: LazyFrame,
}

struct MonthlyData {
    month: i32,
    ldf: LazyFrame,
}

struct ConditionWorkbook {
    workbook: Workbook,
    writer: PolarsXlsxWriter,
}
impl ConditionWorkbook {
    fn new() -> Self {
        let workbook = Workbook::new();
        let writer = PolarsXlsxWriter::new();
        Self { workbook, writer }
    }

    fn write(&mut self, ldf: &LazyFrame, path: &str) -> anyhow::Result<()> {
        match self._write_raw_data(ldf, "data") {
            Ok(_) => self._write_yearly_data(ldf)?,
            Err(e) => return Err(e),
        }
        match self.workbook.save(path) {
            Ok(_) => Ok(()),
            Err(e) => Err(e.into()),
        }
    }

    fn _write_raw_data(&mut self, ldf: &LazyFrame, sheet_name: &str) -> anyhow::Result<()> {
        let worksheet = self.workbook.add_worksheet().set_name(sheet_name)?;

        self.writer
            .write_dataframe_to_worksheet(&ldf.clone().collect()?, worksheet, 0, 0)?;
        Ok(())
    }

    fn _write_yearly_data(&mut self, ldf: &LazyFrame) -> anyhow::Result<()> {
        // # 年間の体調集計データの比較シートを作成
        let comp_sheet_name = "年間体調比較";
        self.workbook.add_worksheet().set_name(comp_sheet_name)?;

        // 集計表用の書式を設定
        let annual_data_format = ConditionalFormatDataBar::new().set_fill_color(Color::Orange);
        let monthly_data_format = ConditionalFormatDataBar::new().set_fill_color(Color::Green);

        let mut row = 0;
        let col = 0;
        // Create a new Excel writer.
        for yearly_data in extract_yearly_frame_vec(ldf) {
            let sheet_name = yearly_data.year.to_string();
            let mut worksheet = Worksheet::new();
            worksheet.set_name(&sheet_name)?;
            // この年のシートにデータを書き込み
            let yearly_ldf = prepare_yearly_frame(&yearly_data.ldf, yearly_data.year);
            let yearly_df = match yearly_ldf.clone().collect() {
                Ok(df) => df,
                Err(e) => return Err(e.into()),
            };
            self.writer
                .write_dataframe_to_worksheet(&yearly_df, &mut worksheet, 0, 0)?;

            let yearly_agg_df = prepare_agg_frame(&yearly_ldf).collect()?;
            self.writer
                .write_dataframe_to_worksheet(&yearly_agg_df, &mut worksheet, 0, 6)?;

            // 集計表に条件付き書式を設定
            worksheet.add_conditional_format(1, 8, 6, 8, &annual_data_format.clone())?;
            worksheet.add_conditional_format(1, 9, 6, 20, &monthly_data_format.clone())?;

            // 年毎体調比較シートに集計データを書き込み
            let workbook_comp = self.workbook.worksheet_from_name(comp_sheet_name)?;
            workbook_comp.write_string(row, col, &sheet_name)?;
            row += 1;
            self.writer
                .write_dataframe_to_worksheet(&yearly_agg_df, workbook_comp, row, col)?;
            let end_row = row + yearly_agg_df.height() as u32;
            // 集計表に条件付き書式を設定
            workbook_comp.add_conditional_format(row + 1, 2, end_row, 2, &annual_data_format)?;
            workbook_comp.add_conditional_format(row + 1, 3, end_row, 14, &monthly_data_format)?;
            row = end_row + 1;

            // この年のシートに月毎の体調推移グラフを挿入
            self._insert_monthly_trend_chart(&mut worksheet, &yearly_ldf)?;
            // workbookにworksheetを追加
            self.workbook.push_worksheet(worksheet);
        }
        Ok(())
    }

    fn _insert_monthly_trend_chart(
        &self,
        worksheet: &mut Worksheet,
        yearly_ldf: &LazyFrame,
    ) -> anyhow::Result<()> {
        // 体調の推移グラフ挿入

        let insert_matrix = (6, 2); // 6行2列に並べる
        let insert_start_cell = (8, 5); // F8セルから挿入
        let per_chart_offset = (8, 11); // グラフの配置間隔がセルで何個分か
        let yearly_ldf_wt_idx = yearly_ldf.clone().with_row_index("cell_row", Some(1));

        for (i, monthly_data) in extract_monthly_frame_vec(&yearly_ldf_wt_idx)
            .iter()
            .enumerate()
        {
            let jp_month_str = format!("{}月", monthly_data.month);

            // 月毎の体調トレンドのグラフ作成
            let monthly_data_df = monthly_data.ldf.clone().collect()?;
            let cell_rows = monthly_data_df.column("cell_row")?.u32()?;

            let start_row = cell_rows.first().unwrap();
            let end_row = cell_rows.last().unwrap();
            let start_col = 0;

            let trend_line_chart =
                self._add_line_chart(&*worksheet.name(), start_row, start_col, end_row, 1);

            let mut base_chart =
                self._add_base_chart(&*worksheet.name(), start_row, start_col, end_row, 4);

            // グラフの結合
            base_chart.combine(&trend_line_chart);
            // 書式調整
            base_chart
                .title()
                .set_name(&jp_month_str)
                .set_font(&ChartFont::new().set_size(14));
            self._set_chart_format(&mut base_chart, monthly_data_df.height() as u32)?;

            // グラフを挿入
            let insert_row = insert_start_cell.0 + (i / insert_matrix.1 * per_chart_offset.0);
            let insert_col = insert_start_cell.1 + (i % insert_matrix.1 * per_chart_offset.1);
            worksheet.insert_chart(insert_row as u32, insert_col as u16, &base_chart)?;
        }
        Ok(())
    }

    fn _add_line_chart(
        &self,
        sheet_name: &str,
        start_row: u32,
        start_col: u16,
        end_row: u32,
        end_col: u16,
    ) -> Chart {
        let mut line_chart = Chart::new_line();

        line_chart
            .add_series()
            .set_name((sheet_name, 0, end_col))
            .set_categories((sheet_name, start_row, start_col, end_row, start_col))
            .set_values((sheet_name, start_row, end_col, end_row, end_col))
            .set_format(ChartFormat::new().set_line(ChartLine::new().set_color(Color::Blue)))
            .set_marker(ChartMarker::new().set_type(ChartMarkerType::Circle));
        line_chart
    }

    fn _add_base_chart(
        &self,
        sheet_name: &str,
        start_row: u32,
        start_col: u16,
        end_row: u32,
        end_col: u16,
    ) -> Chart {
        let mut col_chart = Chart::new_column();

        col_chart
            .add_series()
            .set_name((sheet_name, 0, end_col))
            .set_categories((sheet_name, start_row, start_col, end_row, start_col))
            .set_values((sheet_name, start_row, end_col, end_row, end_col))
            .set_format(
                ChartFormat::new()
                    .set_no_border()
                    .set_solid_fill(ChartSolidFill::new().set_color("#FBE5D6")),
            )
            .set_gap(10);
        col_chart
    }

    fn _set_chart_format(&self, chart: &mut Chart, date_cnt: u32) -> anyhow::Result<()> {
        chart.set_width(620);
        chart.set_height(155);

        chart.legend().set_hidden();
        chart
            .x_axis()
            .set_date_axis(true)
            .set_major_unit_date_type(rust_xlsxwriter::chart::ChartAxisDateUnitType::Days)
            .set_major_unit(1)
            .set_num_format("m/d")
            .set_major_gridlines(true)
            .set_major_gridlines_line(&ChartLine::new().set_color("#D0D0D0"))
            .set_position_between_ticks(false);
        chart
            .y_axis()
            .set_min(1)
            .set_max(5)
            .set_major_unit(1)
            .set_font(&ChartFont::new().set_size(11))
            .set_major_gridlines(true)
            .set_major_gridlines_line(&ChartLine::new().set_color("#D0D0D0"));
        chart.plot_area().set_layout(
            &ChartLayout::new()
                .set_offset(0.05, 0.20)
                .set_dimensions(0.9 * date_cnt as f64 / 31.0, 0.50),
        );
        Ok(())
    }
}

fn extract_yearly_frame_vec(ldf: &LazyFrame) -> Vec<YearlyData> {
    let df_with_year = ldf
        .clone()
        .with_column(col("日付").dt().year().cast(DataType::Int32).alias("year"));

    let binding = df_with_year
        .clone()
        .collect()
        .unwrap()
        .column("year")
        .unwrap()
        .unique()
        .unwrap();

    let years = binding.i32().unwrap();

    let mut yearly_data = vec![];
    for year in years {
        let yearly_ldf = df_with_year
            .clone()
            .filter(col("year").eq(lit(year.unwrap_or_default())))
            .drop([col("year")]);

        yearly_data.push(YearlyData {
            year: year.unwrap_or_default(),
            ldf: yearly_ldf,
        });
    }
    yearly_data
}

fn extract_monthly_frame_vec(ldf: &LazyFrame) -> Vec<MonthlyData> {
    let df_with_month = ldf.clone().with_column(
        col("日付")
            .dt()
            .month()
            .cast(DataType::Int32)
            .alias("month"),
    );

    let binding = df_with_month
        .clone()
        .collect()
        .unwrap()
        .column("month")
        .unwrap()
        .unique()
        .unwrap();

    let months = binding.i32().unwrap();

    let mut monthly_data = vec![];
    for month in months {
        let monthly_df = df_with_month
            .clone()
            .filter(col("month").eq(lit(month.unwrap_or_default())));

        monthly_data.push(MonthlyData {
            month: month.unwrap_or_default(),
            ldf: monthly_df,
        });
    }
    monthly_data
}

fn prepare_yearly_frame(ldf: &LazyFrame, year: i32) -> LazyFrame {
    // 1年分の日付列を準備
    let start = NaiveDate::from_ymd_opt(year, 1, 1).unwrap().into();
    let end = NaiveDate::from_ymd_opt(year, 12, 31).unwrap().into();
    let interval = Duration::parse("1d");
    let dates = date_range(
        "日付".into(),
        start,
        end,
        interval,
        ClosedWindow::Left,
        TimeUnit::Microseconds,
        None,
    )
    .unwrap();
    let date_col = dates.cast(&DataType::Date).unwrap();
    let mut yearly_ldf = DataFrame::new(vec![date_col.into()]).unwrap().lazy();

    yearly_ldf = yearly_ldf.left_join(ldf.clone(), col("日付"), col("日付"));
    // 曜日を追加
    let weekdays_int_series: Series = [1, 2, 3, 4, 5, 6, 7].iter().collect();
    let weekdays_str_series: Series = ["月", "火", "水", "木", "金", "土", "日"]
        .iter()
        .map(|&s| s)
        .collect();
    let holidays_int_series: Series = [0, 0, 0, 0, 0, 5, 5].iter().collect();
    yearly_ldf = yearly_ldf.with_columns(vec![
        col("日付")
            .dt()
            .weekday()
            .replace_strict(
                lit(weekdays_int_series.clone()),
                lit(weekdays_str_series),
                None,
                Some(DataType::String),
            )
            .alias("曜日"),
        col("日付")
            .dt()
            .weekday()
            .replace_strict(
                lit(weekdays_int_series),
                lit(holidays_int_series),
                None,
                Some(DataType::Int32),
            )
            .alias("土日判定"),
    ]);
    yearly_ldf
}

fn prepare_agg_frame(yearly_ldf: &LazyFrame) -> LazyFrame {
    // # 年間の体調の集計dfを作成
    let mut agg_ldf = df!(
        "調子" => ["↑", "↗", "→", "↘", "↓", "⇓"],
        "体調" => [5, 4, 3, 2, 1, 0]
    )
    .unwrap()
    .lazy();

    let yearly_agg_ldf = yearly_ldf
        .clone()
        .group_by([col("体調")])
        .agg([col("日付").count().alias("年間")])
        .select([col("体調"), col("年間")]);

    agg_ldf = agg_ldf.left_join(yearly_agg_ldf, col("体調"), col("体調"));

    // 月毎の集計を追加
    for monthly_data in extract_monthly_frame_vec(yearly_ldf) {
        let jp_month_str = format!("{}月", monthly_data.month);
        // 集計表に月毎の体調の集計を追加
        let temp_agg_ldf = monthly_data
            .ldf
            .group_by([col("体調")])
            .agg([col("体調").count().alias(&jp_month_str)])
            .select([col("体調"), col(&jp_month_str)]);
        agg_ldf = agg_ldf.left_join(temp_agg_ldf, col("体調"), col("体調"));
    }
    agg_ldf.fill_null(lit(0))
}

#[tauri::command]
fn write_excel(csv_path: &str, excel_path: &str, save_path: &str) -> Result<(), String> {
    let additional_condition_df = match read_csv(Some(csv_path.into())) {
        Ok(df) => df,
        Err(e) => return Err(e.to_string()),
    };
    let ori_condition_df = match read_excel(excel_path) {
        Ok(df) => df,
        Err(e) => return Err(e.to_string()),
    };
    let merged_ldf = merge_condition_data(&additional_condition_df, &ori_condition_df);
    let mut workbook = ConditionWorkbook::new();
    match workbook.write(&merged_ldf, save_path) {
        Ok(_) => Ok(()),
        Err(e) => Err(e.to_string()),
    }
}

fn read_csv(path: Option<PathBuf>) -> PolarsResult<DataFrame> {
    let schema = Schema::from_iter(vec![
        Field::new("日付".into(), DataType::Date),
        Field::new("体調".into(), DataType::Int32),
        Field::new("コメント".into(), DataType::String),
    ]);
    match CsvReadOptions::default()
        .with_has_header(false)
        .with_skip_rows(2)
        .with_schema(Some(Arc::new(schema)))
        .try_into_reader_with_file_path(path)
    {
        Ok(csv_reader) => csv_reader
            .finish()?
            .drop_nulls(Some(&vec!["日付".to_string()]))?
            .sort(
                ["日付"],
                SortMultipleOptions::new()
                    .with_order_descending(false)
                    .with_nulls_last(true),
            ),
        Err(e) => return Err(e.into()),
    }
}

fn read_excel(path: &str) -> anyhow::Result<DataFrame> {
    let mut excel: Xlsx<_> = open_workbook(path).unwrap();

    let range = match excel.worksheet_range("data") {
        Ok(range) => range,
        Err(e) => return Err(anyhow::anyhow!(e)),
    };

    let mut dates: Vec<Option<NaiveDate>> = vec![];
    let mut conditions: Vec<Option<i32>> = vec![];
    let mut comments: Vec<Option<String>> = vec![];
    for row in range.rows().skip(1) {
        let date = match calamine::DataType::as_date(&row[0]) {
            Some(value) => Some(value),
            None => None,
        };
        let condition = match calamine::DataType::as_i64(&row[1]) {
            Some(value) => Some(value as i32),
            None => None,
        };
        let comment = match calamine::DataType::as_string(&row[2]) {
            Some(value) => Some(value),
            None => None,
        };

        dates.push(date);
        conditions.push(condition);
        comments.push(comment);
    }
    df!(
        "日付" => dates,
        "体調" => conditions,
        "コメント" => comments,
    )
    .map_err(|e| anyhow::anyhow!(e))
}

fn merge_condition_data(csv_df: &DataFrame, condition_df: &DataFrame) -> LazyFrame {
    condition_df
        .vstack(csv_df)
        .unwrap()
        .lazy()
        .unique(Some(vec!["日付".to_string()]), UniqueKeepStrategy::Last)
        .sort(
            ["日付"],
            SortMultipleOptions::new()
                .with_order_descending(false)
                .with_nulls_last(true),
        )
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![write_excel])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::NaiveDate;
    use rust_xlsxwriter::{ExcelDateTime, Format};
    // use rust_xlsxwriter::*;
    use std::fs::File;
    use std::io::{BufWriter, Write};
    use tempfile::tempdir;

    #[test]
    fn test_read_csv() {
        // const NEW_LINE_CODE: &str = "\r\n";
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("test.csv");
        // let file_path = "../test_data/test.csv";

        // 書き込み専用で開く
        let f = File::create(&file_path).unwrap();
        let mut bfw = BufWriter::new(f);
        const BOM: &[u8; 3] = &[0xEF, 0xBB, 0xBF]; // UTF-8 BOM
        bfw.write_all(BOM).unwrap();

        bfw.write("日付,\"愛さん体調\",\"愛さん体調\"\n".as_bytes())
            .unwrap();
        bfw.write(",\"\",コメント\n".as_bytes()).unwrap();
        bfw.write("2025/01/27,,\n".as_bytes()).unwrap();
        bfw.write(",2,Should be dropped\n".as_bytes()).unwrap();
        bfw.write("2025/01/28,3,Test comment2\n".as_bytes())
            .unwrap();
        bfw.flush().unwrap();

        let expected_df = df!(
            "日付" => [Some(NaiveDate::from_ymd_opt(2025, 01, 27).unwrap()), Some(NaiveDate::from_ymd_opt(2025, 01, 28).unwrap())],
            "体調" => [None, Some(3i32)],
            "コメント" => [None, Some("Test comment2")]
        )
        .unwrap();

        let df = match read_csv(Some(file_path.into())) {
            Ok(data) => data,
            Err(e) => {
                println!("error: {}", e);
                return;
            }
        };
        temp_dir.close().unwrap();
        assert!(df.equals_missing(&expected_df));
    }

    #[test]
    fn test_read_excel() {
        let temp_dir = tempdir().unwrap();
        let file_path = temp_dir.path().join("test.xlsx");
        let mut workbook = Workbook::new();
        let date_format = Format::new().set_num_format("yyyy/mm/dd");
        // Add a worksheet to the workbook.
        let worksheet = workbook.add_worksheet().set_name("data").unwrap();
        // Write a header.
        worksheet.write_string(0, 0, "日付").unwrap();
        worksheet.write_string(0, 1, "体調").unwrap();
        worksheet.write_string(0, 2, "コメント").unwrap();
        // Write dates
        worksheet
            .write_with_format(
                1,
                0,
                ExcelDateTime::from_ymd(2025, 1, 25).unwrap(),
                &date_format,
            )
            .unwrap()
            // .write_with_format(
            //     2,
            //     0,
            //     ExcelDateTime::from_ymd(2025, 1, 26).unwrap(),
            //     &date_format,
            // )
            // .unwrap()
            .write_number(2, 1, 2)
            .unwrap()
            .write_string(2, 2, "Test comment")
            .unwrap();

        // Save the workbook
        workbook.save(&file_path).unwrap();
        assert!(file_path.exists());

        let df = read_excel(file_path.to_str().unwrap()).unwrap();
        temp_dir.close().unwrap();

        let expected_df = df!(
            "日付" => [Some(NaiveDate::from_ymd_opt(2025, 01, 25).unwrap()), None],
            "体調" => [None, Some(2i32)],
            "コメント" => [None, Some("Test comment")]
        )
        .unwrap();
        assert!(df.equals_missing(&expected_df));
    }

    #[test]
    fn test_merge_condition_df() {
        let csv_df = df!(
            "日付" => [NaiveDate::from_ymd_opt(2025, 01, 27).unwrap(), NaiveDate::from_ymd_opt(2025, 01, 28).unwrap()],
            "体調" => [None, Some(4i32)],
            "コメント" => [None, Some("Test comment")]
        )
        .unwrap();
        let excel_df = df!(
            "日付" => [NaiveDate::from_ymd_opt(2025, 01, 25).unwrap(), NaiveDate::from_ymd_opt(2025, 01, 26).unwrap()],
            "体調" => [None, Some(2i32)],
            "コメント" => [None, Some("テストコメント")]
        )
        .unwrap();

        let expected_df = df!(
            "日付" => [
                NaiveDate::from_ymd_opt(2025, 01, 25).unwrap(),
                NaiveDate::from_ymd_opt(2025, 01, 26).unwrap(),
                NaiveDate::from_ymd_opt(2025, 01, 27).unwrap(),
                NaiveDate::from_ymd_opt(2025, 01, 28).unwrap(),

            ],
            "体調" => [None, Some(2i32), None, Some(4i32)],
            "コメント" => [None, Some("テストコメント"), None, Some("Test comment")]
        )
        .unwrap();
        let ldf = merge_condition_data(&csv_df, &excel_df);
        print!("ldf: {:?}", ldf.clone().collect().unwrap());
        println!("expected_df: {:?}", expected_df);
        assert!(ldf.collect().unwrap().equals_missing(&expected_df));
    }

    #[test]
    fn test_write_excel() {
        let test_df = df!(
            "日付" => [
                NaiveDate::from_ymd_opt(2024, 01, 25).unwrap(),
                NaiveDate::from_ymd_opt(2024, 01, 26).unwrap(),
                NaiveDate::from_ymd_opt(2024, 01, 27).unwrap(),
                NaiveDate::from_ymd_opt(2024, 01, 28).unwrap(),
                NaiveDate::from_ymd_opt(2025, 01, 25).unwrap(),
                NaiveDate::from_ymd_opt(2025, 01, 26).unwrap(),
                NaiveDate::from_ymd_opt(2025, 01, 27).unwrap(),
                NaiveDate::from_ymd_opt(2025, 01, 28).unwrap(),
            ],
            "体調" => [None, Some(2i32), None, Some(3i32), None, Some(4i32), None, Some(5i32)],
            "コメント" => [None, Some("Test comment1"), None, Some("Test comment2"), None, Some("Test comment3"), None, Some("Test comment4")]
        )
        .unwrap();

        let mut wb = ConditionWorkbook::new();
        // パスの指定は、src-tauriディレクトリに対して相対パスで指定する必要がある
        // 例えば、root/src-tauri/src/lib.rsからroot/testのファイルを指定する場合は、../testのように指定する
        let path = "../test_data/test.xlsx";
        match wb.write(&test_df.lazy(), path) {
            Ok(_) => assert!(true),
            Err(e) => panic!("Failed to write Excel file: {}", e),
        }
    }
    #[test]
    fn test_write_excel_with_csv() {
        let csv_path = "../test_data/RhythmCareData.csv";
        let excel_path = "../test_data/体調記録_Sean_20250331.xlsx";
        let save_path = "../test_data/test.xlsx";
        match write_excel(csv_path, excel_path, save_path) {
            Ok(_) => assert!(true),
            Err(e) => panic!("Failed to write Excel file: {}", e),
        }
    }
}
