use std::path::PathBuf;

use anyhow::Result;
use calamine::{open_workbook, Reader, Xlsx};
use chrono::NaiveDate;
use polars::prelude::*;
use polars_excel_writer::PolarsXlsxWriter;
use rust_xlsxwriter::{
    chart::{Chart, ChartType}, conditional_format::{ConditionalFormat, ConditionalFormatCell, ConditionalFormatCellRule, ConditionalFormatDataBar}, worksheet::Worksheet, Color, ExcelDateTime, Format, Workbook
};
use serde::{Deserialize, Serialize};
use tauri::{Emitter, Listener, Manager};
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

    fn write(&mut self, ldf: &LazyFrame, path: &str) -> Result<()> {
        self._write_raw_data(ldf, "data")?;
        self._write_yearly_data(ldf)?;
        self.writer.save(path)?;
        Ok(())
    }

    fn _write_raw_data(&mut self, ldf: &LazyFrame, sheet_name: &str) -> Result<()> {
        let worksheet = self.workbook.add_worksheet().set_name(sheet_name)?;
        self.writer
            .write_dataframe_to_worksheet(&ldf.clone().collect()?, worksheet, 0, 0)?;
        Ok(())
    }

    fn _write_yearly_data(&mut self, ldf: &LazyFrame) -> Result<()> {
        // # 年間の体調集計データの比較シートを作成
        let worksheet_comp = self.workbook.add_worksheet().set_name("年毎体調比較")?;
        let comparison_data_position = (0, 0);
        // Create a new Excel writer.
        for yearly_data in extract_yearly_df_vec(ldf) {
            let sheet_name = yearly_data.year.to_string();
            let worksheet = self.workbook.add_worksheet().set_name(&sheet_name)?;
            // yearly_data.dfを整形して書き込み
            let yearly_ldf = prepare_yearly_df(&yearly_data.ldf, yearly_data.year);
            self.writer
                .write_dataframe_to_worksheet(&yearly_ldf.clone().collect()?, worksheet, 0, 0)?;

            // # 年毎の体調の集計データを書き込み
            self._write_yearly_agg_data(&yearly_ldf, worksheet);
            // # 年毎の体調の推移のグラフを挿入

            let mut chart = Chart::new(ChartType::Line);
            chart.add_series();
            // .set_values((sheet_name, 0, 0, 1, 0));
            worksheet.insert_chart(0, 2, &chart)?;

            // # データを書き込み
            // self._write_frame(yearly_df, worksheet.name)
            // # この年の集計データを書込
            // self._write_yearly_agg_data(yearly_df, worksheet.name, position="G1")
            // # 年毎の体調の推移の表を書き込み
            // worksheet_comp.write(*comparison_data_position, yearly_date.year)
            // comparison_data_position = (
            //     comparison_data_position[0] + 1,
            //     comparison_data_position[1],
            // )
            // self._write_yearly_agg_data(
            //     yearly_df, worksheet_comp.name, position=comparison_data_position
            // )
            // comparison_data_position = (
            //     comparison_data_position[0] + 7,
            //     comparison_data_position[1],
            // )
            // # 月毎の体調の推移のグラフを挿入
            // self._insert_monthly_trend_chart(yearly_df, worksheet.name)
        }
        Ok(())
    }

    fn _write_yearly_agg_data(&mut self, yearly_ldf: &LazyFrame, worksheet: &mut Worksheet) -> Result<()> {
        let mut agg_ldf = self._prepare_agg_frame(yearly_ldf);

        // # 体調の集計表を書き込んでを条件付き書式(データバー)を追加
        let months = (1..=12)
            .map(|num| format!("{}月", num))
            .collect::<Vec<_>>();
        agg_ldf = agg_ldf.fill_null(0);
        self.writer.write_dataframe_to_worksheet(&agg_ldf.collect()?, worksheet, 0, 6);

        // Write a conditional format over a range.
        let annual_data_format = ConditionalFormatDataBar::new().set_fill_color(Color::Orange);
        let monthly_data_format = ConditionalFormatDataBar::new().set_fill_color(Color::Green);
        worksheet.add_conditional_format(1, 8, 6, 8, &annual_data_format)?;
        worksheet.add_conditional_format(1, 9, 6, 20, &monthly_data_format)?;
        Ok(())
    }

    fn _prepare_agg_frame(&self, yearly_ldf: &LazyFrame) -> LazyFrame {
        // # 年間の体調の集計dfを作成
        let mut agg_ldf = DataFrame::new(vec![
            Column::new("調子".into(), vec!["↑", "↗", "→", "↘", "↓", "⇓"]),
            Column::new("体調".into(), vec![5, 4, 3, 2, 1, 0]),
        ])
        .unwrap()
        .lazy();

        agg_ldf = agg_ldf.left_join(
            yearly_ldf
                .clone()
                .group_by([col("体調")])
                .agg([col("年間").count()])
                .cast(
                    {
                        col("年間");
                        {
                            let mut map = PlHashMap::new();
                            map.insert("年間", DataType::Int32);
                            map
                        }
                    },
                    true,
                ),
            col("体調"),
            col("体調"),
        );

        // 月毎の集計を追加
        for monthly_data in extract_monthly_df_vec(yearly_ldf) {
            let jp_month_str = format!("{}月", monthly_data.month);
            // 集計表に月毎の体調の集計を追加
            let temp_agg_ldf = monthly_data
                .ldf
                .group_by([col("体調")])
                .agg([col("体調").count().alias(&jp_month_str)])
                .cast(
                    {
                        col(&jp_month_str);
                        {
                            let mut map = PlHashMap::new();
                            map.insert(jp_month_str.as_str(), DataType::Int32);
                            map
                        }
                    },
                    true,
                );
            agg_ldf = agg_ldf.left_join(temp_agg_ldf, col("体調"), col("体調"));
        }
        agg_ldf
    }

    fn _insert_monthly_trend_chart(&self) -> Result<()> {
        // """体調の推移グラフ挿入."""
        // insert_matrix = (6, 2)  # 6行2列にグラフを並べる
        // insert_cell = (8, 5)  # F8セルから挿入
        // per_chart_offset = (8, 11)  # グラフの配置間隔がセルで何個分か
        // yearly_df_wt_idx = yearly_df.with_row_index("cell_row", 1)
        // for i, (monthly_date, monthly_df) in enumerate(
        //     self._iter_yearly_data(yearly_df_wt_idx, "1mo"),
        // ):
        //     jp_month_str = f"{monthly_date.month}月"

        //     # 月毎の体調の推移をグラフ化
        //     # データの参照範囲を取得
        //     start_row = monthly_df[0, "cell_row"]
        //     end_row = monthly_df[-1, "cell_row"]

        //     trend_line_chart = self._add_line_chart(
        //         sheet_name,
        //         (start_row, 0),
        //         (end_row, 0),
        //         (start_row, 1),
        //         (end_row, 1),
        //     )

        //     base_chart = self._add_base_chart(
        //         sheet_name,
        //         (start_row, 0),
        //         (end_row, 0),
        //         (start_row, 4),
        //         (end_row, 4),
        //     )

        //     # グラフ結合
        //     base_chart.combine(trend_line_chart)
        //     # 書式調整
        //     base_chart.set_title({"name": jp_month_str, "name_font": {"size": 14}})
        //     self._set_chart_format(base_chart, monthly_df.height)

        //     # グラフを挿入
        //     insert_row = insert_cell[0] + (i // insert_matrix[1] * per_chart_offset[0])
        //     insert_col = insert_cell[1] + (i % insert_matrix[1] * per_chart_offset[1])
        //     worksheet: Worksheet = self.get_worksheet_by_name(sheet_name)
        //     worksheet.insert_chart(
        //         insert_row,
        //         insert_col,
        //         base_chart,
        //         {"object_position": 3},
        //     )
        Ok(())
    }

    fn _add_line_chart(&self) -> Result<()> {
        // chart = self.add_chart({"type": "line"})
        // chart.add_series(
        //     {
        //         "name": [sheet_name, 1, 0],
        //         "categories": [sheet_name, *category_start, *category_end],
        //         "values": [sheet_name, *values_start, *values_end],
        //         "line": {"color": "blue"},
        //         "marker": {"type": "circle"},
        //     },
        // )
        // return chart
        Ok(())
    }

    fn _add_base_chart(&self) -> Result<()> {
        // chart = self.add_chart({"type": "column"})
        // chart.add_series(
        //     {
        //         "name": [sheet_name, 4, 0],
        //         "categories": [sheet_name, *category_start, *category_end],
        //         "values": [sheet_name, *values_start, *values_end],
        //         "border": {"none": True},
        //         "fill": {"color": "#FBE5D6"},
        //         "gap": 10,
        //     },
        // )
        // return chart
        Ok(())
    }

    fn _set_chart_format(&self) -> Result<()> {
        // chart.set_size({"width": 620, "height": 155})
        // chart.set_legend({"none": True})
        // chart.set_x_axis(
        //     {
        //         "date_axis": True,
        //         "major_unit_type": "days",
        //         "major_unit": 1,
        //         "num_format": "m/d",
        //         "major_gridlines": {
        //             "visible": True,
        //             "line": {"color": "#D0D0D0"},
        //         },
        //         "position_axis": "on_tick",
        //     },
        // )
        // chart.set_y_axis(
        //     {
        //         "min": 1,
        //         "max": 5,
        //         "major_unit": 1,
        //         "num_font": {"size": 11},
        //         "major_gridlines": {
        //             "visible": True,
        //             "line": {"color": "#D0D0D0"},
        //         },
        //     },
        // )
        // chart.set_plotarea(
        //     {
        //         "layout": {
        //             "x": 0.05,
        //             "y": 0.20,
        //             "width": 0.9 * date_cnt / 31,
        //             "height": 0.5,
        //         },
        //     },
        // )
        Ok(())
    }
}

fn extract_yearly_df_vec(ldf: &LazyFrame) -> Vec<YearlyData> {
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
        let yearly_df = df_with_year
            .clone()
            .filter(col("year").eq(lit(year.unwrap_or_default())));

        yearly_data.push(YearlyData {
            year: year.unwrap_or_default(),
            ldf: yearly_df,
        });
    }
    yearly_data
}

fn extract_monthly_df_vec(ldf: &LazyFrame) -> Vec<MonthlyData> {
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

fn prepare_yearly_df(ldf: &LazyFrame, year: i32) -> LazyFrame {
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
    yearly_ldf.with_columns(vec![
        col("日付").dt().weekday().alias("曜日"),
        col("日付").dt().weekday().alias("土日判定"),
    ])

    // # 曜日と土日の判定列を追加
    // weekday_mapping = dict(zip(range(1, 8), "月火水木金土日", strict=False))
    // holiday_mapping = {x: 5 if x in [6, 7] else 0 for x in range(1, 8)}
    // return yearly_df.with_columns(
    //     pl.col("日付")
    //     .dt.weekday()
    //     .replace_strict(weekday_mapping, return_dtype=pl.String)
    //     .alias("曜日"),
    //     pl.col("日付").dt.weekday().replace(holiday_mapping).alias("土日判定"),
    // )
}

#[tauri::command]
fn write_excel(csv_path_str: &str, ori_excel_path_str: &str, save_path: &str) -> String {
    let additional_condition_df = match read_csv(Some(csv_path_str.into())) {
        Ok(df) => df,
        Err(e) => return e.to_string(),
    };
    let ori_condition_df = match read_excel(ori_excel_path_str) {
        Ok(df) => df,
        Err(e) => return e.to_string(),
    };
    let merged_ldf = merge_condition_data(&additional_condition_df, &ori_condition_df);
    let mut workbook = ConditionWorkbook::new();
    workbook.write(&merged_ldf, save_path).unwrap();
    "Excelファイルが保存されました".to_string()
}

fn read_csv(path: Option<PathBuf>) -> PolarsResult<DataFrame> {
    let schema = Schema::from_iter(vec![
        Field::new("日付".into(), DataType::Date),
        Field::new("体調".into(), DataType::Int64),
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

fn read_excel(path: &str) -> Result<DataFrame> {
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
        let condition = match calamine::DataType::as_i32(&row[1]) {
            Some(value) => Some(value),
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
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![write_excel])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::NaiveDate;
    use rust_xlsxwriter::*;
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
            "コメント" => [None, Some("Test comment")]
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
            "コメント" => [None, Some("Test comment"), None, Some("Test comment")]
        )
        .unwrap();
        let ldf = merge_condition_data(&csv_df, &excel_df);
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
        wb.write(&test_df.lazy(), "./test_data/test.xlsx").unwrap();
        assert!(std::path::Path::new("./test_data/test.xlsx").exists());
    }
}
