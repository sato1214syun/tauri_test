import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import { useState } from "react";
import "./App.css";

function App() {
	const [csv_path, setCSVPath] = useState("");
	const [excel_path, setExcelPath] = useState("");
	const [result_message, setResultMsg] = useState("");

	function open_dialog(file_type: string) {
		let filter: { name: string; extensions: string[] }[];
		if (file_type === "csv") {
			filter = [
				{
					name: "CSV file",
					extensions: ["csv"],
				},
			];
		} else if (file_type === "excel") {
			filter = [
				{
					name: "Excel file",
					extensions: ["xlsx"],
				},
			];
		} else {
			return;
		}
		open({ multiple: false, filters: filter, directory: false }).then(
			(files) => {
				if (files && files.length > 0) {
					if (file_type === "csv") {
						setCSVPath(files);
					} else if (file_type === "excel") {
						setExcelPath(files);
					}
				} else {
					if (file_type === "csv") {
						setCSVPath("No file selected");
					} else if (file_type === "excel") {
						setExcelPath("No file selected");
					}
				}
			},
		);
	}

	async function save_excel() {
		const save_path = await save({
			filters: [
				{
					name: "Excel file",
					extensions: ["xlsx"],
				},
			],
		});
		if (save_path === null) {
			setResultMsg("保存先が選択されていません");
			return;
		}
		invoke("write_excel", {
			csvPath: csv_path,
			excelPath: excel_path,
			savePath: save_path,
		})
			.then(() => {
				setResultMsg("更新された体調の管理エクセルを保存しました");
			})
			.catch((error) => {
				setResultMsg(`エラーが発生しました:\n${error}`);
			});
	}

	return (
		<main className="container">
			<h1>体調データ解析アプリ</h1>
			<div style={{ textAlign: "left" }}>
				<h2>使用手順</h2>
				<p>1. Rhythm Careからデータをcsvで保存する(エクスポート)</p>

				<div>
					<p>2. csvファイルをアップロードする</p>
					<button type="button" onClick={() => open_dialog("csv")}>
						Upload CSV
					</button>
					<p>Selected CSV file: {csv_path.split("/").pop()}</p>
				</div>

				<div>
					<p>3. 前回のエクセルファイルがあればアップロードする</p>
					<button type="button" onClick={() => open_dialog("excel")}>
						Upload Excel
					</button>
					<p>Selected Excel file: {excel_path.split("/").pop()}</p>
				</div>

				<div>
					<p>4. 解析結果(エクセル)を保存する</p>
					<button type="button" onClick={() => save_excel()}>
						Save Excel
					</button>
					<p>{result_message}</p>
				</div>
			</div>
		</main>
	);
}

export default App;
