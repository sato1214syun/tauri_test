import { invoke } from "@tauri-apps/api/core";
// import { emit, listen } from "@tauri-apps/api/event";
// import { getCurrentWebview } from "@tauri-apps/api/webview";
import { open, save } from "@tauri-apps/plugin-dialog";
import { useState } from "react";
import "./App.css";

function App() {
	const [csv_path, setCSVPath] = useState("");
	const [excel_path, setExcelPath] = useState("");
	const [save_result, setSaveResult] = useState("");

	function open_dialog(file_type: string) {
		open({ multiple: false, directory: false }).then((files) => {
			if (files && files.length > 0) {
				if (file_type === "csv") {
					setCSVPath(`Selected CSV file: ${files}`);
				} else if (file_type === "excel") {
					setExcelPath(`Selected Excel file: ${files}`);
				}
			} else {
				if (file_type === "csv") {
					setCSVPath("No file selected");
				} else if (file_type === "excel") {
					setExcelPath("No file selected");
				}
			}
		});
	}

	async function save_excel() {
		const path = await save({
			filters: [
				{
					name: "My Filter",
					extensions: ["png", "jpeg"],
				},
			],
		});
		invoke("save_excel", { path }).then((response) => {
			setSaveResult(response);
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
					<p>{csv_path}</p>
				</div>

				<div>
					<p>3. 前回のエクセルファイルがあればアップロードする</p>
					<button type="button" onClick={() => open_dialog("excel")}>
						Upload Excel
					</button>
					<p>{excel_path}</p>
				</div>

				<div>
					<p>4. 解析結果(エクセル)を保存する</p>
					<button type="button" onClick={() => save_excel()}>
						Save Excel
					</button>
					<p>{save_result}</p>
				</div>
			</div>
		</main>
	);
}

export default App;
