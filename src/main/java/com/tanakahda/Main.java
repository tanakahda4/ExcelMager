/*
 * Copyright © 2024 tanakahda. All rights reserved.
 * https://github.com/tanakahda4/ExcelMerger
 */
package com.tanakahda;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory;

/**
 * 複数のExcelを１つのシートへマージするユーティリティクラスです。
 *
 * @author tanakahda
 *
 */
public class Main {

	/** ログ出力クラス */
	private static Logger _log = LoggerFactory.getLogger(Main.class);
	/** 設定ファイルのモデル */
	private static Config _config = null;

	/**
	 * メインルーティン
	 * @param args
	 */
	public static void main(String[] args) {

		_log.info("処理を開始します。");

		Workbook destWorkbook = new XSSFWorkbook();
		Sheet destSheet = destWorkbook.createSheet("Sheet1");
		_config = getConfig("src/main/resource/config.yaml");

		List<String> srcPaths = _config.getSrcExcelPath();

		srcPaths.forEach(System.out::println);

		// 順番に転記元のExcelファイルを開き、転記先のExcelへ書き込む
		for (String dir : srcPaths) {
			try {
				FileInputStream fis = new FileInputStream(dir);
				XSSFWorkbook srcWorkbook = new XSSFWorkbook(fis);
				Sheet srcSheet = srcWorkbook.getSheetAt(0);
				execSheet(destSheet, srcSheet);
				srcWorkbook.close();
			} catch (IOException e) {
				e.printStackTrace();
				_log.error(e.getMessage() + "：" + dir + "の転記に失敗したためスキップしました。");
			}
		}

		try (FileOutputStream out = new FileOutputStream(_config.getDestPath())) {
			destWorkbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
			_log.error(e.getMessage() + "：転記先のブックの出力に失敗しました。");
			_log.error("転記先のブック：" + _config.getDestPath());
		}

		try {
			destWorkbook.close();
		} catch (IOException e) {
			e.printStackTrace();
			_log.warn(e.getMessage());
		}

		_log.info("処理を終了しました。");
	}

	/**
	 * 設定ファイルのモデルクラスを取得します。
	 * @return 設定ファイル
	 */
	private static Config getConfig(String path) {
		String yaml;
		Config config = null;
		try {
			yaml = Files.readString(Paths.get(path));
			var mapper = new ObjectMapper(new YAMLFactory());
			config = mapper.readValue(yaml, Config.class);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return config;
	}

	/**
	 * シートの処理を実行します。
	 * @param destSheet 転記先のシート
	 * @param srcSheet 転記元のシート
	 */
	private static void execSheet(Sheet destSheet, Sheet srcSheet) {

		// 転記先ブックの最終行を取得して1行加算する（加算しないと最終行を上書きしてしまうため）
		final int destLastRowNum = destSheet.getLastRowNum() > 0 ? (destSheet.getLastRowNum() + 1) : 0;
		System.out.println(destLastRowNum);

		// 転記元ブックの最終行を取得して、行頭から行末まで繰り返す
		for (int i = 0; i <= srcSheet.getLastRowNum(); i++) {
			Row row = srcSheet.getRow(i);
			if (row != null) {
				Row newRow = destSheet.createRow(destLastRowNum + i);
				execRow(newRow, row);
			}
		}
	}

	/**
	 * 行の処理を実行します。
	 * @param destRow 転記先の行
	 * @param srcRow 転記元の行
	 */
	private static void execRow(Row destRow, Row srcRow) {
		for (int i = 0; i < srcRow.getLastCellNum(); i++) {
			Cell cell = srcRow.getCell(i);
			if (cell != null) {
				// A列、B列、C列などにコピーを絞り込むときは"^(A|B|C).*")ような文字列が設定されている
				String filterCopyColumn = _config.getCopyColumn();
				if (! filterCopyColumn.isEmpty()) {
					// セルのアドレス（A1など）を取得
					String addr = cell.getAddress().formatAsString();
					if (addr.matches(filterCopyColumn)) {
						Cell destCell = destRow.createCell(i);
						execCell(destCell, cell);
					}
				} else {
					Cell destCell = destRow.createCell(i);
					execCell(destCell, cell);
				}
			}
		}
	}

	/**
	 * セルの処理を実行（転記）します。
	 * @param destCell 転記先のセル
	 * @param srcCell 転記元のセル
	 */
	private static void execCell(Cell destCell, Cell srcCell) {
		switch (srcCell.getCellType()) {
			case STRING:
				destCell.setCellValue(srcCell.getStringCellValue());
				break;
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(srcCell)) {
					destCell.setCellValue(srcCell.getDateCellValue());
				} else {
					destCell.setCellValue(srcCell.getNumericCellValue());
				}
				break;
			case BOOLEAN:
				destCell.setCellValue(srcCell.getBooleanCellValue());
				break;
			case FORMULA:
				destCell.setCellFormula(srcCell.getCellFormula());
				break;
			case BLANK:
				destCell.setBlank();
				break;
			case ERROR:
				destCell.setCellErrorValue(srcCell.getErrorCellValue());
				break;
			default:
				break;
		}
	}
}
