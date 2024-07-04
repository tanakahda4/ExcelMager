package com.tanakahda;

import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

public class Config {

	@JsonProperty("dest_path")
	private String destPath;

	@JsonProperty("copy_column")
	private String copyColumn;

	@JsonProperty("src_excel_path")
	private List<String> srcExcelPath;

	public String getDestPath() {
		return destPath;
	}

	public void setDestPath(String destPath) {
		this.destPath = destPath;
	}

	public String getCopyColumn() {
		return copyColumn;
	}

	public void setCopyColumn(String copyColumn) {
		this.copyColumn = copyColumn;
	}

	public List<String> getSrcExcelPath() {
		return srcExcelPath;
	}

	public void setSrcExcelPath(List<String> srcExcelPath) {
		this.srcExcelPath = srcExcelPath;
	}

}
